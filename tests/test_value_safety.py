from datetime import date, datetime, time, timezone
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

from excel_mcp.changeset import apply_workbook_changeset
from excel_mcp.data import append_table_rows, quick_read, update_rows_by_key, write_data
from excel_mcp.exceptions import DataError
from excel_mcp.tables import append_excel_table_rows, read_excel_table, upsert_excel_table_rows


WRITERS = ["cells", "sheet_append", "sheet_update", "native_append", "native_upsert"]


def make_book(tmp_path, kind, value="KEEP", key=1):
    path = tmp_path / "values.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active.append(["Id", "Value"])
    book.active.append([key, value])
    if kind.startswith("native"):
        book.active.add_table(Table(displayName="Records", ref="A1:B2"))
    book.save(path)
    book.close()
    return path


def write(path, kind, value, *, key=1, dry_run=False):
    args = dict(dry_run=dry_run, include_changes=not (isinstance(value, str) and len(value) > 1000))
    if kind == "cells":
        return write_data(str(path), "Data", [[value]], "B2", **args)
    if kind == "sheet_append":
        return append_table_rows(str(path), "Data", [{"Id": key, "Value": value}], **args)
    if kind == "sheet_update":
        return update_rows_by_key(str(path), "Data", "Id", [{"Id": key, "Value": value}], **args)
    if kind == "native_append":
        return append_excel_table_rows(str(path), "Records", [{"Id": key, "Value": value}], **args)
    return upsert_excel_table_rows(str(path), "Records", "Id", [{"Id": key, "Value": value}], **args)


@pytest.mark.parametrize("kind", WRITERS)
@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("value", [
    9007199254740993, float("nan"), float("inf"), float("-inf"),
    "x" * 32768, "\U0001f642" * 16384, "", "a\rb", "a\x00b", "\ud800",
    Decimal("0.1234567890123456789"), datetime(2026, 1, 1, microsecond=1),
    0.1 + 0.2, datetime(2026, 1, 1, tzinfo=timezone.utc),
], ids=lambda value: f"{type(value).__name__}-{len(value)}" if isinstance(value, str) else repr(value))
def test_writers_reject_lossy_values_without_modifying_source(tmp_path, kind, dry_run, value):
    path = make_book(tmp_path, kind)
    before = path.read_bytes()
    with pytest.raises(DataError):
        write(path, kind, value, dry_run=dry_run)
    assert path.read_bytes() == before


@pytest.mark.parametrize("kind", WRITERS)
@pytest.mark.parametrize("value", [None, True, False, 1.0, 0.125, Decimal("1.25"), "x" * 32767, "\U0001f642" * 16383, "a\nb"], ids=lambda value: f"{type(value).__name__}-{len(value)}" if isinstance(value, str) else repr(value))
def test_supported_values_round_trip(tmp_path, kind, value):
    path = make_book(tmp_path, kind)
    result = write(path, kind, value)
    book = load_workbook(path)
    actual = book["Data"]["B3" if kind.endswith("append") else "B2"].value
    book.close()
    assert actual == value
    if isinstance(value, bool):
        assert type(actual) is bool
    assert result["changed_cells"] >= 1


@pytest.mark.parametrize("kind", ["sheet_update", "native_upsert"])
def test_boolean_key_does_not_select_numeric_identity(tmp_path, kind):
    path = make_book(tmp_path, kind)
    result = write(path, kind, "boolean row", key=True)
    book = load_workbook(path)
    assert book["Data"]["B2"].value == "KEEP"
    if kind == "native_upsert":
        assert book["Data"]["A3"].value is True
        assert result["appended_rows"] == 1
    else:
        assert result["updated_rows"] == 0
        assert result["missing_keys"] == [True]
    book.close()


@pytest.mark.parametrize("kind", ["sheet_update", "native_upsert"])
@pytest.mark.parametrize("before,after", [(1, True), (True, 1), (0, False), (False, 0)])
def test_boolean_type_changes_are_previewed_and_written(tmp_path, kind, before, after):
    path = make_book(tmp_path, kind, before)
    original = path.read_bytes()
    preview = write(path, kind, after, dry_run=True)
    assert preview["changed_cells"] == 1
    assert path.read_bytes() == original
    result = write(path, kind, after)
    assert result["changed_cells"] == 1
    book = load_workbook(path)
    assert type(book["Data"]["B2"].value) is type(after)
    book.close()


@pytest.mark.parametrize("kind", WRITERS[1:])
def test_returned_tokens_describe_serialized_candidate(tmp_path, kind):
    path = make_book(tmp_path, kind)
    original = path.read_bytes()
    preview = write(path, kind, 1.0, dry_run=True)
    assert path.read_bytes() == original
    result = write(path, kind, 1.0)
    observed = read_excel_table(str(path), "Records") if kind.startswith("native") else quick_read(str(path), "Data")
    assert result["new_content_token"] == observed["content_token"] == preview["new_content_token"]
    assert result["new_structure_token"] == observed["structure_token"]


@pytest.mark.parametrize("value", [9007199254740993, "x" * 32768])
def test_changeset_rejects_loss_without_optional_value_assertions(tmp_path, value):
    path = make_book(tmp_path, "cells")
    before = path.read_bytes()
    with pytest.raises(Exception, match="round-trip|32767"):
        apply_workbook_changeset(str(path), [{"tool": "write_data_to_excel", "args": {
            "sheet_name": "Data", "start_cell": "B2", "data": [[value]],
        }}], create_snapshot=False)
    assert path.read_bytes() == before


def test_candidate_value_check_blocks_number_reinterpreted_by_existing_date_style(tmp_path):
    path = make_book(tmp_path, "cells", datetime(2026, 1, 1))
    before = path.read_bytes()
    with pytest.raises(DataError, match="Serialized value"):
        write(path, "cells", 1)
    assert path.read_bytes() == before


@pytest.mark.parametrize("kind", ["sheet_append", "sheet_update", "native_append", "native_upsert"])
def test_row_guards_keep_rich_text_read_identity_and_unchanged_runs(tmp_path, kind):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    path = make_book(tmp_path, kind)
    book = load_workbook(path)
    book["Data"]["B2"] = CellRichText(TextBlock(InlineFont(b=True), "KEEP"))
    book["Data"].append([2, "other"])
    if kind.startswith("native"):
        book["Data"].tables["Records"].ref = "A1:B3"
    else:
        book["Data"]["B1"] = CellRichText(TextBlock(InlineFont(b=True), "Value"))
    book.save(path)
    book.close()
    read = read_excel_table(str(path), "Records") if kind.startswith("native") else quick_read(str(path), "Data")
    result = write(path, kind, "KEEP")
    assert result["previous_content_token"] == read["content_token"]
    assert result["previous_structure_token"] == read["structure_token"]
    book = load_workbook(path, rich_text=True)
    assert isinstance(book["Data"]["B2"].value, CellRichText)
    book.close()


@pytest.mark.parametrize("kind", ["cells", "sheet_update", "native_upsert"])
def test_rich_text_change_details_match_plain_read_values(tmp_path, kind):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    path = make_book(tmp_path, kind)
    book = load_workbook(path)
    book["Data"]["B2"] = CellRichText(TextBlock(InlineFont(b=True), "KEEP"))
    book.save(path)
    book.close()
    result = write(path, kind, "new")
    assert result["changes"][0]["old_value"] == "KEEP"
    assert result["changes"][0]["new_value"] == "new"


@pytest.mark.parametrize("kind", WRITERS)
@pytest.mark.parametrize("value", [date(2026, 1, 1), datetime(2026, 1, 1, 12, 30, microsecond=123000), time(12, 30, microsecond=123000)])
def test_documented_naive_date_time_policy_round_trips(tmp_path, kind, value):
    path = make_book(tmp_path, kind)
    write(path, kind, value)
    book = load_workbook(path)
    actual = book["Data"]["B3" if kind.endswith("append") else "B2"].value
    book.close()
    expected = datetime.combine(value, time()) if type(value) is date else value
    assert actual == expected
