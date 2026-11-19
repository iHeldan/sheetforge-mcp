"""Shared save guards must compose across public mutation families."""

from datetime import datetime
import json
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.table import Table

from excel_mcp import server, workbook
from excel_mcp.exceptions import PreconditionFailedError
from tests.test_package_fidelity import _rewrite


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def make_book(path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    for row in [("Id", "Value"), (1, 10), (2, 20)]:
        sheet.append(row)
    native = book.copy_worksheet(sheet)
    native.title = "Native"
    native.add_table(Table(displayName="Records", ref="A1:B3"))
    book.create_sheet("Formulas")["A1"] = "=6*9"
    book.save(path)
    book.close()
    return path


def protect(path, kind):
    def inject(parts):
        if kind == "custom":
            parts["customXml/item1.xml"] = b'<synthetic xmlns="urn:sheetforge:test">retained</synthetic>'
            rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
            ET.SubElement(rels, f"{{{REL}}}Relationship", Id="rId99", Type=DOC_REL + "customXml", Target="../customXml/item1.xml")
            parts["xl/_rels/workbook.xml.rels"] = ET.tostring(rels)
        else:
            root = ET.fromstring(parts["xl/worksheets/sheet3.xml"])
            root.find(f".//{{{MAIN}}}c[@r='A1']/{{{MAIN}}}v").text = "54"
            parts["xl/worksheets/sheet3.xml"] = ET.tostring(root)
    _rewrite(path, inject)


PREVIEW_OPERATIONS = ["format", "freeze", "copy_range", "name", "sheet_append", "sheet_update", "native_append", "native_upsert", "merge", "insert"]
OPERATIONS = [*PREVIEW_OPERATIONS, "formula", "chart", "copy_sheet", "rename", "create", "delete", "create_table"]


def mutate(path, operation, dry_run=False):
    args = {"dry_run": dry_run}
    path = str(path)
    calls = {
        "format": lambda: server.format_range(path, "Data", "D1", bold=True, **args),
        "freeze": lambda: server.freeze_panes(path, "Data", "B2", **args),
        "copy_range": lambda: server.copy_range(path, "Data", "A2", "B2", "D2", **args),
        "name": lambda: server.create_named_range(path, "Window", "A2:B3", sheet_name="Data", **args),
        "sheet_append": lambda: server.append_table_rows(path, "Data", [{"Id": 3, "Value": 30}], **args),
        "sheet_update": lambda: server.update_rows_by_key(path, "Data", "Id", [{"Id": 1, "Value": 30}], **args),
        "native_append": lambda: server.append_excel_table_rows(path, "Records", [{"Id": 3, "Value": 30}], **args),
        "native_upsert": lambda: server.upsert_excel_table_rows(path, "Records", "Id", [{"Id": 1, "Value": 30}], **args),
        "merge": lambda: server.merge_cells(path, "Data", "D1", "E1", **args),
        "insert": lambda: server.insert_rows(path, "Data", 2, **args),
        "formula": lambda: server.apply_formula(path, "Data", "D1", "=1+1"),
        "chart": lambda: server.create_chart(path, "Data", "bar", "D5", data_range="A1:B3"),
        "copy_sheet": lambda: server.copy_worksheet(path, "Data", "Copy"),
        "rename": lambda: server.rename_worksheet(path, "Data", "Renamed"),
        "create": lambda: server.create_worksheet(path, "Extra"),
        "delete": lambda: server.delete_worksheet(path, "Data"),
        "create_table": lambda: server.create_table(path, "Data", "A1:B3", "MoreRecords"),
    }
    return json.loads(calls[operation]())


@pytest.mark.parametrize("kind", ["custom", "cache"])
@pytest.mark.parametrize("operation", OPERATIONS)
def test_all_mutation_families_refuse_protected_loss_before_save(tmp_path, kind, operation):
    path = make_book(tmp_path / "protected.xlsx")
    protect(path, kind)
    original, stamp = path.read_bytes(), path.stat().st_mtime_ns
    payload = mutate(path, operation)
    assert not payload["ok"], payload
    assert payload["error"]["code"] == "write_package_fidelity_failed", payload
    assert path.read_bytes() == original and path.stat().st_mtime_ns == stamp
    assert not list(tmp_path.glob(".*.sheetforge-*"))


@pytest.mark.parametrize("operation", PREVIEW_OPERATIONS)
def test_previews_check_the_same_package_guard_without_writing(tmp_path, operation):
    path = make_book(tmp_path / "preview.xlsx")
    protect(path, "cache")
    original = path.read_bytes()
    payload = mutate(path, operation, dry_run=True)
    assert not payload["ok"] and payload["error"]["code"] == "write_package_fidelity_failed", payload
    assert path.read_bytes() == original


@pytest.mark.parametrize("operation", OPERATIONS)
def test_editable_saves_keep_unrelated_rich_text(tmp_path, operation):
    path = make_book(tmp_path / "rich.xlsx")
    book = load_workbook(path)
    book["Formulas"]["C2"] = CellRichText("plain ", TextBlock(InlineFont(b=True), "bold"))
    book.save(path)
    book.close()
    assert mutate(path, operation)["ok"]
    book = load_workbook(path, rich_text=True)
    try:
        value = book["Formulas"]["C2"].value
        assert isinstance(value, CellRichText) and value[1].font.b
    finally:
        book.close()


def test_explicit_formula_replacement_can_invalidate_its_own_cache(tmp_path):
    path = make_book(tmp_path / "replace-formula.xlsx")
    protect(path, "cache")
    payload = json.loads(server.apply_formula(str(path), "Formulas", "A1", "=2+2"))
    assert payload["ok"], payload
    book = load_workbook(path)
    assert book["Formulas"]["A1"].value == "=2+2"
    book.close()


@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("kind", ["write_columns", "write_rows", "copy_columns", "copy_rows", "insert_columns", "insert_rows"])
def test_complete_mutation_extents_are_checked_before_writing(tmp_path, kind, dry_run):
    path = make_book(tmp_path / "bounds.xlsx")
    if kind.startswith("insert"):
        book = load_workbook(path)
        book["Data"]["XFD1" if kind.endswith("columns") else "A1048576"] = "edge"
        book.save(path)
        book.close()
    before = path.read_bytes()
    args = {"dry_run": dry_run}
    calls = {
        "write_columns": lambda: server.write_data_to_excel(str(path), "Data", [[1, 2]], "XFD2", **args),
        "write_rows": lambda: server.write_data_to_excel(str(path), "Data", [[1], [2]], "A1048576", **args),
        "copy_columns": lambda: server.copy_range(str(path), "Data", "A1", "B1", "XFD2", **args),
        "copy_rows": lambda: server.copy_range(str(path), "Data", "A1", "A2", "A1048576", **args),
        "insert_columns": lambda: server.insert_columns(str(path), "Data", 16384, **args),
        "insert_rows": lambda: server.insert_rows(str(path), "Data", 1048576, **args),
    }
    assert not json.loads(calls[kind]())["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("target", ["A-2", "A1B2", "A0", "XFE1", "A1:B2"])
def test_copy_rejects_malformed_destinations_instead_of_guessing(tmp_path, target):
    path = make_book(tmp_path / "target.xlsx")
    before = path.read_bytes()
    assert not json.loads(server.copy_range(str(path), "Data", "A1", "A1", target))["ok"]
    assert path.read_bytes() == before


def test_serialized_grid_is_checked_even_if_library_can_open_it(tmp_path):
    path = make_book(tmp_path / "invalid.xlsx")
    def corrupt(parts):
        parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(b'r="B3"', b'r="XFE3"')
    _rewrite(path, corrupt)
    with pytest.raises(ValueError, match="Excel limits"):
        workbook._verify_saved_workbook(str(path))


def test_serialized_grid_parser_rejects_dtds(tmp_path):
    path = make_book(tmp_path / "doctype.xlsx")
    def add_doctype(parts):
        parts["xl/worksheets/sheet1.xml"] = b'<!DOCTYPE worksheet []>' + parts["xl/worksheets/sheet1.xml"]
    _rewrite(path, add_doctype)
    with pytest.raises(workbook.WorkbookError, match="DTD"):
        workbook._verify_saved_workbook(str(path))


def test_scatter_creation_reads_rich_text_header_as_plain_title(tmp_path):
    path = make_book(tmp_path / "rich-header.xlsx")
    book = load_workbook(path)
    book["Data"]["B1"] = CellRichText("Sales ", TextBlock(InlineFont(b=True), "EUR"))
    book.save(path)
    book.close()
    payload = json.loads(server.create_chart(str(path), "Data", "scatter", "D5", data_range="A1:B3"))
    assert payload["ok"], payload
    book = load_workbook(path, rich_text=True)
    try:
        assert book["Data"]._charts[0].series[0].title.v == "Sales EUR"
        assert isinstance(book["Data"]["B1"].value, CellRichText)
    finally:
        book.close()


@pytest.mark.parametrize("formula", ['="a\rb"', '="' + "x" * 8190 + '"'], ids=["CR", "too-long"])
def test_formula_application_rejects_loss_before_saving(tmp_path, formula):
    path = make_book(tmp_path / "formula.xlsx")
    before = path.read_bytes()
    assert not json.loads(server.apply_formula(str(path), "Data", "B2", formula))["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("target_sheet", ["Data", "Native"])
def test_copy_resets_default_style_and_checks_saved_values(tmp_path, target_sheet):
    path = make_book(tmp_path / "date.xlsx")
    book = load_workbook(path)
    book[target_sheet]["C2"] = datetime(2020, 1, 1)
    book.save(path)
    book.close()
    assert json.loads(server.copy_range(str(path), "Data", "B2", "B2", "C2", target_sheet=target_sheet))["ok"]
    book = load_workbook(path)
    assert book[target_sheet]["C2"].value == 10 and book[target_sheet]["C2"].number_format == "General"
    book.close()


def test_overlapping_copy_snapshots_values_and_default_styles(tmp_path):
    path = make_book(tmp_path / "overlap.xlsx")
    book = load_workbook(path)
    book["Data"]["B3"] = datetime(2020, 1, 1)
    book.save(path)
    book.close()
    payload = json.loads(server.copy_range(str(path), "Data", "B2", "B3", "B3"))
    assert payload["ok"], payload
    book = load_workbook(path)
    try:
        assert book["Data"]["B3"].value == 10 and book["Data"]["B3"].number_format == "General"
        assert book["Data"]["B4"].value == datetime(2020, 1, 1)
        assert book["Data"]["B4"].number_format != "General"
    finally:
        book.close()


@pytest.mark.parametrize("before,after", [(1, True), (True, 1), (0, False), (False, 0)])
def test_copy_and_changesets_report_typed_changes(tmp_path, before, after):
    path = make_book(tmp_path / "typed.xlsx")
    book = load_workbook(path)
    book["Data"]["B2"], book["Data"]["C2"] = before, after
    book.save(path)
    book.close()
    preview = json.loads(server.copy_range(str(path), "Data", "C2", "C2", "B2", dry_run=True))
    assert len(preview["changes"]) == 1
    operations = [{"tool": "write_data_to_excel", "args": {"sheet_name": "Data", "start_cell": "B2", "data": [[after]]}}]
    plan = json.loads(server.apply_workbook_changeset(str(path), operations))["data"]
    assert plan["diff"]["cell_changes"]["count"] == 1
    saved = json.loads(server.apply_workbook_changeset(str(path), operations, mode="commit", expected_workbook_sha256=plan["expected_workbook_sha256"], changeset_token=plan["changeset_token"]))
    assert saved["ok"] and saved["data"]["diff"]["cell_changes"]["count"] == 1
    book = load_workbook(path)
    assert type(book["Data"]["B2"].value) is type(after)
    book.close()


def test_callbacks_cannot_bypass_shared_preservation(tmp_path):
    path = make_book(tmp_path / "callback.xlsx")
    protect(path, "custom")
    before = path.read_bytes()
    calls = []
    with pytest.raises(PreconditionFailedError, match="protected XLSX"):
        with workbook.safe_workbook(str(path), save=True, candidate_check=lambda *_: calls.append(True)) as book:
            book["Data"]["B2"] = 99
    assert calls == [True] and path.read_bytes() == before
