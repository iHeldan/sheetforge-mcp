import json

import pytest
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
import excel_mcp.server as server_module

from excel_mcp.data import (
    append_table_rows,
    quick_read as quick_read_impl,
    read_as_table,
    read_excel_range,
    read_excel_range_with_metadata,
    search_cells,
    update_rows_by_key,
    write_data,
)
from excel_mcp.exceptions import DataError
from excel_mcp.server import (
    list_all_sheets,
    quick_read,
    read_data_from_excel,
    read_excel_as_table as read_excel_as_table_tool,
    search_in_sheet,
)


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    assert "operation" in payload
    assert "message" in payload
    return payload


def _create_chartsheet_first_workbook(tmp_path) -> str:
    filepath = str(tmp_path / "chartsheet-first.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("Alice", 30), ("Bob", 25)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    categories = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb._sheets = [chart_sheet, ws]
    wb.active = 0
    wb.save(filepath)
    wb.close()
    return filepath


def test_read_from_explicit_start_cell(tmp_workbook):
    """Bug #1: read_excel_range should respect explicit start_cell when no end_cell given."""
    result = read_excel_range(tmp_workbook, "Sheet1", start_cell="B2")
    flat = [cell for row in result for cell in row]
    assert "Name" not in flat, "start_cell was ignored — returned data from A1"
    assert result[0][0] == 30


def test_read_from_a1_default_uses_full_range(tmp_workbook):
    """Default A1 start should read the full data range."""
    result = read_excel_range(tmp_workbook, "Sheet1")
    assert result[0][0] == "Name"
    assert len(result) == 6


def test_read_with_metadata_respects_start_cell(tmp_workbook):
    """Bug #1 also affects read_excel_range_with_metadata."""
    result = read_excel_range_with_metadata(tmp_workbook, "Sheet1", start_cell="C4")
    cells = result["cells"]
    values = [c["value"] for c in cells]
    assert "Name" not in values, "start_cell was ignored in metadata read"
    assert cells[0]["value"] == "Turku"


def test_read_as_table_returns_headers_and_rows(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1")
    assert result["headers"] == ["Name", "Age", "City"]
    assert len(result["rows"]) == 5
    assert result["rows"][0] == ["Alice", 30, "Helsinki"]
    assert result["rows"][4] == ["Eve", 32, "Espoo"]


def test_read_as_table_with_max_rows(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", max_rows=2)
    assert len(result["rows"]) == 2
    assert result["total_rows"] == 5
    assert result["truncated"] is True
    assert result["next_start_row"] == 4


def test_read_as_table_supports_start_row_pagination(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", start_row=4, max_rows=2)

    assert result["headers"] == ["Name", "Age", "City"]
    assert result["rows"] == [
        ["Carol", 35, "Turku"],
        ["Dave", 28, "Oulu"],
    ]
    assert result["total_rows"] == 5
    assert result["truncated"] is True
    assert result["next_start_row"] == 6


def test_read_as_table_can_omit_headers_for_followup_pages(tmp_workbook):
    result = read_as_table(
        tmp_workbook,
        "Sheet1",
        start_row=4,
        max_rows=2,
        include_headers=False,
    )

    assert "headers" not in result
    assert result["rows"] == [
        ["Carol", 35, "Turku"],
        ["Dave", 28, "Oulu"],
    ]
    assert result["truncated"] is True
    assert result["next_start_row"] == 6


def test_read_as_table_omits_next_start_row_on_final_page(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", start_row=6, max_rows=2)

    assert result["rows"] == [["Eve", 32, "Espoo"]]
    assert result["truncated"] is False
    assert "next_start_row" not in result


def test_read_as_table_rejects_start_row_at_or_above_header(tmp_workbook):
    with pytest.raises(DataError, match="start_row must be greater than header_row"):
        read_as_table(tmp_workbook, "Sheet1", header_row=1, start_row=1)


def test_read_as_table_rejects_non_positive_max_rows(tmp_workbook):
    with pytest.raises(DataError, match="max_rows must be a positive integer"):
        read_as_table(tmp_workbook, "Sheet1", max_rows=0)


def test_read_as_table_ignores_trailing_rows_outside_selected_columns(tmp_path):
    filepath = tmp_path / "selected-columns.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in [
        ("Name", "Age", "City"),
        ("Alice", 30, "Helsinki"),
        ("Bob", 25, "Tampere"),
    ]:
        ws.append(row)
    ws["D20"] = "Ignored outside selected columns"
    wb.save(filepath)
    wb.close()

    result = read_as_table(str(filepath), "Sheet1", start_col="A", end_col="C")

    assert result["rows"] == [
        ["Alice", 30, "Helsinki"],
        ["Bob", 25, "Tampere"],
    ]
    assert result["total_rows"] == 2
    assert result["truncated"] is False


def test_read_as_table_custom_header_row(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", header_row=2)
    assert result["headers"] == ["Alice", 30, "Helsinki"]


def test_read_as_table_can_return_records_and_schema(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", row_mode="objects", infer_schema=True)

    assert result["records"][0] == {
        "name": "Alice",
        "age": 30,
        "city": "Helsinki",
    }
    assert result["schema"] == [
        {"field": "name", "header": "Name", "type": "string", "nullable": False},
        {"field": "age", "header": "Age", "type": "integer", "nullable": False},
        {"field": "city", "header": "City", "type": "string", "nullable": False},
    ]
    assert result["row_mode"] == "objects"
    assert "rows" not in result


def test_read_as_table_schema_normalizes_and_dedupes_headers(tmp_path):
    from openpyxl import Workbook

    filepath = tmp_path / "schema.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "First Name"
    ws["B1"] = None
    ws["C1"] = "First Name"
    ws["A2"] = "Alice"
    ws["B2"] = 30
    ws["C2"] = "Admin"
    wb.save(filepath)
    wb.close()

    result = read_as_table(str(filepath), "Sheet1", row_mode="objects", infer_schema=True)

    assert result["schema"] == [
        {"field": "first_name", "header": "First Name", "type": "string", "nullable": False},
        {"field": "column_2", "header": None, "type": "integer", "nullable": False},
        {"field": "first_name_2", "header": "First Name", "type": "string", "nullable": False},
    ]
    assert result["records"][0] == {
        "first_name": "Alice",
        "column_2": 30,
        "first_name_2": "Admin",
    }


def test_read_as_table_schema_transliterates_accented_headers(tmp_path):
    filepath = tmp_path / "schema-fi.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Näyttökerrat"
    ws["B1"] = "Lisäklikit"
    ws["C1"] = "CTR %"
    ws["A2"] = 1200
    ws["B2"] = 45
    ws["C2"] = 3.75
    wb.save(filepath)
    wb.close()

    result = read_as_table(str(filepath), "Sheet1", row_mode="objects", infer_schema=True)

    assert result["schema"] == [
        {"field": "nayttokerrat", "header": "Näyttökerrat", "type": "integer", "nullable": False},
        {"field": "lisaklikit", "header": "Lisäklikit", "type": "integer", "nullable": False},
        {"field": "ctr", "header": "CTR %", "type": "number", "nullable": False},
    ]
    assert result["records"][0] == {
        "nayttokerrat": 1200,
        "lisaklikit": 45,
        "ctr": 3.75,
    }


def test_read_as_table_rejects_chart_sheet_with_clear_error(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    with pytest.raises(DataError, match="Sheet 'Charts' is a chartsheet"):
        read_as_table(filepath, "Charts")


def test_read_excel_range_rejects_chart_sheet_with_clear_error(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    with pytest.raises(DataError, match="Sheet 'Charts' is a chartsheet"):
        read_excel_range(filepath, "Charts", start_cell="A1", end_cell="B2")

def test_search_cells_finds_exact_match(tmp_workbook):
    results = search_cells(tmp_workbook, "Sheet1", "Alice")
    assert len(results) == 1
    assert results[0]["cell"] == "A2"
    assert results[0]["value"] == "Alice"


def test_search_cells_finds_partial_match(tmp_workbook):
    results = search_cells(tmp_workbook, "Sheet1", "li", exact=False)
    values = [r["value"] for r in results]
    assert "Alice" in values


def test_search_cells_no_match(tmp_workbook):
    results = search_cells(tmp_workbook, "Sheet1", "Nonexistent")
    assert results == []


def test_search_cells_finds_numbers(tmp_workbook):
    results = search_cells(tmp_workbook, "Sheet1", 30)
    assert len(results) == 1
    assert results[0]["cell"] == "B2"


def test_search_cells_matches_numeric_values_from_string_queries(tmp_workbook):
    results = search_cells(tmp_workbook, "Sheet1", "30")
    assert len(results) == 1
    assert results[0]["cell"] == "B2"


def test_search_in_sheet_accepts_numeric_queries(tmp_workbook):
    payload = _load_tool_payload(search_in_sheet(tmp_workbook, "Sheet1", 30))
    assert payload["operation"] == "search_in_sheet"
    assert len(payload["data"]["matches"]) == 1
    assert payload["data"]["matches"][0]["cell"] == "B2"


def test_read_data_from_excel_preview_only_limits_output(tmp_path):
    filepath = tmp_path / "preview.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Value"
    for row in range(2, 17):
        ws[f"A{row}"] = f"Row {row}"
    wb.save(filepath)
    wb.close()

    payload = _load_tool_payload(read_data_from_excel(str(filepath), "Sheet1", preview_only=True))
    preview_rows = {cell["row"] for cell in payload["data"]["cells"]}

    assert len(preview_rows) == 10
    assert payload["data"]["preview_only"] is True
    assert payload["data"]["truncated"] is True


def test_read_data_from_excel_compact_omits_default_validation(tmp_workbook):
    payload = _load_tool_payload(read_data_from_excel(tmp_workbook, "Sheet1", compact=True))
    first_cell = payload["data"]["cells"][0]

    assert first_cell["address"] == "A1"
    assert "validation" not in first_cell


def test_read_data_from_excel_compact_keeps_real_validation(tmp_path):
    from openpyxl.worksheet.datavalidation import DataValidation

    filepath = tmp_path / "validated.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Status"
    ws["A2"] = "Open"
    validation = DataValidation(type="list", formula1='"Open,Closed"')
    validation.add("A2")
    ws.add_data_validation(validation)
    wb.save(filepath)
    wb.close()

    payload = _load_tool_payload(read_data_from_excel(str(filepath), "Sheet1", compact=True))
    cells_by_address = {cell["address"]: cell for cell in payload["data"]["cells"]}

    assert "validation" in cells_by_address["A2"]
    assert cells_by_address["A2"]["validation"]["has_validation"] is True
    assert cells_by_address["A2"]["validation"]["allowed_values"] == ["Open", "Closed"]


def test_read_excel_range_with_metadata_values_only_returns_2d_values(tmp_workbook):
    result = read_excel_range_with_metadata(
        tmp_workbook,
        "Sheet1",
        start_cell="A1",
        end_cell="B3",
        values_only=True,
    )

    assert result == {
        "range": "A1:B3",
        "sheet_name": "Sheet1",
        "values": [
            ["Name", "Age"],
            ["Alice", 30],
            ["Bob", 25],
        ],
    }


def test_read_data_from_excel_values_only_preview_limits_rows(tmp_path):
    filepath = tmp_path / "preview-values.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Value"
    for row in range(2, 17):
        ws[f"A{row}"] = f"Row {row}"
    wb.save(filepath)
    wb.close()

    payload = _load_tool_payload(
        read_data_from_excel(str(filepath), "Sheet1", preview_only=True, values_only=True)
    )

    assert "cells" not in payload["data"]
    assert len(payload["data"]["values"]) == 10
    assert payload["data"]["values"][0] == ["Value"]
    assert payload["data"]["preview_only"] is True
    assert payload["data"]["truncated"] is True


def test_read_data_from_excel_values_only_handles_out_of_bounds_start(tmp_workbook):
    payload = _load_tool_payload(
        read_data_from_excel(
            tmp_workbook,
            "Sheet1",
            start_cell="Z100",
            values_only=True,
        )
    )

    assert payload["data"]["values"] == []
    assert "cells" not in payload["data"]


def test_read_excel_as_table_compact_omits_nonessential_metadata(tmp_workbook):
    payload = _load_tool_payload(read_excel_as_table_tool(tmp_workbook, "Sheet1", compact=True))

    assert payload["operation"] == "read_excel_as_table"
    assert payload["data"] == {
        "headers": ["Name", "Age", "City"],
        "rows": [
            ["Alice", 30, "Helsinki"],
            ["Bob", 25, "Tampere"],
            ["Carol", 35, "Turku"],
            ["Dave", 28, "Oulu"],
            ["Eve", 32, "Espoo"],
        ],
    }


def test_read_excel_as_table_compact_preserves_truncation_metadata(tmp_workbook):
    payload = _load_tool_payload(read_excel_as_table_tool(tmp_workbook, "Sheet1", max_rows=2, compact=True))

    assert payload["data"]["headers"] == ["Name", "Age", "City"]
    assert payload["data"]["rows"] == [["Alice", 30, "Helsinki"], ["Bob", 25, "Tampere"]]
    assert payload["data"]["total_rows"] == 5
    assert payload["data"]["truncated"] is True


def test_read_excel_as_table_tool_can_return_records_and_schema(tmp_workbook):
    payload = _load_tool_payload(
        read_excel_as_table_tool(
            tmp_workbook,
            "Sheet1",
            compact=True,
            row_mode="objects",
            infer_schema=True,
        )
    )

    assert payload["data"]["headers"] == ["Name", "Age", "City"]
    assert payload["data"]["records"][0] == {
        "name": "Alice",
        "age": 30,
        "city": "Helsinki",
    }
    assert payload["data"]["schema"][1]["type"] == "integer"
    assert payload["data"]["row_mode"] == "objects"
    assert "rows" not in payload["data"]


def test_read_excel_as_table_tool_rejects_chart_sheet_with_clear_error(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    payload = json.loads(read_excel_as_table_tool(filepath, "Charts"))

    assert payload["ok"] is False
    assert payload["error"]["type"] == "DataError"
    assert "Sheet 'Charts' is a chartsheet" in payload["error"]["message"]


def test_read_data_from_excel_tool_rejects_chart_sheet_with_clear_error(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    payload = json.loads(read_data_from_excel(filepath, "Charts", start_cell="A1", end_cell="B2"))

    assert payload["ok"] is False
    assert payload["error"]["type"] == "DataError"
    assert "Sheet 'Charts' is a chartsheet" in payload["error"]["message"]


def test_list_all_sheets_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(list_all_sheets(tmp_workbook))
    assert payload["operation"] == "list_all_sheets"
    assert payload["data"]["sheets"][0]["name"] == "Sheet1"


def test_quick_read_uses_first_sheet_when_sheet_name_omitted(multi_sheet_workbook):
    result = quick_read_impl(multi_sheet_workbook)
    assert result["sheet_name"] == "Sales"
    assert result["auto_selected_sheet"] is True
    assert result["headers"] == ["Product", "Revenue"]


def test_quick_read_respects_explicit_sheet_name(multi_sheet_workbook):
    result = quick_read_impl(multi_sheet_workbook, sheet_name="Inventory")
    assert result["sheet_name"] == "Inventory"
    assert result["auto_selected_sheet"] is False
    assert result["headers"] == ["Item", "Count"]


def test_quick_read_can_return_records_and_schema(multi_sheet_workbook):
    result = quick_read_impl(
        multi_sheet_workbook,
        sheet_name="Inventory",
        row_mode="objects",
        infer_schema=True,
    )

    assert result["records"] == [{"item": "Widget", "count": 42}]
    assert result["schema"] == [
        {"field": "item", "header": "Item", "type": "string", "nullable": False},
        {"field": "count", "header": "Count", "type": "integer", "nullable": False},
    ]
    assert result["row_mode"] == "objects"
    assert "rows" not in result


def test_quick_read_tool_returns_json_envelope(multi_sheet_workbook):
    payload = _load_tool_payload(quick_read(multi_sheet_workbook))
    assert payload["operation"] == "quick_read"
    assert payload["data"]["sheet_name"] == "Sales"
    assert payload["data"]["auto_selected_sheet"] is True


def test_quick_read_tool_supports_start_row_pagination(tmp_workbook):
    payload = _load_tool_payload(
        quick_read(tmp_workbook, sheet_name="Sheet1", start_row=4, max_rows=2)
    )

    assert payload["data"]["rows"] == [
        ["Carol", 35, "Turku"],
        ["Dave", 28, "Oulu"],
    ]
    assert payload["data"]["truncated"] is True
    assert payload["data"]["next_start_row"] == 6


def test_quick_read_tool_can_omit_headers_for_followup_pages(tmp_workbook):
    payload = _load_tool_payload(
        quick_read(
            tmp_workbook,
            sheet_name="Sheet1",
            start_row=4,
            max_rows=2,
            include_headers=False,
        )
    )

    assert "headers" not in payload["data"]
    assert payload["data"]["rows"] == [
        ["Carol", 35, "Turku"],
        ["Dave", 28, "Oulu"],
    ]


def test_quick_read_skips_chartsheet_when_first_sheet_is_not_a_worksheet(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    result = quick_read_impl(filepath)

    assert result["sheet_name"] == "Data"
    assert result["auto_selected_sheet"] is True
    assert result["headers"] == ["Name", "Value"]


def test_quick_read_tool_skips_chartsheet_when_first_sheet_is_not_a_worksheet(tmp_path):
    filepath = _create_chartsheet_first_workbook(tmp_path)

    payload = _load_tool_payload(quick_read(filepath))

    assert payload["operation"] == "quick_read"
    assert payload["data"]["sheet_name"] == "Data"
    assert payload["data"]["auto_selected_sheet"] is True


def test_quick_read_rejects_invalid_row_mode(multi_sheet_workbook):
    payload = json.loads(quick_read(multi_sheet_workbook, row_mode="dicts"))
    assert payload["ok"] is False
    assert payload["error"]["message"] == "row_mode must be 'arrays' or 'objects'"


def test_tool_responses_are_compact_json(multi_sheet_workbook):
    raw = quick_read(multi_sheet_workbook)

    assert "\n" not in raw
    assert raw.startswith('{"ok":true')


def test_read_data_from_excel_returns_guided_error_before_oversized_payload(
    tmp_workbook,
    monkeypatch,
):
    monkeypatch.setattr(server_module, "MCP_RESPONSE_CHAR_LIMIT", 200)

    payload = json.loads(read_data_from_excel(tmp_workbook, "Sheet1"))

    assert payload["ok"] is False
    assert payload["error"]["type"] == "ResponseTooLargeError"
    assert payload["error"]["estimated_size"] > payload["error"]["limit"]
    assert any("values_only=True" in hint for hint in payload["error"]["hints"])
    assert any("preview_only=True" in hint for hint in payload["error"]["hints"])
    assert any("start_cell/end_cell" in hint for hint in payload["error"]["hints"])


def test_quick_read_returns_guided_error_before_oversized_payload(
    tmp_workbook,
    monkeypatch,
):
    monkeypatch.setattr(server_module, "MCP_RESPONSE_CHAR_LIMIT", 120)

    payload = json.loads(quick_read(tmp_workbook, sheet_name="Sheet1"))

    assert payload["ok"] is False
    assert payload["error"]["type"] == "ResponseTooLargeError"
    assert any("max_rows" in hint for hint in payload["error"]["hints"])
    assert any("start_row" in hint for hint in payload["error"]["hints"])


def test_write_data_dry_run_does_not_persist(tmp_workbook):
    result = write_data(tmp_workbook, "Sheet1", [["Mallory", 44, "Lahti"]], start_cell="A2", dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["cell"] == "A2"

    table = read_as_table(tmp_workbook, "Sheet1")
    assert table["rows"][0] == ["Alice", 30, "Helsinki"]


def test_write_data_defaults_to_summary_without_changes(tmp_workbook):
    result = write_data(tmp_workbook, "Sheet1", [["Mallory", 44, "Lahti"]], start_cell="A2")

    assert result["dry_run"] is False
    assert result["changed_cells"] == 3
    assert "changes" not in result


def test_write_data_can_include_changes_explicitly(tmp_workbook):
    result = write_data(
        tmp_workbook,
        "Sheet1",
        [["Mallory", 44, "Lahti"]],
        start_cell="A2",
        include_changes=True,
    )

    assert result["changed_cells"] == 3
    assert result["changes"][0]["cell"] == "A2"


def test_append_table_rows_appends_using_headers(tmp_workbook):
    result = append_table_rows(
        tmp_workbook,
        "Sheet1",
        [{"Name": "Mallory", "Age": 44, "City": "Lahti"}],
    )

    assert result["rows_appended"] == 1
    assert result["dry_run"] is False

    table = read_as_table(tmp_workbook, "Sheet1")
    assert table["rows"][-1] == ["Mallory", 44, "Lahti"]


def test_append_table_rows_defaults_to_summary_without_changes(tmp_workbook):
    result = append_table_rows(
        tmp_workbook,
        "Sheet1",
        [{"Name": "Mallory", "Age": 44, "City": "Lahti"}],
    )

    assert result["changed_cells"] == 3
    assert "changes" not in result


def test_append_table_rows_dry_run_does_not_persist(tmp_workbook):
    result = append_table_rows(
        tmp_workbook,
        "Sheet1",
        [{"Name": "Mallory", "Age": 44, "City": "Lahti"}],
        dry_run=True,
    )

    assert result["dry_run"] is True
    assert result["start_row"] == 7

    table = read_as_table(tmp_workbook, "Sheet1")
    assert table["total_rows"] == 5


def test_update_rows_by_key_updates_matching_rows_and_reports_missing_keys(tmp_workbook):
    result = update_rows_by_key(
        tmp_workbook,
        "Sheet1",
        "Name",
        [
            {"Name": "Alice", "City": "Vantaa", "Age": 31},
            {"Name": "Missing", "City": "Nowhere"},
        ],
    )

    assert result["updated_rows"] == 1
    assert result["missing_keys"] == ["Missing"]

    table = read_as_table(tmp_workbook, "Sheet1")
    assert table["rows"][0] == ["Alice", 31, "Vantaa"]


def test_update_rows_by_key_defaults_to_summary_without_changes(tmp_workbook):
    result = update_rows_by_key(
        tmp_workbook,
        "Sheet1",
        "Name",
        [{"Name": "Alice", "City": "Vantaa"}],
    )

    assert result["updated_rows"] == 1
    assert result["changed_cells"] == 1
    assert "changes" not in result


def test_update_rows_by_key_dry_run_does_not_persist(tmp_workbook):
    result = update_rows_by_key(
        tmp_workbook,
        "Sheet1",
        "Name",
        [{"Name": "Alice", "City": "Vantaa"}],
        dry_run=True,
    )

    assert result["dry_run"] is True
    assert result["changes"][0]["cell"] == "C2"

    table = read_as_table(tmp_workbook, "Sheet1")
    assert table["rows"][0] == ["Alice", 30, "Helsinki"]
