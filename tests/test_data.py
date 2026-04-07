import json

from excel_mcp.data import read_excel_range, read_excel_range_with_metadata
from excel_mcp.server import read_data_from_excel, search_in_sheet


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


from excel_mcp.data import read_as_table


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


def test_read_as_table_custom_header_row(tmp_workbook):
    result = read_as_table(tmp_workbook, "Sheet1", header_row=2)
    assert result["headers"] == ["Alice", 30, "Helsinki"]


from excel_mcp.data import search_cells


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
    results = json.loads(search_in_sheet(tmp_workbook, "Sheet1", 30))
    assert len(results) == 1
    assert results[0]["cell"] == "B2"


def test_read_data_from_excel_preview_only_limits_output(tmp_path):
    from openpyxl import Workbook

    filepath = tmp_path / "preview.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Value"
    for row in range(2, 17):
        ws[f"A{row}"] = f"Row {row}"
    wb.save(filepath)
    wb.close()

    payload = json.loads(read_data_from_excel(str(filepath), "Sheet1", preview_only=True))
    preview_rows = {cell["row"] for cell in payload["cells"]}

    assert len(preview_rows) == 10
    assert payload["preview_only"] is True
    assert payload["truncated"] is True
