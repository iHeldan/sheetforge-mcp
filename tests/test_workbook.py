import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from excel_mcp.chart import create_chart_in_sheet
from excel_mcp.server import (
    analyze_range_impact as analyze_range_impact_tool,
    get_workbook_metadata as get_workbook_metadata_tool,
    list_all_sheets as list_all_sheets_tool,
    profile_workbook as profile_workbook_tool,
)
from excel_mcp.tables import create_excel_table
from excel_mcp.workbook import (
    analyze_range_impact,
    get_or_create_workbook,
    get_workbook_info,
    list_sheets,
    profile_workbook,
)


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    assert "operation" in payload
    assert "message" in payload
    return payload


def test_get_or_create_raises_on_missing_file(tmp_path):
    """get_or_create_workbook should raise when file doesn't exist."""
    missing = str(tmp_path / "nonexistent.xlsx")
    with pytest.raises(FileNotFoundError):
        get_or_create_workbook(missing)


def test_get_or_create_loads_existing_file(tmp_workbook):
    """get_or_create_workbook should load existing files normally."""
    wb = get_or_create_workbook(tmp_workbook)
    assert "Sheet1" in wb.sheetnames
    wb.close()

def test_list_sheets_returns_names(multi_sheet_workbook):
    result = list_sheets(multi_sheet_workbook)
    assert len(result) == 2
    assert result[0]["name"] == "Sales"
    assert result[1]["name"] == "Inventory"
    assert result[0]["rows"] >= 2
    assert result[0]["columns"] >= 2
    assert result[0]["is_empty"] is False


def test_list_sheets_marks_empty_sheet(empty_workbook):
    result = list_sheets(empty_workbook)
    assert result == [
        {
            "name": "Sheet",
            "sheet_type": "worksheet",
            "rows": 0,
            "columns": 0,
            "column_range": None,
            "is_empty": True,
        }
    ]


def test_list_sheets_handles_chart_sheets(tmp_path):
    filepath = str(tmp_path / "chartsheet-list.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("A", 1), ("B", 2)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    categories = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb.save(filepath)
    wb.close()

    result = list_sheets(filepath)

    assert result == [
        {
            "name": "Data",
            "sheet_type": "worksheet",
            "rows": 3,
            "columns": 2,
            "column_range": "A-B",
            "is_empty": False,
        },
        {
            "name": "Charts",
            "sheet_type": "chartsheet",
            "rows": 0,
            "columns": 0,
            "column_range": None,
            "is_empty": False,
        },
    ]


def test_list_all_sheets_tool_handles_chart_sheets(tmp_path):
    filepath = str(tmp_path / "chartsheet-tool.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("A", 1), ("B", 2)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    categories = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb.save(filepath)
    wb.close()

    payload = json.loads(list_all_sheets_tool(filepath))

    assert payload["operation"] == "list_all_sheets"
    assert payload["data"]["sheets"][1]["sheet_type"] == "chartsheet"


def test_profile_workbook_summarizes_tables_and_charts(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="Customers")
    create_chart_in_sheet(
        filepath=tmp_workbook,
        sheet_name="Sheet1",
        chart_type="bar",
        target_cell="E1",
        data_range="A1:B6",
        title="Customers by Age",
    )

    result = profile_workbook(tmp_workbook)

    assert result["sheet_count"] == 1
    assert result["table_count"] == 1
    assert result["chart_count"] == 1
    assert result["named_range_count"] == 0

    sheet = result["sheets"][0]
    assert sheet["name"] == "Sheet1"
    assert sheet["used_range"] == "A1:C6"
    assert sheet["table_count"] == 1
    assert sheet["chart_count"] == 1
    assert sheet["tables"][0]["table_name"] == "Customers"
    assert sheet["charts"][0]["chart_type"] == "bar"
    assert sheet["charts"][0]["anchor"] == "E1"
    assert sheet["charts"][0]["occupied_range"].startswith("E1:")


def test_profile_workbook_tool_returns_json_envelope(tmp_workbook):
    payload = json.loads(profile_workbook_tool(tmp_workbook))

    assert payload["operation"] == "profile_workbook"
    assert payload["data"]["sheet_count"] == 1


def test_analyze_range_impact_reports_overlapping_structures(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="Customers")
    create_chart_in_sheet(
        filepath=tmp_workbook,
        sheet_name="Sheet1",
        chart_type="bar",
        target_cell="E1",
        data_range="A1:B6",
        title="Customers by Age",
    )

    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws.merge_cells("B2:C2")
    ws.auto_filter.ref = "A1:C6"
    ws.print_area = "A1:F10"
    ws["D3"] = "=SUM(B2:C2)"
    ws["H2"] = "=SUM(B2:C3)"
    dependent_sheet = workbook.create_sheet("Dependent")
    dependent_sheet["A1"] = "=SUM(Sheet1!B2:C3)"
    workbook.defined_names["ImpactArea"] = DefinedName(
        "ImpactArea",
        attr_text="Sheet1!$B$2:$F$4",
    )
    workbook.save(tmp_workbook)
    workbook.close()

    result = analyze_range_impact(tmp_workbook, "Sheet1", "A1:F4")

    assert result["summary"]["risk_level"] == "high"
    assert result["summary"]["table_count"] == 1
    assert result["summary"]["chart_count"] == 1
    assert result["summary"]["merged_range_count"] == 1
    assert result["summary"]["named_range_count"] == 1
    assert result["summary"]["formula_cell_count"] == 1
    assert result["summary"]["dependent_formula_count"] == 2
    assert result["summary"]["autofilter_overlap"] is True
    assert result["summary"]["print_area_overlap"] is True
    assert result["tables"][0]["covers_header"] is True
    assert result["charts"][0]["anchor"] == "E1"
    assert result["merged_ranges"][0]["range"] == "B2:C2"
    assert result["named_ranges"][0]["name"] == "ImpactArea"
    assert result["formula_cells"]["sample"] == ["D3"]
    assert result["dependent_formulas"]["count"] == 2
    assert result["dependent_formulas"]["sample"][0]["references"][0]["intersection_range"] == "B2:C3"
    dependent_cells = {
        (item["sheet_name"], item["cell"]) for item in result["dependent_formulas"]["sample"]
    }
    assert dependent_cells == {("Sheet1", "H2"), ("Dependent", "A1")}


def test_analyze_range_impact_reports_low_risk_for_empty_area(tmp_workbook):
    result = analyze_range_impact(tmp_workbook, "Sheet1", "H20:I21")

    assert result["summary"]["risk_level"] == "low"
    assert result["summary"]["table_count"] == 0
    assert result["summary"]["chart_count"] == 0
    assert result["summary"]["dependent_formula_count"] == 0
    assert result["hints"] == ["No overlapping workbook structures detected for this range."]


def test_analyze_range_impact_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(analyze_range_impact_tool(tmp_workbook, "Sheet1", "A1:C3"))

    assert payload["operation"] == "analyze_range_impact"
    assert payload["data"]["range"] == "A1:C3"


def test_profile_workbook_handles_chart_sheets(tmp_path):
    filepath = str(tmp_path / "chartsheet.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("A", 1), ("B", 2), ("C", 3)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=4)
    categories = Reference(ws, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb.save(filepath)
    wb.close()

    result = profile_workbook(filepath)

    assert result["sheet_count"] == 2
    assert result["chart_count"] == 1
    assert result["table_count"] == 0
    assert result["sheets"][0]["sheet_type"] == "worksheet"
    chart_sheet = result["sheets"][1]
    assert chart_sheet["name"] == "Charts"
    assert chart_sheet["sheet_type"] == "chartsheet"
    assert chart_sheet["visibility"] == "visible"
    assert chart_sheet["table_count"] == 0
    assert chart_sheet["chart_count"] == 1
    assert chart_sheet["tables"] == []
    assert chart_sheet["charts"][0]["chart_index"] == 1
    assert chart_sheet["charts"][0]["chart_type"] == "bar"
    assert chart_sheet["charts"][0]["series_count"] == 1
    assert "occupied_range" not in chart_sheet["charts"][0]


def test_get_workbook_info_include_ranges_skips_chart_sheets(tmp_path):
    filepath = str(tmp_path / "chartsheet-ranges.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("A", 1), ("B", 2)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    categories = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb.save(filepath)
    wb.close()

    result = get_workbook_info(filepath, include_ranges=True)

    assert result["sheets"] == ["Data", "Charts"]
    assert result["used_ranges"] == {"Data": "A1:B3"}


def test_get_workbook_metadata_tool_handles_chart_sheets_with_ranges(tmp_path):
    filepath = str(tmp_path / "chartsheet-metadata.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [("Name", "Value"), ("A", 1), ("B", 2)]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    categories = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart_sheet = wb.create_chartsheet("Charts")
    chart_sheet.add_chart(chart)
    wb.save(filepath)
    wb.close()

    payload = json.loads(get_workbook_metadata_tool(filepath, include_ranges=True))

    assert payload["ok"] is True
    assert payload["operation"] == "get_workbook_metadata"
    assert payload["data"]["sheets"] == ["Data", "Charts"]
    assert payload["data"]["used_ranges"] == {"Data": "A1:B3"}
