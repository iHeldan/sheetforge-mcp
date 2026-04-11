import pytest
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from excel_mcp.chart import create_chart_in_sheet
from excel_mcp.server import profile_workbook as profile_workbook_tool
from excel_mcp.tables import create_excel_table
from excel_mcp.workbook import get_or_create_workbook, list_sheets, profile_workbook


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
            "rows": 0,
            "columns": 0,
            "column_range": None,
            "is_empty": True,
        }
    ]


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


def test_profile_workbook_tool_returns_json_envelope(tmp_workbook):
    payload = profile_workbook_tool(tmp_workbook)

    assert '"operation": "profile_workbook"' in payload
    assert '"sheet_count": 1' in payload


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
