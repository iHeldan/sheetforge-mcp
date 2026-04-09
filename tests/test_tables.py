import pytest
from openpyxl import load_workbook

from excel_mcp.server import list_tables as list_tables_tool
from excel_mcp.tables import create_excel_table, list_excel_tables
from excel_mcp.exceptions import DataError


def test_create_table_with_auto_name(tmp_workbook):
    result = create_excel_table(tmp_workbook, "Sheet1", "A1:C6")
    assert "Successfully created table" in result["message"]
    assert result["range"] == "A1:C6"
    assert result["table_name"].startswith("Table_")


def test_create_table_with_custom_name(tmp_workbook):
    result = create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="MyTable")
    assert result["table_name"] == "MyTable"

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    table_names = [t.displayName for t in ws.tables.values()]
    assert "MyTable" in table_names
    wb.close()


def test_create_table_with_custom_style(tmp_workbook):
    result = create_excel_table(
        tmp_workbook, "Sheet1", "A1:C6", table_style="TableStyleLight1"
    )
    assert "Successfully created table" in result["message"]


def test_create_table_persists_to_file(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="PersistTest")

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert len(ws.tables) == 1
    wb.close()


def test_create_table_sheet_not_found(tmp_workbook):
    with pytest.raises(DataError, match="not found"):
        create_excel_table(tmp_workbook, "NoSheet", "A1:C6")


def test_create_table_duplicate_name(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C3", table_name="Dupl")
    with pytest.raises(DataError):
        create_excel_table(tmp_workbook, "Sheet1", "A1:C3", table_name="Dupl")


def test_list_tables_returns_created_table(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="Customers", table_style="TableStyleLight1")

    result = list_excel_tables(tmp_workbook)

    assert result == [
        {
            "sheet_name": "Sheet1",
            "table_name": "Customers",
            "range": "A1:C6",
            "style": "TableStyleLight1",
            "headers": ["Name", "Age", "City"],
            "column_count": 3,
            "data_row_count": 5,
            "header_row_count": 1,
            "totals_row_count": 0,
            "totals_row_shown": False,
            "show_first_column": False,
            "show_last_column": False,
            "show_row_stripes": True,
            "show_column_stripes": False,
        }
    ]


def test_list_tables_can_filter_by_sheet(multi_sheet_workbook):
    create_excel_table(multi_sheet_workbook, "Sales", "A1:B2", table_name="SalesTable")
    create_excel_table(multi_sheet_workbook, "Inventory", "A1:B2", table_name="InventoryTable")

    result = list_excel_tables(multi_sheet_workbook, sheet_name="Inventory")

    assert result == [
        {
            "sheet_name": "Inventory",
            "table_name": "InventoryTable",
            "range": "A1:B2",
            "style": "TableStyleMedium9",
            "headers": ["Item", "Count"],
            "column_count": 2,
            "data_row_count": 1,
            "header_row_count": 1,
            "totals_row_count": 0,
            "totals_row_shown": False,
            "show_first_column": False,
            "show_last_column": False,
            "show_row_stripes": True,
            "show_column_stripes": False,
        }
    ]


def test_list_tables_tool_returns_json_envelope(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="Customers")

    payload = list_tables_tool(tmp_workbook)

    assert '"operation": "list_tables"' in payload
    assert '"table_name": "Customers"' in payload
