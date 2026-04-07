import pytest
from excel_mcp.workbook import get_or_create_workbook


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


from excel_mcp.workbook import list_sheets


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
