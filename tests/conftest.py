import pytest
from openpyxl import Workbook


@pytest.fixture
def tmp_workbook(tmp_path):
    """Create a temporary Excel workbook with sample data."""
    filepath = str(tmp_path / "test.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["C1"] = "City"

    # Data rows
    data = [
        ("Alice", 30, "Helsinki"),
        ("Bob", 25, "Tampere"),
        ("Carol", 35, "Turku"),
        ("Dave", 28, "Oulu"),
        ("Eve", 32, "Espoo"),
    ]
    for i, (name, age, city) in enumerate(data, start=2):
        ws[f"A{i}"] = name
        ws[f"B{i}"] = age
        ws[f"C{i}"] = city

    wb.save(filepath)
    wb.close()
    return filepath


@pytest.fixture
def empty_workbook(tmp_path):
    """Create an empty Excel workbook."""
    filepath = str(tmp_path / "empty.xlsx")
    wb = Workbook()
    wb.save(filepath)
    wb.close()
    return filepath


@pytest.fixture
def multi_sheet_workbook(tmp_path):
    """Create a workbook with multiple sheets."""
    filepath = str(tmp_path / "multi.xlsx")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sales"
    ws1["A1"] = "Product"
    ws1["B1"] = "Revenue"
    ws1["A2"] = "Widget"
    ws1["B2"] = 1500

    ws2 = wb.create_sheet("Inventory")
    ws2["A1"] = "Item"
    ws2["B1"] = "Count"
    ws2["A2"] = "Widget"
    ws2["B2"] = 42

    wb.save(filepath)
    wb.close()
    return filepath
