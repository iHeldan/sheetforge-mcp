import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp.server import freeze_panes as freeze_panes_tool
from excel_mcp.server import list_named_ranges as list_named_ranges_tool
from excel_mcp.server import autofit_columns as autofit_columns_tool
from excel_mcp.server import get_worksheet_protection as get_worksheet_protection_tool
from excel_mcp.server import set_print_area as set_print_area_tool
from excel_mcp.server import set_print_titles as set_print_titles_tool
from excel_mcp.server import set_column_widths as set_column_widths_tool
from excel_mcp.server import set_row_heights as set_row_heights_tool
from excel_mcp.server import set_worksheet_protection as set_worksheet_protection_tool
from excel_mcp.server import set_worksheet_visibility as set_worksheet_visibility_tool
from excel_mcp.server import set_autofilter as set_autofilter_tool
from excel_mcp.sheet import (
    autofit_columns,
    get_sheet_protection,
    set_auto_filter,
    set_column_widths,
    set_freeze_panes,
    set_print_area,
    set_print_titles,
    set_row_heights,
    set_sheet_protection,
    set_sheet_visibility,
)
from excel_mcp.workbook import list_named_ranges


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    assert "operation" in payload
    assert "message" in payload
    return payload


@pytest.fixture
def named_range_workbook(tmp_path):
    filepath = tmp_path / "named-ranges.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["A2"] = "Alice"
    ws["B2"] = 30
    wb.defined_names["PeopleTable"] = DefinedName("PeopleTable", attr_text="Sheet1!$A$1:$B$2")
    wb.save(filepath)
    wb.close()
    return str(filepath)


def test_set_freeze_panes_persists_value(tmp_workbook):
    result = set_freeze_panes(tmp_workbook, "Sheet1", "B2")
    assert result["freeze_panes"] == "B2"

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].freeze_panes == "B2"
    wb.close()


def test_set_freeze_panes_dry_run_does_not_persist(tmp_workbook):
    result = set_freeze_panes(tmp_workbook, "Sheet1", "B2", dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"] == "B2"

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].freeze_panes is None
    wb.close()


def test_set_autofilter_infers_used_range(tmp_workbook):
    result = set_auto_filter(tmp_workbook, "Sheet1")
    assert result["range"] == "A1:C6"

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].auto_filter.ref == "A1:C6"
    wb.close()


def test_set_autofilter_dry_run_does_not_persist(tmp_workbook):
    result = set_auto_filter(tmp_workbook, "Sheet1", "A1:C3", dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"] == "A1:C3"

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].auto_filter.ref is None
    wb.close()


def test_set_worksheet_visibility_persists_value(multi_sheet_workbook):
    result = set_sheet_visibility(multi_sheet_workbook, "Inventory", "hidden")
    assert result["visibility"] == "hidden"

    wb = load_workbook(multi_sheet_workbook)
    assert wb["Inventory"].sheet_state == "hidden"
    wb.close()


def test_set_worksheet_visibility_dry_run_does_not_persist(multi_sheet_workbook):
    result = set_sheet_visibility(multi_sheet_workbook, "Inventory", "veryHidden", dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"] == "veryHidden"

    wb = load_workbook(multi_sheet_workbook)
    assert wb["Inventory"].sheet_state == "visible"
    wb.close()


def test_set_worksheet_visibility_rejects_hiding_only_visible_sheet(tmp_workbook):
    with pytest.raises(Exception, match="only visible sheet"):
        set_sheet_visibility(tmp_workbook, "Sheet1", "hidden")


def test_get_worksheet_protection_reports_defaults(tmp_workbook):
    result = get_sheet_protection(tmp_workbook, "Sheet1")
    assert result["enabled"] is False
    assert result["password_protected"] is False
    assert "formatCells" in result["options"]


def test_set_worksheet_protection_persists_state(tmp_workbook):
    result = set_sheet_protection(
        tmp_workbook,
        "Sheet1",
        enabled=True,
        password="secret",
        options={"selectUnlockedCells": True, "formatCells": False},
    )
    assert result["enabled"] is True
    assert result["password_protected"] is True
    assert result["options"]["selectUnlockedCells"] is True
    assert result["options"]["formatCells"] is False

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.protection.sheet is True
    assert bool(ws.protection.password) is True
    assert ws.protection.selectUnlockedCells is True
    assert ws.protection.formatCells is False
    wb.close()


def test_set_worksheet_protection_dry_run_does_not_persist(tmp_workbook):
    result = set_sheet_protection(
        tmp_workbook,
        "Sheet1",
        enabled=True,
        options={"selectUnlockedCells": True},
        dry_run=True,
    )
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"]["enabled"] is True

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.protection.sheet is False
    assert ws.protection.selectUnlockedCells is False
    wb.close()


def test_set_print_area_persists_value(tmp_workbook):
    result = set_print_area(tmp_workbook, "Sheet1", "A1:C4")
    assert result["print_area"] == "A1:C4"

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].print_area == "'Sheet1'!$A$1:$C$4"
    wb.close()


def test_set_print_area_can_clear_value(tmp_workbook):
    set_print_area(tmp_workbook, "Sheet1", "A1:C4")
    result = set_print_area(tmp_workbook, "Sheet1", None)
    assert result["print_area"] is None

    wb = load_workbook(tmp_workbook)
    assert wb["Sheet1"].print_area == ""
    wb.close()


def test_set_print_titles_persists_rows_and_columns(tmp_workbook):
    result = set_print_titles(tmp_workbook, "Sheet1", rows="1:2", columns="A:B")
    assert result["print_title_rows"] == "1:2"
    assert result["print_title_columns"] == "A:B"

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.print_title_rows == "$1:$2"
    assert ws.print_title_cols == "$A:$B"
    wb.close()


def test_set_print_titles_can_clear_rows_or_columns(tmp_workbook):
    set_print_titles(tmp_workbook, "Sheet1", rows="1:2", columns="A:B")
    result = set_print_titles(tmp_workbook, "Sheet1", rows="", columns=None)
    assert result["print_title_rows"] is None
    assert result["print_title_columns"] == "A:B"

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.print_title_rows is None
    assert ws.print_title_cols == "$A:$B"
    wb.close()


def test_set_column_widths_persists_values(tmp_workbook):
    result = set_column_widths(tmp_workbook, "Sheet1", {"A": 24, "c": 18.5})
    assert result["widths"] == {"A": 24.0, "C": 18.5}

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["A"].width == 24.0
    assert ws.column_dimensions["C"].width == 18.5
    wb.close()


def test_set_column_widths_dry_run_does_not_persist(tmp_workbook):
    result = set_column_widths(tmp_workbook, "Sheet1", {"B": 30}, dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"] == 30.0

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["B"].width != 30.0
    wb.close()


def test_autofit_columns_persists_computed_width(tmp_workbook):
    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    ws["A2"] = "Extraordinarily long customer name"
    wb.save(tmp_workbook)
    wb.close()

    result = autofit_columns(tmp_workbook, "Sheet1", columns=["A"])
    assert result["columns_fitted"] == 1
    assert result["widths"]["A"] > 20

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["A"].width == result["widths"]["A"]
    wb.close()


def test_autofit_columns_dry_run_does_not_persist(tmp_workbook):
    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    ws["A2"] = "Extraordinarily long customer name"
    wb.save(tmp_workbook)
    original_width = ws.column_dimensions["A"].width
    wb.close()

    result = autofit_columns(tmp_workbook, "Sheet1", columns=["A"], dry_run=True)
    assert result["dry_run"] is True
    assert result["widths"]["A"] > 20

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["A"].width == original_width
    wb.close()


def test_set_row_heights_persists_values(tmp_workbook):
    result = set_row_heights(tmp_workbook, "Sheet1", {"1": 22, "3": 28.5})
    assert result["heights"] == {1: 22.0, 3: 28.5}

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.row_dimensions[1].height == 22.0
    assert ws.row_dimensions[3].height == 28.5
    wb.close()


def test_set_row_heights_dry_run_does_not_persist(tmp_workbook):
    result = set_row_heights(tmp_workbook, "Sheet1", {"2": 31}, dry_run=True)
    assert result["dry_run"] is True
    assert result["changes"][0]["new_value"] == 31.0

    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    assert ws.row_dimensions[2].height != 31.0
    wb.close()


def test_list_named_ranges_returns_destinations(named_range_workbook):
    result = list_named_ranges(named_range_workbook)
    assert result == [
        {
            "name": "PeopleTable",
            "type": "RANGE",
            "value": "Sheet1!$A$1:$B$2",
            "destinations": [{"sheet_name": "Sheet1", "range": "$A$1:$B$2"}],
            "local_sheet": None,
            "hidden": False,
        }
    ]


def test_freeze_panes_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(freeze_panes_tool(tmp_workbook, "Sheet1", "B2", dry_run=True))
    assert payload["operation"] == "freeze_panes"
    assert payload["dry_run"] is True
    assert payload["data"]["freeze_panes"] == "B2"


def test_set_autofilter_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(set_autofilter_tool(tmp_workbook, "Sheet1", dry_run=True))
    assert payload["operation"] == "set_autofilter"
    assert payload["dry_run"] is True
    assert payload["data"]["range"] == "A1:C6"


def test_list_named_ranges_tool_returns_json_envelope(named_range_workbook):
    payload = _load_tool_payload(list_named_ranges_tool(named_range_workbook))
    assert payload["operation"] == "list_named_ranges"
    assert payload["data"]["named_ranges"][0]["name"] == "PeopleTable"


def test_set_worksheet_visibility_tool_returns_json_envelope(multi_sheet_workbook):
    payload = _load_tool_payload(
        set_worksheet_visibility_tool(multi_sheet_workbook, "Inventory", "hidden", dry_run=True)
    )
    assert payload["operation"] == "set_worksheet_visibility"
    assert payload["dry_run"] is True
    assert payload["data"]["visibility"] == "hidden"


def test_get_worksheet_protection_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(get_worksheet_protection_tool(tmp_workbook, "Sheet1"))
    assert payload["operation"] == "get_worksheet_protection"
    assert payload["data"]["enabled"] is False


def test_set_worksheet_protection_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(
        set_worksheet_protection_tool(
            tmp_workbook,
            "Sheet1",
            enabled=True,
            options={"selectUnlockedCells": True},
            dry_run=True,
        )
    )
    assert payload["operation"] == "set_worksheet_protection"
    assert payload["dry_run"] is True
    assert payload["data"]["enabled"] is True


def test_set_print_area_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(set_print_area_tool(tmp_workbook, "Sheet1", "A1:C4", dry_run=True))
    assert payload["operation"] == "set_print_area"
    assert payload["dry_run"] is True
    assert payload["data"]["print_area"] == "A1:C4"


def test_set_print_titles_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(
        set_print_titles_tool(tmp_workbook, "Sheet1", rows="1:2", columns="A:B", dry_run=True)
    )
    assert payload["operation"] == "set_print_titles"
    assert payload["dry_run"] is True
    assert payload["data"]["print_title_rows"] == "1:2"
    assert payload["data"]["print_title_columns"] == "A:B"


def test_set_column_widths_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(set_column_widths_tool(tmp_workbook, "Sheet1", {"A": 20}, dry_run=True))
    assert payload["operation"] == "set_column_widths"
    assert payload["dry_run"] is True
    assert payload["data"]["widths"]["A"] == 20.0


def test_autofit_columns_tool_returns_json_envelope(tmp_workbook):
    wb = load_workbook(tmp_workbook)
    ws = wb["Sheet1"]
    ws["A2"] = "Extraordinarily long customer name"
    wb.save(tmp_workbook)
    wb.close()

    payload = _load_tool_payload(autofit_columns_tool(tmp_workbook, "Sheet1", ["A"], dry_run=True))
    assert payload["operation"] == "autofit_columns"
    assert payload["dry_run"] is True
    assert payload["data"]["widths"]["A"] > 20


def test_set_row_heights_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(set_row_heights_tool(tmp_workbook, "Sheet1", {"1": 24}, dry_run=True))
    assert payload["operation"] == "set_row_heights"
    assert payload["dry_run"] is True
    assert payload["data"]["heights"]["1"] == 24.0
