import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp.server import freeze_panes as freeze_panes_tool
from excel_mcp.server import list_named_ranges as list_named_ranges_tool
from excel_mcp.server import set_autofilter as set_autofilter_tool
from excel_mcp.sheet import set_auto_filter, set_freeze_panes
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
