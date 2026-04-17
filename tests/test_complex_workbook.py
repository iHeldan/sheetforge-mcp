import json

from excel_mcp.data import quick_read as quick_read_impl
from excel_mcp.server import read_data_from_excel
from excel_mcp.workbook import analyze_range_impact, get_workbook_info, list_sheets, profile_workbook


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    return payload


def test_complex_workbook_orientation_tools_handle_chart_sheet_fixture(complex_workbook):
    sheets = list_sheets(complex_workbook)
    metadata = get_workbook_info(complex_workbook, include_ranges=True)
    profile = profile_workbook(complex_workbook)

    assert [sheet["name"] for sheet in sheets] == ["Charts", "Dashboard", "Data"]
    assert sheets[0]["sheet_type"] == "chartsheet"
    assert metadata["sheets"] == ["Charts", "Dashboard", "Data"]
    assert metadata["used_ranges"] == {
        "Dashboard": "A1:C2",
        "Data": "A1:H6",
    }
    assert profile["sheet_count"] == 3
    assert profile["chart_count"] == 2
    assert profile["table_count"] == 1
    assert profile["named_range_count"] == 2
    assert profile["sheets"][0]["sheet_type"] == "chartsheet"
    assert profile["sheets"][1]["freeze_panes"] == "A2"
    assert profile["sheets"][2]["tables"][0]["table_name"] == "SalesData"


def test_complex_workbook_range_impact_finds_cross_feature_dependencies(complex_workbook):
    result = analyze_range_impact(complex_workbook, "Data", "B2:D5")

    assert result["summary"]["risk_level"] == "high"
    assert result["summary"]["table_count"] == 1
    assert result["summary"]["named_range_count"] == 1
    assert result["summary"]["data_validation_count"] == 1
    assert result["summary"]["conditional_format_count"] == 1
    assert result["summary"]["dependent_formula_count"] == 3
    assert result["summary"]["dependent_validation_count"] == 1
    assert result["summary"]["dependent_conditional_format_count"] == 2
    assert result["tables"][0]["table_name"] == "SalesData"
    assert result["data_validations"]["sample"][0]["applies_to"] == "D2:D6"
    assert result["conditional_formats"]["sample"][0]["applies_to"] == "C2:C6"
    dependent_formula_cells = {
        (item["sheet_name"], item["cell"]) for item in result["dependent_formulas"]["sample"]
    }
    assert dependent_formula_cells == {
        ("Dashboard", "B2"),
        ("Dashboard", "C2"),
        ("Data", "H2"),
    }
    assert result["dependent_validations"]["sample"][0]["sheet_name"] == "Dashboard"
    dependent_cf_sheets = {
        item["sheet_name"] for item in result["dependent_conditional_formats"]["sample"]
    }
    assert dependent_cf_sheets == {"Dashboard", "Data"}


def test_complex_workbook_reads_support_auto_sheet_selection_and_cursors(complex_workbook):
    quick_read_result = quick_read_impl(complex_workbook)
    first_page = _load_tool_payload(
        read_data_from_excel(
            complex_workbook,
            "Data",
            start_cell="A1",
            end_cell="D6",
            max_rows=2,
            max_cols=2,
            values_only=True,
        )
    )
    down_page = _load_tool_payload(
        read_data_from_excel(
            complex_workbook,
            "Data",
            cursor=first_page["data"]["continuations"]["down"]["cursor"],
            values_only=True,
        )
    )

    assert quick_read_result["sheet_name"] == "Dashboard"
    assert quick_read_result["auto_selected_sheet"] is True
    assert quick_read_result["headers"] == ["Executive Dashboard", None, None]

    assert first_page["data"]["range"] == "A1:B2"
    assert set(first_page["data"]["continuations"]) == {"down", "right"}
    assert down_page["data"]["range"] == "A3:B4"
    assert down_page["data"]["values"] == [
        ["Gadget", 24],
        ["Thing", 18],
    ]
