import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
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
    list_named_ranges,
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
    dependent_sheet["B1"] = "=SUM(ImpactArea)"
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
    assert result["summary"]["dependent_formula_count"] == 3
    assert result["summary"]["autofilter_overlap"] is True
    assert result["summary"]["print_area_overlap"] is True
    assert result["tables"][0]["covers_header"] is True
    assert result["charts"][0]["anchor"] == "E1"
    assert result["merged_ranges"][0]["range"] == "B2:C2"
    assert result["named_ranges"][0]["name"] == "ImpactArea"
    assert result["formula_cells"]["sample"] == ["D3"]
    assert result["dependent_formulas"]["count"] == 3
    dependent_cells = {
        (item["sheet_name"], item["cell"]) for item in result["dependent_formulas"]["sample"]
    }
    assert dependent_cells == {("Sheet1", "H2"), ("Dependent", "A1"), ("Dependent", "B1")}
    named_range_reference = next(
        reference
        for item in result["dependent_formulas"]["sample"]
        if item["cell"] == "B1"
        for reference in item["references"]
        if reference.get("via_named_range") == "ImpactArea"
    )
    assert named_range_reference["intersection_range"] == "B2:F4"


def test_analyze_range_impact_reports_low_risk_for_empty_area(tmp_workbook):
    result = analyze_range_impact(tmp_workbook, "Sheet1", "H20:I21")

    assert result["summary"]["risk_level"] == "low"
    assert result["summary"]["table_count"] == 0
    assert result["summary"]["chart_count"] == 0
    assert result["summary"]["dependent_formula_count"] == 0
    assert result["hints"] == ["No overlapping workbook structures detected for this range."]


def test_analyze_range_impact_tracks_local_named_range_dependencies(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws["D1"] = "=SUM(LocalImpact)"
    ws.defined_names.add(
        DefinedName(
            "LocalImpact",
            attr_text="Sheet1!$B$2:$B$4",
        )
    )
    workbook.save(tmp_workbook)
    workbook.close()

    result = analyze_range_impact(tmp_workbook, "Sheet1", "B2:B4")

    assert result["summary"]["dependent_formula_count"] == 1
    dependency = result["dependent_formulas"]["sample"][0]
    assert dependency["cell"] == "D1"
    assert dependency["sheet_name"] == "Sheet1"
    assert dependency["references"][0]["via_named_range"] == "LocalImpact"
    assert dependency["references"][0]["intersection_range"] == "B2:B4"


def test_analyze_range_impact_tracks_transitive_formula_dependencies(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    for row in range(2, 6):
        ws[f"D{row}"] = f"=B{row}+1"
    ws["D6"] = "=SUM(D2:D5)"
    ws["E1"] = "=D6*2"
    workbook.save(tmp_workbook)
    workbook.close()

    result = analyze_range_impact(tmp_workbook, "Sheet1", "B2:C5")

    assert result["summary"]["dependent_formula_count"] == 6
    assert result["summary"]["direct_formula_count"] == 4
    assert result["summary"]["transitive_formula_count"] == 2
    assert result["dependent_formulas"]["count"] == 6
    assert result["dependent_formulas"]["direct_count"] == 4
    assert result["dependent_formulas"]["transitive_count"] == 2

    dependencies = {
        item["cell"]: item for item in result["dependent_formulas"]["sample"]
    }
    assert dependencies["D2"]["dependency_depth"] == 1
    assert dependencies["D2"]["dependency_type"] == "direct"
    assert dependencies["D6"]["dependency_depth"] == 2
    assert dependencies["D6"]["dependency_type"] == "transitive"
    assert {
        predecessor["cell"] for predecessor in dependencies["D6"]["transitive_via"]
    } == {"D2", "D3", "D4", "D5"}
    assert {
        reference["intersection_range"] for reference in dependencies["D6"]["references"]
    } == {"D2:D2", "D3:D3", "D4:D4", "D5:D5"}
    assert dependencies["E1"]["dependency_depth"] == 3
    assert dependencies["E1"]["dependency_type"] == "transitive"
    assert dependencies["E1"]["transitive_via"] == [{"sheet_name": "Sheet1", "cell": "D6"}]
    assert dependencies["E1"]["references"][0]["intersection_range"] == "D6:D6"


def test_analyze_range_impact_tracks_table_structured_reference_dependencies(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="People")

    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws["H2"] = "=SUM(People[Age])"
    ws["H3"] = "=COUNTA(People[#Headers])"
    ws["H4"] = "=COUNTA(People[[#All],[Name]])"
    ws["H5"] = "=SUM(People[[Age]:[City]])"
    workbook.save(tmp_workbook)
    workbook.close()

    data_result = analyze_range_impact(tmp_workbook, "Sheet1", "B2:C6")
    header_result = analyze_range_impact(tmp_workbook, "Sheet1", "A1:C1")

    assert data_result["summary"]["dependent_formula_count"] == 2
    assert {
        item["cell"] for item in data_result["dependent_formulas"]["sample"]
    } == {"H2", "H5"}
    assert any(
        reference.get("via_table") == "People" and reference.get("structured_reference") == "People[Age]"
        for item in data_result["dependent_formulas"]["sample"]
        for reference in item["references"]
    )
    assert any(
        reference.get("intersection_range") == "B2:C6"
        for item in data_result["dependent_formulas"]["sample"]
        for reference in item["references"]
    )

    assert header_result["summary"]["dependent_formula_count"] == 2
    assert {item["cell"] for item in header_result["dependent_formulas"]["sample"]} == {"H3", "H4"}
    assert any(
        reference.get("structured_reference") == "People[#Headers]"
        for item in header_result["dependent_formulas"]["sample"]
        for reference in item["references"]
    )
    assert any(
        reference.get("structured_reference") == "People[[#All],[Name]]"
        and reference.get("intersection_range") == "A1:A1"
        for item in header_result["dependent_formulas"]["sample"]
        for reference in item["references"]
    )


def test_analyze_range_impact_tracks_this_row_structured_references(tmp_workbook):
    create_excel_table(tmp_workbook, "Sheet1", "A1:C6", table_name="People")

    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws["H3"] = "=SUM(People[@Age])"
    workbook.save(tmp_workbook)
    workbook.close()

    matching_row = analyze_range_impact(tmp_workbook, "Sheet1", "B3:B3")
    other_row = analyze_range_impact(tmp_workbook, "Sheet1", "B2:B2")

    assert matching_row["summary"]["dependent_formula_count"] == 1
    dependency = matching_row["dependent_formulas"]["sample"][0]
    assert dependency["cell"] == "H3"
    assert dependency["references"][0]["via_table"] == "People"
    assert dependency["references"][0]["structured_reference"] == "People[@Age]"
    assert dependency["references"][0]["intersection_range"] == "B3:B3"
    assert other_row["summary"]["dependent_formula_count"] == 0


def test_analyze_range_impact_reports_validation_and_conditional_format_overlaps(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    validation = DataValidation(type="whole", operator="between", formula1="18", formula2="65")
    validation.add("B2:B6")
    ws.add_data_validation(validation)
    ws.conditional_formatting.add(
        "C2:C6",
        CellIsRule(operator="equal", formula=["\"Turku\""], fill=None),
    )
    workbook.save(tmp_workbook)
    workbook.close()

    result = analyze_range_impact(tmp_workbook, "Sheet1", "B3:C4")

    assert result["summary"]["data_validation_count"] == 1
    assert result["summary"]["conditional_format_count"] == 1
    assert result["data_validations"]["count"] == 1
    assert result["conditional_formats"]["count"] == 1
    assert result["data_validations"]["sample"][0]["applies_to"] == "B2:B6"
    assert result["data_validations"]["sample"][0]["intersection_ranges"] == ["B3:B4"]
    assert result["conditional_formats"]["sample"][0]["applies_to"] == "C2:C6"
    assert result["conditional_formats"]["sample"][0]["intersection_ranges"] == ["C3:C4"]
    assert "Selected range overlaps worksheet data validation rules." in result["hints"]
    assert "Selected range overlaps conditional formatting rules." in result["hints"]


def test_analyze_range_impact_tracks_validation_and_conditional_format_dependencies(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    dependent_sheet = workbook.create_sheet("Checks")

    validation = DataValidation(
        type="whole",
        operator="between",
        formula1="Sheet1!$B$2",
        formula2="Sheet1!$B$4",
    )
    validation.add("A1:A3")
    dependent_sheet.add_data_validation(validation)
    dependent_sheet.conditional_formatting.add(
        "B1:B3",
        FormulaRule(formula=["COUNTIF(Sheet1!$B$2:$B$4,B1)>0"]),
    )
    workbook.save(tmp_workbook)
    workbook.close()

    result = analyze_range_impact(tmp_workbook, "Sheet1", "B2:B4")

    assert result["summary"]["dependent_validation_count"] == 1
    assert result["summary"]["dependent_conditional_format_count"] == 1
    assert result["dependent_validations"]["count"] == 1
    assert result["dependent_conditional_formats"]["count"] == 1

    dependent_validation = result["dependent_validations"]["sample"][0]
    assert dependent_validation["sheet_name"] == "Checks"
    assert dependent_validation["applies_to"] == "A1:A3"
    assert {
        reference["intersection_range"] for reference in dependent_validation["references"]
    } == {"B2:B2", "B4:B4"}

    dependent_cf = result["dependent_conditional_formats"]["sample"][0]
    assert dependent_cf["sheet_name"] == "Checks"
    assert dependent_cf["applies_to"] == "B1:B3"
    assert dependent_cf["formula"] == ["COUNTIF(Sheet1!$B$2:$B$4,B1)>0"]
    assert dependent_cf["references"][0]["intersection_range"] == "B2:B4"
    assert "Validation rules elsewhere in the workbook reference the selected range." in result["hints"]
    assert "Conditional formatting rules reference the selected range." in result["hints"]


def test_list_named_ranges_includes_local_and_workbook_scope(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws.defined_names.add(DefinedName("LocalImpact", attr_text="Sheet1!$B$2:$B$4"))
    workbook.defined_names["GlobalImpact"] = DefinedName(
        "GlobalImpact",
        attr_text="Sheet1!$A$1:$A$2",
    )
    workbook.save(tmp_workbook)
    workbook.close()

    result = list_named_ranges(tmp_workbook)
    scopes = {(item["name"], item["local_sheet"]) for item in result}

    assert ("LocalImpact", "Sheet1") in scopes
    assert ("GlobalImpact", None) in scopes


def test_analyze_range_impact_prefers_same_sheet_local_named_range_over_workbook_scope(tmp_workbook):
    workbook = load_workbook(tmp_workbook)
    ws = workbook["Sheet1"]
    ws["D1"] = "=SUM(ScopedImpact)"
    ws.defined_names.add(DefinedName("ScopedImpact", attr_text="Sheet1!$B$2:$B$4"))
    workbook.defined_names["ScopedImpact"] = DefinedName(
        "ScopedImpact",
        attr_text="Sheet1!$A$1:$A$2",
    )
    workbook.save(tmp_workbook)
    workbook.close()

    local_result = analyze_range_impact(tmp_workbook, "Sheet1", "B2:B4")
    global_result = analyze_range_impact(tmp_workbook, "Sheet1", "A1:A2")

    assert local_result["summary"]["dependent_formula_count"] == 1
    assert global_result["summary"]["dependent_formula_count"] == 0
    local_reference = local_result["dependent_formulas"]["sample"][0]["references"][0]
    assert local_reference["via_named_range"] == "ScopedImpact"
    assert local_reference["intersection_range"] == "B2:B4"


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
