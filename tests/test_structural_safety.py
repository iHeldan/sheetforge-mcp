"""Reference identity and sparse structural mutation regressions."""

import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import TableFormula
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp import server, sheet
from tests.test_mutation_guards import make_book


@pytest.mark.parametrize("formula_type", [ArrayFormula, DataTableFormula])
@pytest.mark.parametrize("operation", ["insert_rows", "delete_sheet_rows", "insert_columns", "delete_sheet_columns"])
@pytest.mark.parametrize("dry_run", [False, True])
def test_full_axis_edit_refuses_array_extent_even_when_anchor_precedes_edit(tmp_path, formula_type, operation, dry_run):
    path = tmp_path / "array.xlsx"
    book = Workbook()
    book.active.title = "Data"
    formula = (ArrayFormula(ref="A1:C5", text="=ROW(A1:C5)") if formula_type is ArrayFormula
               else DataTableFormula(ref="A1:C5", r1="D1"))
    book.active["A1"] = formula
    book.active["C5"] = 5
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(getattr(server, operation)(str(path), "Data", 3, dry_run=dry_run))
    assert not payload["ok"] and "array/data-table" in payload["error"]["message"], payload
    assert path.read_bytes() == before


@pytest.mark.parametrize("operation,index", [("insert_rows", 6), ("delete_sheet_rows", 6), ("insert_columns", 4), ("delete_sheet_columns", 4)])
def test_full_axis_edit_after_array_extent_remains_supported(tmp_path, operation, index):
    path = tmp_path / "after-array.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["A1"] = ArrayFormula(ref="A1:C5", text="=ROW(A1:C5)")
    book.active["F8"] = 8
    book.save(path)
    book.close()
    payload = json.loads(getattr(server, operation)(str(path), "Data", index))
    assert payload["ok"], payload
    book = load_workbook(path)
    assert book["Data"]["A1"].value.ref == "A1:C5"
    book.close()


@pytest.mark.parametrize("operation", ["insert_rows", "delete_sheet_rows", "insert_columns", "delete_sheet_columns", "rename"])
def test_unsupported_cross_sheet_array_dependencies_fail_closed(tmp_path, operation):
    path = tmp_path / "external-array.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["D5"] = 5
    book.create_sheet("Other")["A1"] = ArrayFormula(ref="A1:A5", text="=Data!D5:D9*2")
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.rename_worksheet(str(path), "Data", "Renamed") if operation == "rename"
                         else getattr(server, operation)(str(path), "Data", 3))
    assert not payload["ok"] and "array/data-table" in payload["error"]["message"], payload
    assert path.read_bytes() == before


def test_copy_worksheet_refuses_unverifiable_array_formula_rewriting(tmp_path):
    path = tmp_path / "copy-array.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["A1"] = ArrayFormula(ref="A1:A5", text="=Data!B1:B5*2")
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.copy_worksheet(str(path), "Data", "Copy"))
    assert not payload["ok"] and "array/data-table" in payload["error"]["message"], payload
    assert path.read_bytes() == before


@pytest.mark.parametrize("formula_type", [ArrayFormula, DataTableFormula])
@pytest.mark.parametrize("operation,index", [("insert_rows", 6), ("delete_sheet_rows", 6), ("insert_columns", 4), ("delete_sheet_columns", 4)])
def test_edits_outside_array_extent_refuse_changed_array_input_references(tmp_path, formula_type, operation, index):
    path = tmp_path / "array-inputs.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["A1"] = (ArrayFormula(ref="A1:C5", text="=E8:E12*2") if formula_type is ArrayFormula
                         else DataTableFormula(ref="A1:C5", r1="E8"))
    book.active["E8"] = 8
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(getattr(server, operation)(str(path), "Data", index))
    assert not payload["ok"] and "array/data-table" in payload["error"]["message"], payload
    assert path.read_bytes() == before


@pytest.mark.parametrize("target,allowed", [("C1", False), ("C2", False), ("D1", True)])
@pytest.mark.parametrize("dry_run", [False, True])
def test_range_copy_respects_complete_target_array_extent(tmp_path, target, allowed, dry_run):
    path = tmp_path / "array-target.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["A2"] = 20
    book.active["C1"] = ArrayFormula(ref="C1:C3", text="=ROW(C1:C3)")
    book.active["C3"] = 3
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.copy_range(str(path), "Data", "A2", "A2", target, dry_run=dry_run))
    assert payload["ok"] is allowed, payload
    if not allowed or dry_run:
        assert path.read_bytes() == before


def test_case_only_rename_uses_exact_title_for_every_reference_family(tmp_path):
    path = make_book(tmp_path / "case.xlsx")
    book = load_workbook(path)
    book.create_sheet("Data_pivot")["A1"] = 5
    formulas = book["Formulas"]
    formulas["A1"], formulas["B1"] = "=Data!B2", "=Data_pivot!A1"
    book.defined_names.add(DefinedName("Window", attr_text="'Data'!$B$2:$B$3"))
    validation = DataValidation(type="whole", formula1="Data!B2", formula2="Data!B3", sqref="C1:C4")
    formulas.add_data_validation(validation)
    chart = BarChart()
    chart.add_data(Reference(book["Data"], min_col=2, min_row=1, max_row=3), titles_from_data=True)
    formulas.add_chart(chart, "F1")
    native = book["Native"].tables["Records"]
    native._initialise_columns()
    for column, name in zip(native.tableColumns, ["Id", "Value"]):
        column.name = name
    native.tableColumns[1].calculatedColumnFormula = TableFormula(attr_text="Data!B2")
    book.save(path)
    book.close()
    payload = json.loads(server.rename_worksheet(str(path), "Data", "data"))
    assert payload["ok"], payload
    book = load_workbook(path)
    try:
        assert "data" in book.sheetnames and "data_pivot" in book.sheetnames
        assert "data1" not in book.sheetnames
        assert book["Formulas"]["A1"].value in {"=data!B2", "='data'!B2"}
        assert "data_pivot" in book["Formulas"]["B1"].value
        assert "data" in book.defined_names["Window"].attr_text
        assert "data" in book["Formulas"].data_validations.dataValidation[0].formula1
        assert "data" in book["Formulas"]._charts[0].series[0].val.numRef.f
        assert "data" in book["Native"].tables["Records"].tableColumns[1].calculatedColumnFormula.text
    finally:
        book.close()


@pytest.mark.parametrize("operation", ["rename", "copy", "create", "write_create"])
def test_sheet_name_collisions_are_rejected_before_save(tmp_path, operation):
    path = make_book(tmp_path / "collision.xlsx")
    before = path.read_bytes()
    calls = {
        "rename": lambda: server.rename_worksheet(str(path), "Data", "native"),
        "copy": lambda: server.copy_worksheet(str(path), "Data", "native"),
        "create": lambda: server.create_worksheet(str(path), "native"),
        "write_create": lambda: server.write_data_to_excel(str(path), "native", [[1]]),
    }
    assert not json.loads(calls[operation]())["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("name", ["x" * 32, "Bad/Name", "'Name", " ", "bad\x00name"])
def test_invalid_titles_never_reach_saved_workbooks(tmp_path, name):
    path = make_book(tmp_path / "name.xlsx")
    before = path.read_bytes()
    assert not json.loads(server.rename_worksheet(str(path), "Data", name))["ok"]
    assert path.read_bytes() == before
    new_path = tmp_path / "new.xlsx"
    assert not json.loads(server.create_workbook(str(new_path), sheet_name=name))["ok"]
    assert not new_path.exists()


@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("direction", ["up", "left"])
def test_partial_delete_rewrites_dependencies_and_keeps_last_moved_value(tmp_path, dry_run, direction):
    path = make_book(tmp_path / "partial.xlsx")
    book = load_workbook(path)
    ws, formulas = book["Data"], book["Formulas"]
    if direction == "up":
        ws["A2"], ws["A3"], ws["A4"], ws["A5"] = 1, "=$A$5", 30, 40
        formulas["A1"] = "=Data!$A$4"
        named = "Data!$A$4:$A$5"
        start, expected_formula, moved, expected_value = "A2", "=Data!$A$3", "A4", 40
        validation = DataValidation(type="whole", formula1="Data!A4", sqref="D1:D3")
    else:
        ws["A2"], ws["B2"], ws["C2"], ws["D2"] = 1, "=$D$2", 30, 40
        formulas["A1"] = "=Data!$C$2"
        named = "Data!$C$2:$D$2"
        start, expected_formula, moved, expected_value = "A2", "=Data!$B$2", "C2", 40
        validation = DataValidation(type="whole", formula1="Data!C2", sqref="D1:D3")
    formulas.add_data_validation(validation)
    book.defined_names.add(DefinedName("Window", attr_text=named))
    formulas["B1"] = '="Data!A4"'
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.delete_range(str(path), "Data", start, start, shift_direction=direction, dry_run=dry_run))
    assert payload["ok"], payload
    assert payload["data"]["reference_updates"]["formula_updates"] >= 2
    if dry_run:
        assert path.read_bytes() == before
        return
    book = load_workbook(path)
    try:
        assert book["Formulas"]["A1"].value == expected_formula
        assert book["Formulas"]["B1"].value == '="Data!A4"'
        assert book["Data"][moved].value == expected_value
        assert book["Data"]["A2"].value in {"=$A$4", "=$C$2"}
        assert book.defined_names["Window"].attr_text != named
        assert book["Formulas"].data_validations.dataValidation[0].formula1 != validation.formula1
    finally:
        book.close()


@pytest.mark.parametrize("formula", ["=SUM(Data!A2:B5)", '=INDIRECT("Data!A4")', "=SUM(Data:Native!A4)"])
@pytest.mark.parametrize("dry_run", [False, True])
def test_unverifiable_partial_delete_dependencies_fail_closed(tmp_path, formula, dry_run):
    path = make_book(tmp_path / "ambiguous.xlsx")
    book = load_workbook(path)
    book["Formulas"]["A1"] = formula
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.delete_range(str(path), "Data", "A2", "A2", dry_run=dry_run))
    assert not payload["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("dry_run", [False, True])
def test_partial_delete_refuses_global_relative_name_without_sheet_context(tmp_path, dry_run):
    path = make_book(tmp_path / "relative-name.xlsx")
    book = load_workbook(path)
    book.defined_names.add(DefinedName("RelativeWindow", attr_text="A4:A5"))
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.delete_range(str(path), "Data", "A2", "A2", dry_run=dry_run))
    assert not payload["ok"] and "unqualified" in payload["error"]["message"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("dry_run", [False, True])
def test_partial_delete_reports_boolean_to_numeric_shift(tmp_path, dry_run):
    path = make_book(tmp_path / "typed-shift.xlsx")
    book = load_workbook(path)
    book["Data"]["B2"], book["Data"]["B3"] = True, 1
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.delete_range(str(path), "Data", "B2", "B2", dry_run=dry_run))
    assert payload["ok"], payload
    change = next(item for item in payload["changes"] if item["cell"] == "B2")
    assert change["old_value"] is True and type(change["new_value"]) is int
    if dry_run:
        assert path.read_bytes() == before
    else:
        book = load_workbook(path)
        assert type(book["Data"]["B2"].value) is int
        book.close()


@pytest.mark.parametrize("kind", ["validation", "filter", "merge", "print", "freeze"])
def test_partial_delete_refuses_layout_intersections(tmp_path, kind):
    path = make_book(tmp_path / "layout.xlsx")
    book = load_workbook(path)
    ws = book["Data"]
    if kind == "validation":
        ws.add_data_validation(DataValidation(type="whole", formula1="1", sqref="A2:A5"))
    elif kind == "filter":
        ws.auto_filter.ref = "A1:B5"
    elif kind == "merge":
        ws.merge_cells("A2:B2")
    elif kind == "print":
        ws.print_area = "A1:D5"
    else:
        ws.freeze_panes = "A2"
    book.save(path)
    book.close()
    before = path.read_bytes()
    assert not json.loads(server.delete_range(str(path), "Data", "A2", "A2"))["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("source,target", [("B2", "A1"), ("B2", "A2"), ("B2", "B1")])
def test_copy_formula_underflow_is_rejected_atomically(tmp_path, dry_run, source, target):
    path = make_book(tmp_path / "translate.xlsx")
    book = load_workbook(path)
    book["Data"]["B2"] = "=A1"
    book.save(path)
    book.close()
    before = path.read_bytes()
    assert not json.loads(server.copy_range(str(path), "Data", source, source, target, dry_run=dry_run))["ok"]
    assert path.read_bytes() == before


@pytest.mark.parametrize("operation", ["insert_rows", "insert_columns", "delete_sheet_rows", "delete_sheet_columns"])
def test_structural_operations_keep_unrelated_full_extent_sheet_sparse(tmp_path, monkeypatch, operation):
    path = make_book(tmp_path / "sparse.xlsx")
    book = load_workbook(path)
    other = book.create_sheet("Unrelated")
    other["A1"] = "keep"
    other["XFD1048576"].font = Font(bold=True)
    book.save(path)
    book.close()
    original = sheet._update_workbook_structure_references
    observed = []
    def inspect(book, **kwargs):
        before = len(book["Unrelated"]._cells)
        result = original(book, **kwargs)
        observed.append((before, len(book["Unrelated"]._cells)))
        return result
    monkeypatch.setattr(sheet, "_update_workbook_structure_references", inspect)
    payload = json.loads(getattr(server, operation)(str(path), "Data", 2))
    assert payload["ok"], payload
    assert observed == [(2, 2)]


def test_target_structural_movement_and_range_copy_are_sparse(tmp_path, monkeypatch):
    path = make_book(tmp_path / "target-sparse.xlsx")
    book = load_workbook(path)
    book["Data"]["XFD1000"].font = Font(bold=True)
    book.save(path)
    book.close()
    def never_densify(*args, **kwargs):
        pytest.fail("Dense worksheet iteration must not be used")
    from openpyxl.worksheet.worksheet import Worksheet
    monkeypatch.setattr(Worksheet, "iter_rows", never_densify)
    assert json.loads(server.insert_rows(str(path), "Data", 2))["ok"]
    assert json.loads(server.copy_range(str(path), "Data", "G1", "Z1000", "AB1"))["ok"]
    book = load_workbook(path)
    assert len(book["Data"]._cells) < 20
    assert book["Data"]["XFD1001"].font.bold
    book.close()
