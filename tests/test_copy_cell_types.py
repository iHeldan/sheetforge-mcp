"""Copy and reference rewriting must preserve explicit XLSX cell semantics."""

import json

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp import server, workbook
from excel_mcp.exceptions import DataError, WorkbookError
from excel_mcp.values import verify_cell_writes


def make_book(path, value="=A2", data_type="s"):
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet["A2"] = 7
    sheet["C1"] = value
    sheet["C1"].data_type = data_type
    sheet["C1"].quotePrefix = data_type == "s"
    book.save(path)
    book.close()
    return path


@pytest.mark.parametrize("value,data_type,expected", [
    ("=A2", "s", "=A2"), ("#N/A", "s", "#N/A"),
    ("#N/A", "e", "#N/A"), ("=A2", "f", "=B2"),
    ("=" + "a" * 9000, "s", "=" + "a" * 9000),
])
@pytest.mark.parametrize("dry_run", [False, True])
def test_copy_preserves_source_type_and_only_translates_real_formulas(
    tmp_path, monkeypatch, value, data_type, expected, dry_run,
):
    path = make_book(tmp_path / "types.xlsx", value, data_type)
    before = path.read_bytes()
    book = load_workbook(path)
    assert (book["Data"]["C1"].value, book["Data"]["C1"].data_type) == (value, data_type)
    book.close()
    original, observed = workbook.inspect_package_fidelity, []

    def inspect(baseline, candidate, **kwargs):
        book = load_workbook(candidate)
        try:
            cell = book["Data"]["D1"]
            observed.append((cell.value, cell.data_type, cell.quotePrefix))
        finally:
            book.close()
        return original(baseline, candidate, **kwargs)

    monkeypatch.setattr(workbook, "inspect_package_fidelity", inspect)
    payload = json.loads(server.copy_range(str(path), "Data", "C1", "C1", "D1", dry_run=dry_run))
    assert payload["ok"], payload
    assert observed == [(expected, data_type, data_type == "s")]
    if dry_run:
        assert path.read_bytes() == before
    else:
        book = load_workbook(path)
        assert (book["Data"]["D1"].value, book["Data"]["D1"].data_type) == (expected, data_type)
        book.close()


@pytest.mark.parametrize("operation", ["rename", "insert", "delete", "copy_sheet", "partial_delete"])
def test_structure_rewrites_skip_literal_formula_text(tmp_path, operation):
    value = "=Data!B2"
    path = make_book(tmp_path / "structure.xlsx", value)
    book = load_workbook(path)
    report = book.create_sheet("Report")
    report["A1"] = value
    report["A1"].data_type = "s"
    report["A1"].quotePrefix = True
    report["B1"] = value
    book.save(path)
    book.close()
    calls = {
        "rename": lambda: server.rename_worksheet(str(path), "Data", "Renamed"),
        "insert": lambda: server.insert_rows(str(path), "Data", 2),
        "delete": lambda: server.delete_sheet_rows(str(path), "Data", 2),
        "copy_sheet": lambda: server.copy_worksheet(str(path), "Data", "Copy"),
        "partial_delete": lambda: server.delete_range(str(path), "Data", "B2", "B2"),
    }
    payload = json.loads(calls[operation]())
    assert payload["ok"], payload
    book = load_workbook(path)
    try:
        cell = book["Report"]["A1"]
        assert (cell.value, cell.data_type) == (value, "s")
        target = book["Renamed" if operation == "rename" else "Copy" if operation == "copy_sheet" else "Data"]
        assert (target["C1"].value, target["C1"].data_type) == (value, "s")
        assert book["Report"]["B1"].data_type == "f"
    finally:
        book.close()


def test_copy_reports_type_only_change_and_keeps_overlapping_source_snapshot(tmp_path):
    path = make_book(tmp_path / "overlap.xlsx")
    book = load_workbook(path)
    book["Data"]["C2"] = "#N/A"
    book["Data"]["C2"].data_type = "s"
    book.save(path)
    book.close()
    payload = json.loads(server.copy_range(str(path), "Data", "C1", "C2", "C2"))
    assert payload["ok"], payload
    book = load_workbook(path)
    assert [(book["Data"][f"C{row}"].value, book["Data"][f"C{row}"].data_type) for row in (2, 3)] == [
        ("=A2", "s"), ("#N/A", "s"),
    ]
    book["Data"]["D2"] = "=A2"
    book.save(path)
    book.close()
    payload = json.loads(server.copy_range(str(path), "Data", "C2", "C2", "D2"))
    assert payload["ok"], payload
    change = payload["changes"][0]
    assert change["old_value"] == change["new_value"] == "=A2"
    assert (change["old_data_type"], change["new_data_type"]) == ("f", "s")


def test_serialized_copy_verification_rejects_type_only_drift():
    book = Workbook()
    book.active["A1"] = "=1+1"
    with pytest.raises(DataError, match="Serialized cell type"):
        verify_cell_writes(book, {("Sheet", 1, 1): "=1+1"}, data_types={("Sheet", 1, 1): "s"})
    book.close()


def test_formula_readers_do_not_treat_literal_text_as_an_active_formula(tmp_path):
    path = make_book(tmp_path / "reader.xlsx", "=Missing!A1")
    book = load_workbook(path)
    assert list(workbook._iter_recorded_formula_cells(book["Data"])) == []
    assert workbook._cells_with_broken_formula_references(book, ws=book["Data"], sheet_name="Data") == []
    book.close()
    with pytest.raises(WorkbookError, match="does not contain a formula"):
        workbook.explain_formula_cell(str(path), "Data", "C1")
