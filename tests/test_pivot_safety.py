"""One-version pivot computation and checked generated values."""

import json
from pathlib import Path
from threading import Event, Thread, current_thread

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp import pivot, server, workbook


def make_pivot_book(path, values=(10, 20)):
    book = Workbook()
    book.active.title = "Data"
    book.active.append(["Group", "Value"])
    for value in values:
        book.active.append(["North", value])
    book.save(path)
    book.close()
    return path


def call(path, **kwargs):
    return json.loads(server.create_pivot_table(str(path), "Data", "A1:B3", ["Group"], ["Value"], **kwargs))


def test_pivot_holds_lock_from_source_capture_through_verified_save(tmp_path, monkeypatch):
    path = make_pivot_book(tmp_path / "race.xlsx")
    captured, release, edited = Event(), Event(), Event()
    original_resolve = pivot._resolve_field_names
    original_verify = workbook._verify_saved_workbook
    results, verified_pairs = {}, []

    def pause_after_read(*args):
        value = original_resolve(*args)
        if not captured.is_set():
            captured.set()
            assert release.wait(5)
        return value

    def verify(filepath):
        original_verify(filepath)
        if Path(filepath) == path and current_thread().name == "pivot-writer":
            book = load_workbook(path)
            try:
                verified_pairs.append((sum(book["Data"].cell(r, 2).value for r in (2, 3)), book["Data_pivot"]["B2"].value))
            finally:
                book.close()

    def editor():
        results["editor"] = json.loads(server.write_data_to_excel(str(path), "Data", [[99]], "B2"))
        edited.set()

    monkeypatch.setattr(pivot, "_resolve_field_names", pause_after_read)
    monkeypatch.setattr(workbook, "_verify_saved_workbook", verify)
    writer = Thread(target=lambda: results.update(pivot=call(path)), name="pivot-writer")
    editor_thread = Thread(target=editor)
    writer.start()
    try:
        assert captured.wait(5)
        editor_thread.start()
        assert not edited.wait(0.2)
    finally:
        release.set()
        writer.join(5)
        if editor_thread.ident is not None:
            editor_thread.join(5)
    assert not writer.is_alive() and not editor_thread.is_alive()
    assert results["pivot"]["ok"] and results["editor"]["ok"], results
    assert verified_pairs == [(30, 30)]
    book = load_workbook(path)
    assert book["Data"]["B2"].value == 99
    book.close()


def test_pivot_rejects_external_drift_after_capture_without_overwriting(tmp_path, monkeypatch):
    path = make_pivot_book(tmp_path / "outside-race.xlsx")
    captured, release = Event(), Event()
    resolve = pivot._resolve_field_names
    results = []

    def pause(*args):
        result = resolve(*args)
        if not captured.is_set():
            captured.set()
            assert release.wait(5)
        return result

    monkeypatch.setattr(pivot, "_resolve_field_names", pause)
    writer = Thread(target=lambda: results.append(call(path)))
    writer.start()
    try:
        assert captured.wait(5)
        book = load_workbook(path)
        book["Data"]["B2"] = 99
        book.save(path)
        book.close()
        outside = path.read_bytes()
    finally:
        release.set()
        writer.join(5)
    assert not writer.is_alive()
    assert not results[0]["ok"] and results[0]["error"]["code"] == "stale_workbook"
    assert path.read_bytes() == outside
    assert not list(tmp_path.glob(".*.sheetforge-*"))


def test_pivot_refuses_unrepresentable_exact_integer_total(tmp_path):
    values = (6000000000000000, 4000000000000001)
    path = make_pivot_book(tmp_path / "integer.xlsx", values)
    book = load_workbook(path)
    assert tuple(book["Data"].cell(r, 2).value for r in (2, 3)) == values
    assert all(type(book["Data"].cell(r, 2).value) is int for r in (2, 3))
    book.close()
    before = path.read_bytes()
    payload = call(path)
    assert not payload["ok"] and "round-trip" in payload["error"]["message"], payload
    assert path.read_bytes() == before


@pytest.mark.parametrize("values,aggregation,expected", [((0.1, 0.2), "SUM", 0.3), ((1, 6), "average", 3.5)])
def test_pivot_generated_float_outputs_use_documented_writer_precision(tmp_path, values, aggregation, expected):
    path = make_pivot_book(tmp_path / "float.xlsx", values)
    payload = call(path, agg_func=aggregation)
    assert payload["ok"], payload
    book = load_workbook(path)
    assert book["Data_pivot"]["B2"].value == expected
    book.close()


def test_pivot_replacement_never_selects_case_variant_user_sheet(tmp_path):
    path = make_pivot_book(tmp_path / "names.xlsx")
    book = load_workbook(path)
    book.create_sheet("data_PIVOT")["A1"] = "KEEP"
    book.save(path)
    book.close()
    before = path.read_bytes()
    assert not call(path, replace_existing=True)["ok"]
    assert path.read_bytes() == before


def test_pivot_bounds_generated_cross_product_before_expansion(tmp_path):
    path = tmp_path / "wide.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active.append(["Row", "Column", "Value"])
    for index in range(1000):
        book.active.append([f"row-{index}", f"column-{index}", 1])
    book.save(path)
    book.close()
    before = path.read_bytes()
    payload = json.loads(server.create_pivot_table(str(path), "Data", "A1:C1001", ["Row"], ["Value"], columns=["Column"]))
    assert not payload["ok"] and "output is limited" in payload["error"]["message"]
    assert path.read_bytes() == before
