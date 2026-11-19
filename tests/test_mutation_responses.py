import hashlib
import json

import pytest
from openpyxl import load_workbook

from excel_mcp import server, tables
from excel_mcp.workbook import safe_workbook
from tests.test_value_safety import make_book


@pytest.mark.parametrize("kind", ["cells", "sheet_append", "native_append", "native_upsert"])
def test_real_response_budget_rejects_before_writing_and_retry_does_not_duplicate(tmp_path, kind):
    path = make_book(tmp_path, kind)
    original = path.read_bytes()
    rows = [{"Id": 2, "Value": "a" * 30000}, {"Id": 3, "Value": "b" * 30000}]
    def action():
        if kind == "cells":
            return server.write_data_to_excel(str(path), "Data", [["a" * 30000, "b" * 30000]], "C2", include_changes=True)
        if kind == "sheet_append":
            return server.append_table_rows(str(path), "Data", rows, include_changes=True)
        if kind == "native_append":
            return server.append_excel_table_rows(str(path), "Records", rows, include_changes=True)
        keys = [{"Id": "a" * 30000, "Value": 2}, {"Id": "b" * 30000, "Value": 3}]
        return server.upsert_excel_table_rows(str(path), "Records", "Id", keys)
    for _ in range(2):
        result = json.loads(action())
        assert not result["ok"]
        assert result["error"]["type"] == "ResponseTooLargeError"
        assert path.read_bytes() == original


def test_large_append_can_explicitly_omit_change_details(tmp_path):
    path = make_book(tmp_path, "native_append")
    payload = json.loads(server.append_excel_table_rows(str(path), "Records", [
        {"Id": 2, "Value": "a" * 30000}, {"Id": 3, "Value": "b" * 30000},
    ], include_changes=False))
    assert payload["ok"]
    book = load_workbook(path)
    assert book["Data"].tables["Records"].ref == "A1:B4"
    book.close()


def test_post_commit_metadata_failure_returns_success_with_explicit_warning(tmp_path, monkeypatch):
    path = make_book(tmp_path, "native_append")
    provider = tables._mutation_snapshot_metadata
    calls = 0

    def unavailable_after_commit(*args, **kwargs):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("injected metadata failure")
        return provider(*args, **kwargs)

    monkeypatch.setattr(tables, "_mutation_snapshot_metadata", unavailable_after_commit)
    payload = json.loads(server.append_excel_table_rows(str(path), "Records", [{"Id": 2, "Value": 3}]))
    assert payload["ok"]
    assert payload["data"]["snapshot_metadata"]["persisted"] is True
    assert payload["data"]["snapshot_metadata"]["metadata_available"] is False
    assert payload["warnings"] == [{"code": "snapshot_metadata_unavailable"}]


def test_pre_commit_metadata_failure_leaves_source_unchanged(tmp_path, monkeypatch):
    path = make_book(tmp_path, "native_append")
    original = path.read_bytes()
    monkeypatch.setattr(tables, "_mutation_snapshot_metadata", lambda *a, **k: (_ for _ in ()).throw(OSError("metadata failure")))
    payload = json.loads(server.append_excel_table_rows(str(path), "Records", [{"Id": 2, "Value": 3}]))
    assert not payload["ok"]
    assert path.read_bytes() == original


@pytest.mark.parametrize("mode", ["oversized", "serialization", "followup"])
def test_verified_commit_outcome_is_not_lost_when_late_response_fails(tmp_path, mode):
    path = make_book(tmp_path, "cells")

    class BadString:
        def __str__(self):
            raise ValueError("injected string encoding failure")

    def action():
        target = server.get_excel_path(str(path))
        with safe_workbook(target, save=True) as book:
            book["Data"]["B2"] = "saved"
        if mode == "followup":
            raise OSError("injected post-save step failure")
        return {"detail": "x" * 60000 if mode == "oversized" else BadString()}

    payload = json.loads(server._run_tool("test_saved_response", action))
    details = payload["error"]["details"] if mode == "followup" else payload["data"]
    assert details["persisted"] is True
    assert details["retry_safe"] is False
    assert details["after_sha256"] == hashlib.sha256(path.read_bytes()).hexdigest()
    assert payload["ok"] is (mode != "followup")
    if mode == "followup":
        assert payload["error"]["code"] == "write_committed_followup_failed"


def test_staging_saves_do_not_claim_that_the_requested_source_was_committed(tmp_path):
    path = make_book(tmp_path, "cells")
    staged = tmp_path / "stage.xlsx"
    staged.write_bytes(path.read_bytes())
    before = path.read_bytes()

    def action():
        server.get_excel_path(str(path))
        with safe_workbook(str(staged), save=True) as book:
            book["Data"]["B2"] = "staged only"
        return {"detail": "x" * 60000}

    payload = json.loads(server._run_tool("test_stage_response", action))
    assert not payload["ok"]
    assert payload["error"]["type"] == "ResponseTooLargeError"
    assert path.read_bytes() == before
