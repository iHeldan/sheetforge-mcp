"""Path advice must not relocate a request or weaken server confinement."""

import asyncio
import json
import os
from pathlib import Path
import sys

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook

from excel_mcp import server
from excel_mcp.exceptions import WorkbookPathError


def assert_path_schema(node):
    if node.get("type") == "string":
        assert node["description"] == server.WORKBOOK_PATH_HELP
    elif node.get("type") == "array":
        assert_path_schema(node["items"])
    else:
        branches = [item for item in node["anyOf"] if item.get("type") != "null"]
        assert len(branches) == 1
        assert_path_schema(branches[0])


def test_every_tool_workbook_path_has_guidance():
    count = 0
    for tool in server.mcp._tool_manager._tools.values():
        for name, node in tool.parameters["properties"].items():
            if name == "filepath" or name.endswith(("_filepath", "filepaths")):
                assert_path_schema(node)
                count += 1
    assert count == 88


def test_compact_schema_guidance_keeps_standalone_path_rules():
    guide = server.WORKBOOK_PATH_HELP
    for rule in (
        "Reuse supplied server path",
        "NEVER prepend client cwd",
        "Rooted: relative OK",
        "unrooted stdio: absolute",
        "ask if missing",
    ):
        assert rule in guide
    assert len(guide) <= 128


def test_initialization_keeps_full_policy_and_safe_recovery():
    instructions = server.mcp._mcp_server.instructions
    assert instructions.count(server.WORKBOOK_PATH_POLICY) == 1
    assert "relative paths use EXCEL_FILES_PATH, not cwd" in instructions
    assert (
        "Without a root (normal stdio): ask for a supplied absolute path"
        in instructions
    )
    assert "do not guess" in instructions
    assert "same tool" in instructions and "bypass confinement" in instructions


def test_complete_input_schema_payload_has_a_size_budget():
    schemas = {
        name: tool.parameters for name, tool in server.mcp._tool_manager._tools.items()
    }
    assert len(json.dumps(schemas, separators=(",", ":"))) <= 61_000


@pytest.mark.parametrize("filename", [None, False, 42, "", "  "])
def test_invalid_paths_have_structured_compatible_errors(monkeypatch, filename):
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", None)
    with pytest.raises(ValueError) as caught:
        server.get_excel_path(filename)
    assert isinstance(caught.value, WorkbookPathError)
    assert caught.value.code == "filepath_invalid"
    assert not caught.value.details["automatic_relocation"]


def test_stdio_requires_owner_path_not_client_cwd(monkeypatch):
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", None)
    reply = json.loads(server.profile_workbook("report.xlsx"))
    assert not reply["ok"]
    error = reply["error"]
    assert error["code"] == "filepath_absolute_required"
    assert error["suggested_next_tool"] == "profile_workbook"
    assert error["details"]["action"] == "ask_for_absolute_workbook_path"
    assert "client cwd" in error["message"]
    assert "replace_with" not in json.dumps(error)


def test_rejected_path_never_falls_back_to_existing_basename(tmp_path, monkeypatch):
    base = tmp_path / "root"
    base.mkdir()
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    intended = base / "report.xlsx"
    wb = Workbook()
    wb.active["A1"] = "Keep this data"
    wb.save(intended)
    wb.close()
    before, mtime = intended.read_bytes(), intended.stat().st_mtime_ns
    rejected = tmp_path / "client-cwd" / "report.xlsx"
    reply = json.loads(server.write_data_to_excel(str(rejected), "Sheet", [["bad"]]))
    assert not reply["ok"]
    error = reply["error"]
    assert error["code"] == "filepath_outside_root"
    assert error["suggested_next_tool"] == "write_data_to_excel"
    assert error["details"]["relative_to"] == "EXCEL_FILES_PATH"
    assert error["details"]["action"] == "reuse_supplied_relative_path_or_ask"
    assert not error["details"]["automatic_relocation"]
    assert "basename" in error["details"]["instruction"]
    assert str(base) not in json.dumps(reply) and str(rejected) not in json.dumps(reply)
    assert intended.read_bytes() == before and intended.stat().st_mtime_ns == mtime
    assert not rejected.exists()


@pytest.mark.parametrize("relative", ["../outside.xlsx", "nested/../../outside.xlsx"])
def test_traversal_has_recovery_but_stays_rejected(tmp_path, monkeypatch, relative):
    base = tmp_path / "root"
    base.mkdir()
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    reply = json.loads(server.profile_workbook(relative))
    assert reply["error"]["code"] == "filepath_outside_root"
    assert "ask for an approved" in reply["error"]["details"]["instruction"]


def test_symlink_escape_is_not_repaired(tmp_path, monkeypatch):
    base = tmp_path / "root"
    base.mkdir()
    (base / "escape").symlink_to(tmp_path, target_is_directory=True)
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    with pytest.raises(WorkbookPathError) as caught:
        server.get_excel_path("escape/report.xlsx")
    assert caught.value.code == "filepath_outside_root"


def test_sibling_with_same_prefix_is_not_inside_root(tmp_path, monkeypatch):
    base = tmp_path / "root"
    base.mkdir()
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    with pytest.raises(WorkbookPathError):
        server.get_excel_path(str(tmp_path / "root-other" / "report.xlsx"))


def test_bulk_inputs_validate_all_paths_before_domain_call(tmp_path, monkeypatch):
    base = tmp_path / "root"
    base.mkdir()
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    called = []
    monkeypatch.setattr(server, "union_tables_impl", lambda *a, **kw: called.append(a))
    result = json.loads(server.union_tables(["one.xlsx", "../two.xlsx"]))
    assert result["error"]["code"] == "filepath_outside_root"
    assert result["error"]["suggested_next_tool"] == "union_tables"
    assert called == []


def test_secondary_snapshot_path_still_blocks_before_mutation(tmp_path, monkeypatch):
    base = tmp_path / "root"
    base.mkdir()
    monkeypatch.setattr(server, "EXCEL_FILES_PATH", str(base))
    assert json.loads(server.create_workbook("source.xlsx"))["ok"]
    before = (base / "source.xlsx").read_bytes()
    result = json.loads(server.create_workbook_snapshot("source.xlsx", "../bad.xlsx"))
    assert result["error"]["code"] == "filepath_outside_root"
    assert result["error"]["suggested_next_tool"] == "create_workbook_snapshot"
    assert (base / "source.xlsx").read_bytes() == before
    assert not (tmp_path / "bad.xlsx").exists()


def exercise_path_workflow(tmp_path, *, installed=False):
    base = tmp_path / "exposed"
    base.mkdir()
    work = tmp_path / "client-cwd"
    work.mkdir()
    path = base / "report.xlsx"
    wb = Workbook()
    wb.active["A1"] = "Keep"
    wb.save(path)
    wb.close()
    before, mtime = path.read_bytes(), path.stat().st_mtime_ns
    source = Path(__file__).resolve().parents[1] / "src"
    # A rooted stdio adapter exercises the same path policy as HTTP/SSE, without
    # network binding or modifying global client configuration.
    params = StdioServerParameters(
        command=sys.executable,
        args=[
            "-c",
            "from excel_mcp import server; import os; server.EXCEL_FILES_PATH=os.environ['PATH_TEST_ROOT']; server.mcp.run(transport='stdio')",
        ],
        cwd=work,
        env={
            **os.environ,
            "PYTHONPATH": "" if installed else str(source),
            "PATH_TEST_ROOT": str(base),
        },
    )

    async def exercise():
        with (tmp_path / "stdio-stderr.log").open("w") as log:
            async with stdio_client(params, errlog=log) as (read, write):
                async with ClientSession(read, write) as session:
                    initialized = await session.initialize()
                    assert "NEVER prepend client cwd" in initialized.instructions
                    inventory = (await session.list_tools()).tools
                    assert len(inventory) == 83
                    profile = next(t for t in inventory if t.name == "profile_workbook")
                    assert_path_schema(profile.inputSchema["properties"]["filepath"])
                    rejected = await session.call_tool(
                        "profile_workbook", {"filepath": str(work / "report.xlsx")}
                    )
                    payload = json.loads(
                        "".join(c.text for c in rejected.content if c.type == "text")
                    )
                    assert not rejected.isError and not payload["ok"]
                    assert payload["error"]["code"] == "filepath_outside_root"
                    assert payload["error"]["suggested_next_tool"] == "profile_workbook"
                    corrected = await session.call_tool(
                        "profile_workbook", {"filepath": "report.xlsx"}
                    )
                    result = json.loads(
                        "".join(c.text for c in corrected.content if c.type == "text")
                    )
                    assert result["ok"] and result["data"]["sheet_count"] == 1
                    assert (
                        path.read_bytes() == before and path.stat().st_mtime_ns == mtime
                    )

    asyncio.run(asyncio.wait_for(exercise(), timeout=30))
    return {
        "tools": 83,
        "calls": 2,
        "same_tool_recovery": True,
        "source_unchanged": True,
        "installed": installed,
    }


def test_actual_stdio_path_error_and_same_tool_recovery(tmp_path):
    exercise_path_workflow(tmp_path)
