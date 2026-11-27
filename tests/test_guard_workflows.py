"""Real stdio regression workflow for the four reporting guard findings."""

import asyncio
import hashlib
import json
from pathlib import Path
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import load_workbook

from tests.test_reporting_guard_regressions import (
    drift_case,
    mixed_group_book,
    numeric_xml,
)
from tests.test_reporting_preflight import book, state


def test_reporting_guards_over_actual_stdio(tmp_path):
    root = Path(__file__).resolve().parents[1]
    cases = {}
    for drift in ("key", "read_field", "write_field", "allowed"):
        folder = tmp_path / drift
        folder.mkdir()
        cases[drift] = drift_case(folder, "read_field" if drift == "allowed" else drift)
    numeric_xml(cases["allowed"][0], {"B2": str(int(1e23))})
    frozen = {
        name: (state(base), state(current))
        for name, (base, current, _, _) in cases.items()
    }
    groups = mixed_group_book(tmp_path)
    text = book(
        tmp_path / "text.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 12], [2, "North", 3]],
    )
    group_before, text_before = state(groups), state(text)
    rule = {
        "version": 1,
        "source": {"table_name": "Sales"},
        "group_by": ["Region"],
        "metrics": [{"op": "sum", "field": "Amount", "as": "Revenue"}],
        "output": {
            "sheet_name": "Report",
            "range_ref": "D4:F10",
            "allow_overwrite": True,
        },
    }

    async def exercise():
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "excel_mcp", "stdio"],
            cwd=root,
            env={"PYTHONPATH": str(root / "src")},
        )
        with (tmp_path / "server.log").open("w") as log:
            async with stdio_client(params, errlog=log) as (read, write):
                async with ClientSession(read, write) as client:
                    await client.initialize()
                    assert len((await client.list_tools()).tools) == 86

                    async def call(name, args, ok=True):
                        raw = await client.call_tool(name, args)
                        assert not raw.isError
                        payload = json.loads(raw.content[0].text)
                        assert payload["ok"] is ok, payload
                        return payload["data"] if ok else payload["error"]

                    for drift, (base, current, patch, expected) in cases.items():
                        args = {
                            "base_filepath": str(base),
                            "filepath": str(current),
                            "patch": patch,
                            "expected_base_sha256": hashlib.sha256(
                                frozen[drift][0][0]
                            ).hexdigest(),
                        }
                        preview = await call("apply_table_patch", args)
                        assert (state(base), state(current)) == frozen[drift]
                        assert preview["ready_to_commit"] is (drift == "allowed")
                        args.update(
                            mode="commit",
                            rebase_token=preview["rebase_token"],
                            expected_workbook_sha256=preview[
                                "expected_workbook_sha256"
                            ],
                        )
                        saved = await call(
                            "apply_table_patch", args, ok=drift == "allowed"
                        )
                        if drift == "allowed":
                            assert (
                                saved["persisted"] and state(base) == frozen[drift][0]
                            )
                            assert (
                                Path(saved["receipt"]["snapshot"]["path"]).read_bytes()
                                == frozen[drift][1][0]
                            )
                            wb = load_workbook(current)
                            try:
                                assert wb["Data"]["C2"].value == "new"
                                assert wb["Data"]["B2"].value == 1e23
                            finally:
                                wb.close()
                        else:
                            assert preview["counts"][expected] == 1
                            assert saved["code"] == "rebase_conflict"
                            assert (state(base), state(current)) == frozen[drift]
                            assert {p.name for p in base.parent.iterdir()} == {
                                "base.xlsx",
                                "current.xlsx",
                            }

                    base, current, _, _ = cases["key"]
                    joined = await call(
                        "rehearse_table_join",
                        {
                            "left_filepath": str(base),
                            "right_filepath": str(current),
                            "left_source": {"table_name": "Items"},
                            "right_source": {"table_name": "Items"},
                            "left_key": "Id",
                            "right_key": "Id",
                            "measure_field": "Premise",
                            "join_type": "inner",
                        },
                    )
                    assert joined["projected_rows"] == 0 and joined["requires_decision"]
                    compared = await call(
                        "probe_recipe_rules",
                        {
                            "filepath": str(groups),
                            "left_recipe": {
                                **rule,
                                "filters": [{"field": "Id", "op": "lte", "value": 2}],
                            },
                            "right_recipe": {
                                **rule,
                                "filters": [{"field": "Id", "op": "gt", "value": 2}],
                            },
                        },
                    )
                    assert (
                        not compared["real_results_differ"]
                        and state(groups) == group_before
                    )
                    denied = await call(
                        "run_workbook_recipe",
                        {
                            "filepath": str(text),
                            "recipe": {
                                **rule,
                                "filters": [
                                    {"field": "Amount", "op": " CONTAINS ", "value": 1}
                                ],
                            },
                            "include_report_receipt": True,
                        },
                        ok=False,
                    )
                    assert (
                        denied["type"] == "ValidationError"
                        and "text field" in denied["message"]
                    )
                    assert state(text) == text_before
                    args = {
                        "filepath": str(text),
                        "include_report_receipt": True,
                        "recipe": {
                            **rule,
                            "filters": [
                                {"field": "Region", "op": " Contains ", "value": "or"}
                            ],
                        },
                    }
                    preview = await call("run_workbook_recipe", args)
                    assert preview["ready_to_commit"] and state(text) == text_before
                    saved = await call(
                        "run_workbook_recipe",
                        {
                            **args,
                            "mode": "commit",
                            "recipe_token": preview["recipe_token"],
                            "expected_workbook_sha256": preview[
                                "expected_workbook_sha256"
                            ],
                        },
                    )
                    assert (
                        saved["persisted"]
                        and saved["report_receipt"]["recipe"]["filters"][0]["op"]
                        == " Contains "
                    )
                    assert (
                        Path(saved["receipt"]["snapshot"]["path"]).read_bytes()
                        == text_before[0]
                    )
                    before_check = state(text)
                    checked = await call(
                        "check_report_receipt",
                        {
                            "filepath": str(text),
                            "report_receipt": saved["report_receipt"],
                        },
                    )
                    assert (
                        checked["status"] == "current" and state(text) == before_check
                    )
                    wb = load_workbook(text)
                    try:
                        assert wb["Report"]["E5"].value == 15
                        assert [
                            [c.value for c in row] for row in wb["Data"]["A2:C3"]
                        ] == [[1, "North", 12], [2, "North", 3]]
                    finally:
                        wb.close()

    asyncio.run(asyncio.wait_for(exercise(), timeout=60))
