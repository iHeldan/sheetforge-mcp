"""Real MCP report-delta workflow, using synthetic fixtures and a Decimal oracle."""

import asyncio
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import load_workbook

from tests.test_recipe_workflows import (
    HEADERS,
    _make_book,
    _oracle,
    _replace_source,
    _sales_rows,
)


def test_stdio_explanation_and_pinned_group_drilldown(tmp_path):
    root = Path(__file__).resolve().parents[1]
    before, after = tmp_path / "before.xlsx", tmp_path / "after.xlsx"
    old = _sales_rows(1000)
    new = [dict(row) for i, row in enumerate(old) if i % 23] + _sales_rows(
        200, start=1000
    )
    for i, row in enumerate(new):
        if i % 11 == 0:
            row["Amount"] += Decimal("5.25")
        if i % 19 == 0:
            row["Status"] = "posted"
        if i % 13 == 0:
            row["Region"] = "North"
    _make_book(before, old)
    _make_book(after, new)
    wb = load_workbook(after)
    try:
        _replace_source(wb, new, list(reversed(HEADERS)))
        wb.save(after)
    finally:
        wb.close()
    recipe = {
        "version": 1,
        "source": {"table_name": "Sales"},
        "filters": [{"field": "Status", "op": "eq", "value": "posted"}],
        "group_by": ["Region", "Product"],
        "metrics": [
            {"op": "sum", "field": "Amount", "as": "Revenue"},
            {"op": "count", "as": "Lines"},
        ],
        "output": {"sheet_name": "Report", "range_ref": "D4:G40"},
        "checks": [
            {"type": "unique_key", "fields": ["LineId"]},
            {
                "type": "sum_matches",
                "source_field": "Amount",
                "output_field": "Revenue",
            },
        ],
    }
    frozen = [(path.read_bytes(), path.stat().st_mtime_ns) for path in (before, after)]
    expected_before, expected_after = _oracle(old), _oracle(new)
    expected_delta = sum((row[0] for row in expected_after.values()), Decimal(0)) - sum(
        (row[0] for row in expected_before.values()), Decimal(0)
    )
    sizes = []
    parameters = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_mcp", "stdio"],
        cwd=root,
        env={"PYTHONPATH": str(root / "src")},
    )

    async def exercise():
        with (tmp_path / "stderr.log").open("w") as stderr:
            async with stdio_client(parameters, errlog=stderr) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    inventory = await session.list_tools()
                    tool = next(
                        tool
                        for tool in inventory.tools
                        if tool.name == "explain_recipe_delta"
                    )
                    assert (
                        tool.annotations.readOnlyHint is True
                        and tool.outputSchema is None
                    )

                    async def call(name, args):
                        reply = await session.call_tool(name, args)
                        assert not reply.isError
                        text = reply.content[0].text
                        sizes.append(len(text))
                        assert len(text) < 50000
                        payload = json.loads(text)
                        assert payload["ok"], payload
                        return payload["data"]

                    for path in (before, after):
                        await call(
                            "describe_dataset",
                            {
                                "filepath": str(path),
                                "table_name": "Sales",
                                "sample_rows": 3,
                            },
                        )
                    args = {
                        "before_filepath": str(before),
                        "after_filepath": str(after),
                        "recipe": recipe,
                        "key_fields": ["LineId"],
                        "sample_limit": 3,
                    }
                    result = await call("explain_recipe_delta", args)
                    assert (
                        Decimal(result["metrics"][0]["contribution_delta"])
                        == expected_delta
                    )
                    assert result["sources"]["before"]["row_count"] == 1000
                    assert result["sources"]["after"]["row_count"] == len(new)
                    group = result["groups"]["sample"][0]["group"]
                    selected = (group["Region"], group["Product"])
                    detail = await call(
                        "explain_recipe_delta",
                        {
                            **args,
                            "metric": "Revenue",
                            "group": group,
                            "expected_before_sha256": result["sources"]["before"][
                                "sha256"
                            ],
                            "expected_after_sha256": result["sources"]["after"][
                                "sha256"
                            ],
                            "expected_recipe_sha256": result["recipe_sha256"],
                        },
                    )
                    expected = (
                        expected_after.get(selected, (Decimal(0),))[0]
                        - expected_before.get(selected, (Decimal(0),))[0]
                    )
                    assert (
                        Decimal(detail["metrics"][0]["contribution_delta"]) == expected
                    )
                    assert detail["metrics"][0]["residual"] == "0"
                    assert detail["verification"]["saved_outputs_checked"] is False
                    for name, path, (content, mtime) in zip(
                        ("before", "after"), (before, after), frozen
                    ):
                        assert (
                            path.read_bytes() == content
                            and path.stat().st_mtime_ns == mtime
                        )
                        assert (
                            result["sources"][name]["sha256"]
                            == hashlib.sha256(content).hexdigest()
                        )

    asyncio.run(asyncio.wait_for(exercise(), timeout=60))
    assert len(sizes) == 4 and sum(sizes) < 100000
    (tmp_path / "explanation-receipt.json").write_text(
        json.dumps(
            {
                "fixture": "synthetic",
                "tool_calls": len(sizes),
                "response_characters": sizes,
                "total_response_characters": sum(sizes),
            }
        )
    )
