"""Scripted controls for recurring agent tasks and negative controls for grading."""

import asyncio
from copy import copy, deepcopy
from decimal import Decimal
import json
from pathlib import Path
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import load_workbook
import pytest

from tests.reporting_cases import build_case, grade_case
from tests.test_recipe_workflows import _source_and_layout


PHASES = {"weekly": 3, "consolidation": 2, "variance": 2}


def explanation_answer(result):
    return {
        "status": "completed",
        "metrics": {item["name"]: {field: item[field] for field in ("before", "after", "delta")}
                    | item["components"] for item in result["metrics"]},
        "groups": {item["group"]["Region"]: item["metrics"]["Revenue"]["delta"]
                   for item in result["groups"]["sample"]},
    }


async def exercise_case(case, call):
    """Reference client, not an LLM; calculations come from public tool responses."""
    answer = None
    if case.workflow == "variance":
        payload = await call("explain_recipe_delta", {
            **case.arguments, "before_filepath": str(case.paths["before"]),
            "after_filepath": str(case.paths["after"]), "sample_limit": 10,
        }, ok=not case.blocked)
        answer = {"status": "blocked", "reason": payload["error"]["message"]} if case.blocked else explanation_answer(payload["data"])
        grade_case(case, answer)
        return answer

    path = str(case.paths["report"])
    if case.workflow == "weekly":
        await call("inspect_workbook_contract", {"filepath": path, "contract": case.arguments["workbook_contract"]})
        name, args = "run_workbook_recipe", {"filepath": path, **case.arguments, "sample_limit": 2}
        token = "recipe_token"
    else:
        aggregate = (await call("bulk_aggregate_workbooks", {
            "filepaths": [str(case.paths[name]) for name in ("north", "south")],
            "table_name": "Sales", "schema_mode": "intersect",
            "filters": [{"field": "Status", "op": "eq", "value": "posted"}],
            "group_by": ["Region"], "metrics": [
                {"op": "sum", "field": "Amount", "as": "Revenue"},
                {"op": "count", "as": "Lines"},
                {"op": "avg", "field": "Amount", "as": "Average"},
            ],
        }))["data"]
        for row in aggregate["rows"]:
            row[1] = float(Decimal(str(row[1])).quantize(Decimal("0.01")))
            row[-1] = float(Decimal(str(row[-1])).quantize(Decimal("0.0001")))
        data = [aggregate["headers"], *aggregate["rows"]]
        data = [row + [None] * (5 - len(row)) for row in data]
        data += [[None] * 5 for _ in range(37 - len(data))]
        name, args = "apply_workbook_changeset", {
            "filepath": path, "sample_limit": 2,
            "operations": [{"tool": "write_data_to_excel", "args": {
                "sheet_name": "Report", "start_cell": "D4", "data": data,
            }}],
            "assertions": [{"type": "range_equals", "sheet_name": "Report", "range_ref": "D4:H40", "expected": data}],
        }
        token = "changeset_token"
    preview = (await call(name, args))["data"]
    assert preview["ready_to_commit"]
    assert case.paths["report"].read_bytes() == case.frozen["report"][0]
    result = (await call(name, {**args, "mode": "commit", token: preview[token],
                                "expected_workbook_sha256": preview["expected_workbook_sha256"]}))["data"]
    assert result["persisted"]
    assert Path(result["receipt"]["snapshot"]["path"]).read_bytes() == case.frozen["report"][0]
    await call("read_data_from_excel", {"filepath": path, "sheet_name": "Report", "start_cell": "D4", "end_cell": "H40"})
    grade_case(case)
    return answer


@pytest.mark.parametrize("workflow", PHASES)
def test_recurring_reporting_over_real_stdio(tmp_path, workflow):
    root = Path(__file__).resolve().parents[1]
    parameters = StdioServerParameters(command=sys.executable, args=["-m", "excel_mcp", "stdio"],
                                       cwd=root, env={"PYTHONPATH": str(root / "src")})
    sizes = []

    async def exercise():
        with (tmp_path / "server.log").open("w") as log:
            async with stdio_client(parameters, errlog=log) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()

                    async def call(name, arguments, *, ok=True):
                        result = await session.call_tool(name, arguments)
                        text = "".join(item.text for item in result.content if item.type == "text")
                        assert len(text) <= 50000
                        sizes.append(len(text))
                        payload = json.loads(text)
                        assert payload["ok"] is ok, payload
                        return payload

                    for phase in range(PHASES[workflow]):
                        case = build_case(tmp_path, workflow, phase)
                        if workflow == "consolidation" and phase == 0:
                            bad = await session.call_tool("read_excel_table", {
                                "filepath": str(case.paths["north"]), "table_name": "Sales", "row_mode": "rows",
                            })
                            assert bad.isError
                            assert case.paths["north"].read_bytes() == case.frozen["north"][0]
                        await exercise_case(case, call)

    asyncio.run(asyncio.wait_for(exercise(), timeout=120))
    assert sum(sizes) <= 150000


@pytest.mark.parametrize("fault", [None, "wrong_total", "stale_row", "stale_column", "gapped_rows", "source", "style", "chart"])
def test_report_grader_accepts_correct_result_and_rejects_false_success(tmp_path, fault):
    case = build_case(tmp_path, "consolidation" if fault == "stale_column" else "weekly", 0)
    wb = load_workbook(case.paths["report"])
    try:
        ws = wb["Report"]
        for row in ws.iter_rows(min_row=4, max_row=40, min_col=4, max_col=8):
            for cell in row:
                cell.value = None
        for index, row in enumerate([
            case.expected["headers"],
            *[[*key, *[float(value) for value in values]] for key, values in case.expected["groups"].items()],
        ], 4):
            for column, value in enumerate(row, 4):
                ws.cell(index, column, value)
        if fault == "wrong_total":
            ws["F5"] = 123
        elif fault == "stale_row":
            ws["D40"] = "stale"
        elif fault == "stale_column":
            ws["H5"] = "stale fifth column"
        elif fault == "gapped_rows":
            for column in range(4, 9):
                ws.cell(30, column, ws.cell(5, column).value)
                ws.cell(5, column).value = None
        elif fault == "source":
            wb["Data"]["B4"] = "changed source"
        elif fault == "style":
            ws["D4"].fill = copy(ws["A1"].fill)
        elif fault == "chart":
            ws._charts[0].title = "Changed outside the reserved output"
        wb.save(case.paths["report"])
    finally:
        wb.close()
    if fault is None:
        grade_case(case)
        return
    with pytest.raises(AssertionError):
        grade_case(case)


@pytest.mark.parametrize("fault", ["wrong_total", "missing_group", "wrong_component", "false_block"])
def test_variance_grader_rejects_false_success(tmp_path, fault):
    case = build_case(tmp_path, "variance", 0)
    answer = {"status": "completed", **deepcopy(case.expected)}
    if fault == "wrong_total":
        answer["metrics"]["Revenue"]["delta"] += 1
    elif fault == "missing_group":
        answer["groups"].pop("North")
    elif fault == "wrong_component":
        answer["metrics"]["Revenue"]["updated"] += 1
    else:
        answer["status"] = "blocked"
    with pytest.raises(AssertionError):
        grade_case(case, answer)


def test_original_replay_grader_checks_chart_title_not_only_root_attributes(tmp_path):
    case = build_case(tmp_path, "weekly", 0)
    path = case.paths["report"]
    baseline = _source_and_layout(path)
    wb = load_workbook(path)
    try:
        wb["Report"]._charts[0].title = "An unintended title change"
        wb.save(path)
    finally:
        wb.close()
    assert _source_and_layout(path) != baseline
