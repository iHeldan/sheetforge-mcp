"""Regressions for the independent reporting-preflight audit."""

import asyncio
from copy import deepcopy
import json
from pathlib import Path
import shutil
import subprocess
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import load_workbook
import pytest

from excel_mcp import recipes, report_receipts, reporting_evidence
from excel_mcp.exceptions import DataError, ValidationError
from excel_mcp.report_receipts import check_report_receipt
from excel_mcp.rule_probe import probe_recipe_rules
from tests.test_reporting_preflight import book, receipt_for, edit, state

NODE = shutil.which("node")


@pytest.fixture
def source(tmp_path):
    return book(
        tmp_path / "sales.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 120.5], [2, "North", 119.5]],
    )


@pytest.fixture
def rule():
    return {
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


def contract():
    return {
        "version": 1,
        "name": "Sales",
        "source": {"table_name": "Sales"},
        "row_grain": "one sale",
        "key_fields": ["id"],
        "fields": [
            {"name": "id", "headers": ["Id"], "type": "integer"},
            {"name": "region", "headers": ["Region"], "type": "string"},
            {"name": "amount", "headers": ["Amount"], "type": "number", "min": 0.0},
        ],
        "filters": [{"field": "amount", "op": "gte", "value": -0.0}],
    }


def node_roundtrip(value):
    return json.loads(
        subprocess.check_output(
            [
                NODE,
                "-e",
                "process.stdout.write(JSON.stringify(JSON.parse(require('fs').readFileSync(0,'utf8'))))",
            ],
            input=json.dumps(value),
            text=True,
            timeout=15,
        )
    )


@pytest.mark.parametrize(
    "number,integer", [(240.0, 240), (-0.0, 0), (1e16, 10**16), (1e20, 10**20)]
)
def test_receipt_numeric_digest_preserves_json_number_identity(number, integer):
    for value in (
        number,
        [number],
        {"rows": [[number]], "recipe": {"value": number}, "contract": {"min": number}},
    ):
        other = json.loads(
            json.dumps(value).replace(json.dumps(number), json.dumps(integer))
        )
        assert report_receipts._receipt_digest(
            value
        ) == report_receipts._receipt_digest(other)


@pytest.mark.parametrize(
    "left,right",
    [
        (1, True),
        (0, False),
        ("1", 1),
        (1, ["number", "1"]),
        ({"x": 1}, [["x", 1]]),
        (2**53 + 1, 2**53),
        (1e-7, 2e-7),
    ],
)
def test_receipt_numeric_digest_does_not_hide_type_value_or_container_changes(
    left, right
):
    assert report_receipts._receipt_digest(left) != report_receipts._receipt_digest(
        right
    )


@pytest.mark.parametrize("use_contract", [False, True])
def test_current_receipt_roundtrip_normalizes_rows_recipe_and_contract(
    source, rule, use_contract
):
    extra = {}
    if use_contract:
        rule["group_by"] = ["region"]
        rule["metrics"][0]["field"] = "amount"
        extra["workbook_contract"] = contract()
    rule["filters"] = [
        {"field": "amount" if use_contract else "Amount", "op": "gte", "value": 0.0}
    ]
    receipt = receipt_for(source, rule, **extra)
    assert receipt["version"] == 3 and type(receipt["rows"][0][-1]) is float
    transported = json.loads(
        json.dumps(receipt)
        .replace("240.0", "240")
        .replace("-0.0", "0")
        .replace("0.0", "0")
    )
    before = state(source)
    assert check_report_receipt(str(source), transported)["status"] == "current"
    assert state(source) == before
    bad = deepcopy(transported)
    bad["rows"][0][-1] = 241
    with pytest.raises(ValidationError, match="integrity"):
        check_report_receipt(str(source), bad)
    assert state(source) == before


@pytest.mark.skipif(
    NODE is None, reason="Optional actual Node JSON interoperability check"
)
@pytest.mark.parametrize("use_contract", [False, True])
def test_actual_node_roundtrip_keeps_valid_receipt(source, rule, use_contract):
    extra = {}
    if use_contract:
        rule["group_by"] = ["region"]
        rule["metrics"][0]["field"] = "amount"
        extra["workbook_contract"] = contract()
    rule["filters"] = [
        {"field": "amount" if use_contract else "Amount", "op": "gte", "value": -0.0}
    ]
    receipt = receipt_for(source, rule, **extra)
    before = state(source)
    result = node_roundtrip(receipt)
    assert type(result["rows"][0][-1]) is int
    assert check_report_receipt(str(source), result)["status"] == "current"
    assert state(source) == before


@pytest.mark.skipif(
    NODE is None, reason="Optional actual Node JSON interoperability check"
)
def test_node_roundtrip_rejects_actual_unsafe_integer_changes(source, rule):
    rule["filters"] = [{"field": "Id", "op": "neq", "value": 2**53 + 1}]
    receipt = receipt_for(source, rule)
    result = node_roundtrip(receipt)
    assert result["recipe"]["filters"][0]["value"] == 2**53
    before = state(source)
    with pytest.raises(ValidationError, match="integrity"):
        check_report_receipt(str(source), result)
    assert state(source) == before


def legacy(receipt, path, version=1):
    # Reproduce the historical algorithms, not the current receipt hashes.
    def atom(value):
        if type(value) in (int, float):
            return ["number", reporting_evidence.text(reporting_evidence.number(value))]
        return [type(value).__name__, value]

    def canonical(value):
        if isinstance(value, dict):
            return ["object", [[k, canonical(value[k])] for k in sorted(value)]]
        if isinstance(value, list):
            return ["array", [canonical(item) for item in value]]
        return atom(value)

    result = deepcopy(receipt)
    result["version"] = version
    compiled = result["recipe"]
    wb = load_workbook(path)
    try:
        if result["workbook_contract"] is not None:
            from excel_mcp.contracts import compile_contract_recipe

            compiled, _ = compile_contract_recipe(
                wb, compiled, result["workbook_contract"], 3
            )
        source = recipes._load_table(wb, compiled)
        fields = recipes._source_fields(compiled)
        indexes = [source["headers"].index(field) for field in fields]
        result["source_content_sha256"] = reporting_evidence.digest(
            {
                "fields": fields,
                "rows": [[atom(row[i]) for i in indexes] for row in source["rows"]],
            }
        )
        output = compiled["output"]
        result["output_grid_sha256"] = reporting_evidence.digest(
            [
                [atom(c.value) for c in row]
                for row in wb[output["sheet_name"]][output["range_ref"]]
            ]
        )
    finally:
        wb.close()
    if version == 1:
        result.pop("output_headers")
    result.pop("receipt_sha256")
    result["receipt_sha256"] = reporting_evidence.digest(
        result if version == 1 else canonical(result)
    )
    return result


def test_legacy_receipt_retains_exact_checksum_validation(source, rule):
    receipt = legacy(receipt_for(source, rule), source)
    before = state(source)
    assert check_report_receipt(str(source), receipt)["status"] == "current"
    broken = json.loads(json.dumps(receipt).replace("240.0", "240"))
    with pytest.raises(ValidationError, match="integrity"):
        check_report_receipt(str(source), broken)
    assert state(source) == before


@pytest.mark.parametrize("use_contract", [False, True])
@pytest.mark.parametrize("legacy_format", [None, 1, 2])
def test_alias_headers_match_the_actual_writer_without_changing_request(
    source, rule, use_contract, legacy_format
):
    rule["metrics"][0]["as"] = " Revenue "
    extra = {}
    if use_contract:
        rule["group_by"] = ["region"]
        rule["metrics"][0]["field"] = "amount"
        extra["workbook_contract"] = contract()
    original = deepcopy(rule)
    receipt = receipt_for(source, rule, **extra)
    expected_alias = " Revenue " if use_contract else "Revenue"
    assert receipt["output_headers"][-1] == expected_alias
    if legacy_format:
        receipt = legacy(receipt, source, legacy_format)
    before = state(source)
    assert check_report_receipt(str(source), receipt)["status"] == "current"
    assert state(source) == before and rule == original
    edit(source, lambda wb: setattr(wb["Report"]["E5"], "value", 241))
    before = state(source)
    checked = check_report_receipt(str(source), receipt)
    assert checked["status"] == "conflict"
    assert checked["comparison_sample"][0]["previous"] == {expected_alias: 240}
    assert checked["comparison_sample"][0]["stored"] == {expected_alias: 241}
    assert state(source) == before


def test_recomputed_checksum_does_not_authorize_inconsistent_headers(source, rule):
    receipt = receipt_for(source, rule)
    receipt["output_headers"][0] = "Other"
    receipt["receipt_sha256"] = report_receipts._receipt_digest(
        {k: v for k, v in receipt.items() if k != "receipt_sha256"}
    )
    before = state(source)
    with pytest.raises(ValidationError, match="headers do not match"):
        check_report_receipt(str(source), receipt)
    assert state(source) == before


def test_invalid_synthetic_sum_case_does_not_abort_real_results(tmp_path, rule):
    path = book(
        tmp_path / "nulls.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["group_by"] = []
    rule["filters"] = [{"field": "Amount", "op": "eq", "value": "pending"}]
    before = state(path)
    result = probe_recipe_rules(str(path), rule, rule)
    assert (
        result["status"] == "no_difference_found" and not result["equivalence_proven"]
    )
    assert result["invalid_synthetic_cases"] > 0
    assert result["real_results"][0]["rows_sample"] == [[0]]
    assert state(path) == before


@pytest.mark.parametrize("reverse", [False, True])
def test_unknown_synthetic_types_do_not_invalidate_valid_all_null_source(
    tmp_path, rule, reverse
):
    path = book(
        tmp_path / "nulls.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["group_by"] = []
    left, right = deepcopy(rule), deepcopy(rule)
    left["filters"] = [{"field": "Amount", "op": "eq", "value": "pending"}]
    right["filters"] = [{"field": "Amount", "op": "eq", "value": 1}]
    if reverse:
        left, right = right, left
    before = state(path)
    result = probe_recipe_rules(str(path), left, right)
    assert not result["real_results_differ"] and result["invalid_synthetic_cases"] > 0
    assert state(path) == before


@pytest.mark.parametrize("exception", [RuntimeError, DataError])
def test_real_source_errors_are_not_swallowed(source, rule, monkeypatch, exception):
    def fail(*args, **kwargs):
        raise exception("actual source failure")

    monkeypatch.setattr(recipes, "_aggregate", fail)
    before = state(source)
    with pytest.raises(exception, match="actual source failure"):
        probe_recipe_rules(str(source), rule, rule)
    assert state(source) == before


def test_unexpected_synthetic_engine_error_is_not_swallowed(source, rule, monkeypatch):
    aggregate = recipes._aggregate

    def fail(data, definition):
        i = data["headers"].index("Amount")
        if any(row[i] is not None and row[i] < 0 for row in data["rows"]):
            raise RuntimeError("synthetic engine failure")
        return aggregate(data, definition)

    monkeypatch.setattr(recipes, "_aggregate", fail)
    before = state(source)
    with pytest.raises(RuntimeError, match="synthetic engine failure"):
        probe_recipe_rules(str(source), rule, rule)
    assert state(source) == before


@pytest.mark.parametrize(
    "group_values", [[None] * 8 + ["North"], [None] * 8 + [True], [None] * 8 + [10]]
)
def test_synthetic_types_use_full_source_not_just_first_eight_rows(
    tmp_path, rule, group_values
):
    rows = [[i, value, 100] for i, value in enumerate(group_values, 1)]
    path = book(
        tmp_path / "leading-nulls.xlsx", "Sales", ["Id", "Region", "Amount"], rows
    )
    right = deepcopy(rule)
    right["filters"] = [{"field": "Amount", "op": "gte", "value": 0}]
    before = state(path)
    result = probe_recipe_rules(str(path), rule, right)
    witness = result["witness"]["record"]
    assert witness["Region"] is None or type(witness["Region"]) is type(
        group_values[-1]
    )
    assert witness["Amount"] < 0
    assert state(path) == before


def test_positions_only_change_plus_source_refresh_remains_explicit_conflict(
    tmp_path, rule
):
    path = book(
        tmp_path / "sort.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 100], [2, "South", 120]],
    )
    receipt = receipt_for(path, rule)

    def mutate(wb):
        ws = wb["Report"]
        for col in (4, 5):
            ws.cell(5, col).value, ws.cell(6, col).value = (
                ws.cell(6, col).value,
                ws.cell(5, col).value,
            )
        wb["Data"]["C2"] = 125

    edit(path, mutate)
    before = state(path)
    checked = check_report_receipt(str(path), receipt)
    assert checked["status"] == "conflict"
    assert (
        checked["output_positions_only_changed"]
        and not checked["output_values_changed"]
    )
    assert checked["conflict_reason"] == "output_positions_changed"
    assert state(path) == before


@pytest.mark.parametrize(
    "cell,value,reason,values_changed",
    [
        ("E5", 999, "output_values_changed", True),
        ("D4", "Changed", "output_structure_changed", None),
        ("F10", "padding", "output_structure_changed", None),
    ],
)
def test_conflict_reasons_do_not_mislabel_content_changes_as_positions(
    source, rule, cell, value, reason, values_changed
):
    receipt = receipt_for(source, rule)
    edit(source, lambda wb: setattr(wb["Report"][cell], "value", value))
    before = state(source)
    result = check_report_receipt(str(source), receipt)
    assert result["status"] == "conflict" and result["conflict_reason"] == reason
    assert result["output_values_changed"] is values_changed
    assert not result["output_positions_only_changed"] and state(source) == before


@pytest.mark.skipif(
    NODE is None, reason="Optional actual Node JSON interoperability check"
)
def test_actual_stdio_checks_transported_receipt_and_invalid_synthetic_recovery(
    source, rule, tmp_path
):
    root = Path(__file__).resolve().parents[1]
    nulls = book(
        tmp_path / "nulls.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["metrics"][0]["as"] = " Revenue "
    rule["filters"] = [{"field": "Amount", "op": "gte", "value": -0.0}]

    async def exercise():
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "excel_mcp", "stdio"],
            cwd=root,
            env={"PYTHONPATH": str(root / "src")},
        )
        with (tmp_path / "stdio.log").open("w") as log:
            async with stdio_client(params, errlog=log) as (read, write):
                async with ClientSession(read, write) as client:
                    await client.initialize()
                    assert len((await client.list_tools()).tools) == 86

                    async def call(name, args, ok=True):
                        result = await client.call_tool(name, args)
                        assert not result.isError
                        payload = json.loads(result.content[0].text)
                        assert payload["ok"] is ok, payload
                        return payload["data"] if ok else payload["error"]

                    args = {
                        "filepath": str(source),
                        "recipe": rule,
                        "include_report_receipt": True,
                    }
                    before = state(source)
                    preview = await call("run_workbook_recipe", args)
                    assert state(source) == before and "report_receipt" not in preview
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
                    receipt = node_roundtrip(saved["report_receipt"])
                    assert receipt["output_headers"] == ["Region", "Revenue"]
                    before = state(source), state(nulls)
                    checked = await call(
                        "check_report_receipt",
                        {"filepath": str(source), "report_receipt": receipt},
                    )
                    assert checked["status"] == "current" and checked["version"] == 3
                    bad = deepcopy(receipt)
                    bad["rows"][0][-1] += 1
                    error = await call(
                        "check_report_receipt",
                        {"filepath": str(source), "report_receipt": bad},
                        ok=False,
                    )
                    assert (
                        error["type"] == "ValidationError"
                        and "integrity" in error["message"]
                    )
                    probe = deepcopy(rule)
                    probe["group_by"] = []
                    probe["filters"] = [
                        {"field": "Amount", "op": "eq", "value": "pending"}
                    ]
                    result = await call(
                        "probe_recipe_rules",
                        {
                            "filepath": str(nulls),
                            "left_recipe": probe,
                            "right_recipe": probe,
                        },
                    )
                    assert (
                        result["status"] == "no_difference_found"
                        and result["invalid_synthetic_cases"] > 0
                    )
                    assert (
                        not result["equivalence_proven"]
                        and (state(source), state(nulls)) == before
                    )

    asyncio.run(asyncio.wait_for(exercise(), timeout=45))


@pytest.mark.skipif(
    NODE is None, reason="Optional actual Node JSON interoperability check"
)
@pytest.mark.parametrize("number", [240.0, -0.0, 1e16, 1e20, 1e-6, 1e-7])
def test_actual_node_digest_identity_across_numeric_spellings(number):
    value = {
        "rows": [[number]],
        "recipe": {"value": number},
        "contract": {"min": number},
    }
    assert report_receipts._receipt_digest(value) == report_receipts._receipt_digest(
        node_roundtrip(value)
    )
