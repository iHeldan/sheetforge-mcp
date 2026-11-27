"""Exact executable identities, bounded transport growth and failure priority."""

import asyncio
import io
import json
from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import load_workbook
import pytest

from excel_mcp import recipes, report_receipts
from excel_mcp.exceptions import (
    PreconditionFailedError,
    ResponseTooLargeError,
    ValidationError,
)
from tests.test_preflight_audit_regressions import NODE, legacy, node_roundtrip
from tests.test_reporting_preflight import book, edit, receipt_for, state


@pytest.fixture
def source(tmp_path):
    return book(
        tmp_path / "sales.xlsx", "Sales", ["Id", "Region", "Amount"], [[1, 1e23, 100]]
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


@pytest.mark.parametrize("integer", [10**23, 10**24, 10**25, 2**53 + 1])
def test_exact_atom_does_not_merge_executable_inequivalent_numbers(integer):
    rounded = float(integer)
    assert rounded != integer
    assert report_receipts._atom(integer) != report_receipts._atom(rounded)
    for container in (lambda x: x, lambda x: [x], lambda x: {"value": x}):
        assert report_receipts._receipt_digest(container(integer)) != (
            report_receipts._receipt_digest(container(rounded))
        )


@pytest.mark.parametrize("number", [0.1, 1e-300, 5e-324, 1.7976931348623157e308])
def test_exact_atom_uses_finite_ratio_without_decimal_expansion(number):
    numerator, denominator = number.as_integer_ratio()
    assert report_receipts._atom(number) == ["number", str(numerator), str(denominator)]


@pytest.mark.parametrize("value", [float("inf"), float("-inf"), float("nan")])
def test_exact_atom_rejects_nonfinite_numbers(value):
    with pytest.raises(ValidationError, match="finite"):
        report_receipts._atom(value)


@pytest.mark.skipif(NODE is None, reason="Optional actual Node JSON check")
def test_node_unsafe_integer_filter_change_is_rejected(source, rule):
    rule["group_by"] = []
    rule["filters"] = [{"field": "Region", "op": "eq", "value": 10**23}]
    receipt = receipt_for(source, rule)
    before = state(source)
    assert receipt["rows"] == [[0]]
    assert (
        report_receipts.check_report_receipt(str(source), receipt)["status"]
        == "current"
    )
    transported = node_roundtrip(receipt)
    assert transported["recipe"]["filters"][0]["value"] != 10**23
    with pytest.raises(ValidationError, match="integrity"):
        report_receipts.check_report_receipt(str(source), transported)
    assert state(source) == before


@pytest.mark.parametrize("version", [1, 2])
@pytest.mark.parametrize("predicate", [10**23, 1e23])
def test_legacy_ambiguous_predicate_fails_closed(source, rule, version, predicate):
    rule["group_by"] = []
    rule["filters"] = [{"field": "Region", "op": "eq", "value": predicate}]
    receipt = legacy(receipt_for(source, rule), source, version)
    before = state(source)
    with pytest.raises(ValidationError, match="ambiguous numeric") as caught:
        report_receipts.check_report_receipt(str(source), receipt)
    assert caught.value.code == "legacy_receipt_numeric_ambiguity"
    assert caught.value.suggested_next_tool == "run_workbook_recipe"
    assert caught.value.details == {
        "requires_new_receipt": True,
        "automatic_upgrade": False,
    }
    assert state(source) == before


@pytest.mark.skipif(NODE is None, reason="Optional actual Node JSON check")
def test_transported_v2_collision_is_refused_not_silently_upgraded(source, rule):
    rule["group_by"] = []
    rule["filters"] = [{"field": "Region", "op": "eq", "value": 10**23}]
    receipt = legacy(receipt_for(source, rule), source, 2)
    transported = node_roundtrip(receipt)
    unsigned = {k: v for k, v in transported.items() if k != "receipt_sha256"}
    assert (
        report_receipts._receipt_digest(unsigned, atom=report_receipts._legacy_atom)
        == receipt["receipt_sha256"]
    )
    before = state(source)
    with pytest.raises(ValidationError, match="ambiguous numeric"):
        report_receipts.check_report_receipt(str(source), transported)
    assert state(source) == before


@pytest.mark.parametrize("version", [1, 2])
@pytest.mark.parametrize("amount", [0.1, 1e20, -0.0])
def test_legacy_unambiguous_values_remain_readable(tmp_path, rule, version, amount):
    path = book(
        tmp_path / "ordinary.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", amount]],
    )
    receipt = legacy(receipt_for(path, rule), path, version)
    before = state(path)
    assert (
        report_receipts.check_report_receipt(str(path), receipt)["status"] == "current"
    )
    assert state(path) == before


@pytest.mark.parametrize("location", ["source", "output"])
@pytest.mark.parametrize("version", [1, 2])
def test_legacy_ambiguous_live_values_are_not_certified(
    tmp_path, rule, version, location
):
    path = book(
        tmp_path / "safe.xlsx", "Sales", ["Id", "Region", "Amount"], [[1, "North", 100]]
    )
    receipt = legacy(receipt_for(path, rule), path, version)
    edit(
        path,
        lambda wb: setattr(
            wb["Data" if location == "source" else "Report"][
                "C2" if location == "source" else "E5"
            ],
            "value",
            1e23,
        ),
    )
    before = state(path)
    with pytest.raises(ValidationError, match="ambiguous numeric"):
        report_receipts.check_report_receipt(str(path), receipt)
    assert state(path) == before


def replace_number(path, coordinate, value, sheet=1):
    # Build a valid synthetic XLSX with an exact XML integer openpyxl cannot emit.
    payload = io.BytesIO()
    with ZipFile(path) as source, ZipFile(payload, "w") as target:
        for member in source.infolist():
            data = source.read(member)
            if member.filename == f"xl/worksheets/sheet{sheet}.xml":
                root = ET.fromstring(data)
                ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                root.find(f".//m:c[@r='{coordinate}']/m:v", ns).text = str(value)
                data = ET.tostring(root)
            target.writestr(member, data)
    path.write_bytes(payload.getvalue())


def test_exact_source_fingerprint_detects_previously_colliding_value(source, rule):
    rule["group_by"] = []
    rule["filters"] = [{"field": "Region", "op": "neq", "value": 0}]
    receipt = receipt_for(source, rule)
    replace_number(source, "B2", 10**23)
    wb = load_workbook(source)
    try:
        assert type(wb["Data"]["B2"].value) is int and wb["Data"]["B2"].value == 10**23
    finally:
        wb.close()
    before = state(source)
    checked = report_receipts.check_report_receipt(str(source), receipt)
    assert checked["source_changed"] and not checked["result_changed"]
    assert checked["status"] == "current" and not checked["output_changed"]
    assert state(source) == before


def test_exact_group_and_grid_fingerprints_detect_previously_colliding_value(
    source, rule
):
    receipt = receipt_for(source, rule)
    replace_number(source, "D5", 10**23, sheet=2)
    before = state(source)
    result = report_receipts.check_report_receipt(str(source), receipt)
    assert result["status"] == "conflict" and result["output_changed"]
    assert result["output_values_changed"] and not result["source_changed"]
    assert result["conflict_reason"] == "output_values_changed"
    assert state(source) == before


@pytest.mark.skipif(NODE is None, reason="Optional actual Node JSON check")
@pytest.mark.parametrize("legacy_version", [None, 2])
def test_near_cap_issued_receipt_survives_node_size_growth(
    tmp_path, rule, legacy_version
):
    path = book(
        tmp_path / "large.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[i, f"G{i:02d}" + "x" * 244, 1e16] for i in range(20)],
    )
    rule["output"]["range_ref"] = "D4:F40"
    receipt = receipt_for(path, rule, sample_limit=1)
    if legacy_version:
        receipt = legacy(receipt, path, legacy_version)
    transported = node_roundtrip(receipt)

    def size(value):
        return len(json.dumps(value, ensure_ascii=True, separators=(",", ":")))

    assert (
        size(receipt)
        <= 6000
        < size(transported)
        <= report_receipts.MAX_RECEIPT_INPUT_CHARS
    )
    before = state(path)
    assert (
        report_receipts.check_report_receipt(str(path), receipt)["status"] == "current"
    )
    assert (
        report_receipts.check_report_receipt(str(path), transported)["status"]
        == "current"
    )
    assert state(path) == before


def test_receipt_checker_input_remains_bounded_with_useful_recovery(source, rule):
    rule["group_by"] = []
    receipt = receipt_for(source, rule)
    receipt["rows"][0][0] = "x" * report_receipts.MAX_RECEIPT_INPUT_CHARS
    before = state(source)
    with pytest.raises(ResponseTooLargeError) as caught:
        report_receipts.check_report_receipt(str(source), receipt)
    assert caught.value.limit == 24000
    assert all("sample_limit" not in hint for hint in caught.value.hints)
    assert any("original" in hint for hint in caught.value.hints)
    assert state(source) == before


@pytest.mark.parametrize("mode", ["preview", "commit"])
def test_failed_checks_take_precedence_over_optional_receipt(tmp_path, rule, mode):
    path = book(
        tmp_path / "null.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["group_by"] = []
    rule["metrics"][0]["op"] = "min"
    rule["checks"] = [{"type": "row_count", "min": 10, "max": 20}]
    before = state(path)
    ordinary = recipes.run_workbook_recipe(str(path), rule)
    assert not ordinary["ready_to_commit"] and not ordinary["verification"]["passed"]
    if mode == "preview":
        result = recipes.run_workbook_recipe(
            str(path), rule, include_report_receipt=True
        )
        assert result == ordinary
    else:
        with pytest.raises(PreconditionFailedError) as caught:
            recipes.run_workbook_recipe(
                str(path),
                rule,
                mode=mode,
                include_report_receipt=True,
                recipe_token=ordinary["recipe_token"],
                expected_workbook_sha256=ordinary["expected_workbook_sha256"],
            )
        assert caught.value.code == "recipe_verification_failed"
        assert caught.value.details["verification"] == ordinary["verification"]
        assert "report_receipt" not in caught.value.details
    assert state(path) == before


@pytest.mark.parametrize("passed", [False, True])
def test_receipt_budget_construction_only_runs_after_passing_checks(
    source, rule, monkeypatch, passed
):
    rule["group_by"] = []
    if not passed:
        rule["checks"] = [{"type": "row_count", "min": 10, "max": 20}]
    calls = []

    def fail(*args, **kwargs):
        calls.append(True)
        raise ResponseTooLargeError(
            "prospective receipt", estimated_size=6001, limit=6000
        )

    monkeypatch.setattr(report_receipts, "make_receipt", fail)
    before = state(source)
    if passed:
        with pytest.raises(ResponseTooLargeError, match="prospective receipt"):
            recipes.run_workbook_recipe(str(source), rule, include_report_receipt=True)
    else:
        result = recipes.run_workbook_recipe(
            str(source), rule, include_report_receipt=True
        )
        assert not result["ready_to_commit"] and "report_receipt" not in result
    assert len(calls) == int(passed) and state(source) == before


@pytest.mark.skipif(NODE is None, reason="Optional actual Node/MCP check")
def test_actual_stdio_rejects_node_changed_executable_predicate(source, rule, tmp_path):
    root = Path(__file__).resolve().parents[1]
    rule["group_by"] = []
    rule["filters"] = [{"field": "Region", "op": "eq", "value": 10**23}]

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
                    preview = await call("run_workbook_recipe", args)
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
                    receipt = saved["report_receipt"]
                    assert receipt["rows"] == [[0]] and receipt["version"] == 3
                    before = state(source)
                    assert (
                        await call(
                            "check_report_receipt",
                            {
                                "filepath": str(source),
                                "report_receipt": receipt,
                            },
                        )
                    )["status"] == "current"
                    error = await call(
                        "check_report_receipt",
                        {
                            "filepath": str(source),
                            "report_receipt": node_roundtrip(receipt),
                        },
                        ok=False,
                    )
                    assert (
                        error["type"] == "ValidationError"
                        and "integrity" in error["message"]
                    )
                    assert state(source) == before

    asyncio.run(asyncio.wait_for(exercise(), timeout=45))
