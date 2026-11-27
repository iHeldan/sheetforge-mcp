"""Independent arithmetic and no-write controls for reporting preflight."""

import asyncio
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table
import pytest

from excel_mcp import report_receipts
from excel_mcp.exceptions import ResponseTooLargeError, ValidationError
from excel_mcp.join_rehearsal import rehearse_table_join
from excel_mcp.recipes import run_workbook_recipe
from excel_mcp.report_receipts import check_report_receipt
from excel_mcp.rule_probe import probe_recipe_rules


def book(path, name, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    ws.add_table(
        Table(displayName=name, ref=f"A1:{chr(64 + len(headers))}{len(rows) + 1}")
    )
    wb.save(path)
    wb.close()
    return path


def state(path):
    return path.read_bytes(), path.stat().st_mtime_ns


def edit(path, fn):
    wb = load_workbook(path)
    try:
        fn(wb)
        wb.save(path)
    finally:
        wb.close()


@pytest.fixture
def source(tmp_path):
    return book(
        tmp_path / "myynti-ää.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 100], [2, "South", 120]],
    )


@pytest.fixture
def recipe():
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


def receipt_for(path, recipe, **extra):
    before = state(path)
    preview = run_workbook_recipe(
        str(path), recipe, include_report_receipt=True, **extra
    )
    assert preview["ready_to_commit"] and "report_receipt" not in preview
    assert state(path) == before
    result = run_workbook_recipe(
        str(path),
        recipe,
        mode="commit",
        include_report_receipt=True,
        expected_workbook_sha256=preview["expected_workbook_sha256"],
        recipe_token=preview["recipe_token"],
        **extra,
    )
    assert result["persisted"]
    return json.loads(json.dumps(result["report_receipt"]))


def test_join_net_zero_does_not_hide_duplicate_and_loss(tmp_path):
    left = book(
        tmp_path / "left.xlsx", "Invoices", ["Id", "Amount"], [[1, 120], [2, 120]]
    )
    right = book(tmp_path / "right.xlsx", "Shipments", ["Id"], [[1], [1]])
    before = state(left), state(right)
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "Invoices"},
        {"table_name": "Shipments"},
        "Id",
        "Id",
        "Amount",
        "inner",
    )
    assert result["measure"] == {
        "field": "Amount",
        "before": "240",
        "after": "240",
        "delta": "0",
        "extra_signed": "120",
        "lost_signed": "120",
        "extra_absolute": "120",
        "lost_absolute": "120",
    }
    assert result["requires_decision"]
    assert result["multiplied_left_rows"] == result["dropped_left_rows"] == 1
    assert result["samples"]["multiplied"][0]["left_row"] == 2
    assert (state(left), state(right)) == before


@pytest.mark.parametrize("join_type,expected", [("left", 360), ("inner", 240)])
def test_join_projection_matches_independent_nested_loop(tmp_path, join_type, expected):
    left_rows, right_rows = [[1, 120], [2, 120]], [[1], [1]]
    left = book(tmp_path / "left.xlsx", "L", ["Id", "Amount"], left_rows)
    right = book(tmp_path / "right.xlsx", "R", ["Id"], right_rows)
    projected = [a for a in left_rows for b in right_rows if a[0] == b[0]]
    if join_type == "left":
        projected += [a for a in left_rows if not any(a[0] == b[0] for b in right_rows)]
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "L"},
        {"table_name": "R"},
        "Id",
        "Id",
        "Amount",
        join_type,
    )
    assert int(result["measure"]["after"]) == sum(r[1] for r in projected) == expected
    assert result["projected_rows"] == len(projected)


def test_join_legitimate_many_to_one_count_measure_is_conserved(tmp_path):
    left = book(
        tmp_path / "ship.xlsx", "Shipments", ["Id", "Count"], [[1, 1], [1, 1], [2, 1]]
    )
    right = book(tmp_path / "inv.xlsx", "Invoices", ["Id"], [[1], [2]])
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "Shipments"},
        {"table_name": "Invoices"},
        "Id",
        "Id",
        "Count",
    )
    assert result["measure"]["after"] == "3" and not result["requires_decision"]
    assert result["left_duplicate_key_count"] == 1


@pytest.mark.parametrize(
    "key,right_key,matches",
    [
        (True, 1, 0),
        ("1", 1, 0),
        ("a", "A", 0),
        (None, None, 0),
        (" ", " ", 0),
        (1, 1.0, 1),
    ],
)
def test_join_typed_blank_and_exact_keys(tmp_path, key, right_key, matches):
    left = book(tmp_path / "l.xlsx", "L", ["Id", "Amount"], [[key, 1]])
    right = book(tmp_path / "r.xlsx", "R", ["Id"], [[right_key]])
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "L"},
        {"table_name": "R"},
        "Id",
        "Id",
        "Amount",
        "inner",
    )
    assert result["projected_rows"] == matches


def test_join_negative_measures_and_nulls(tmp_path):
    left = book(tmp_path / "l.xlsx", "L", ["Id", "Amount"], [[1, -3], [2, None]])
    right = book(tmp_path / "r.xlsx", "R", ["Id"], [[1], [1]])
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "L"},
        {"table_name": "R"},
        "Id",
        "Id",
        "Amount",
    )
    assert result["measure"]["extra_signed"] == "-3"
    assert result["measure"]["extra_absolute"] == "3"
    assert result["null_measure_rows"] == 1


def test_rule_probe_finds_refund_absent_from_real_data(source, recipe):
    other = deepcopy(recipe)
    other["filters"] = [{"field": "Amount", "op": "gte", "value": 0}]
    before = state(source)
    result = probe_recipe_rules(str(source), recipe, other)
    assert not result["real_results_differ"]
    witness = result["witness"]
    assert witness["kind"] == "hypothetical_single_row"
    assert witness["record"]["Amount"] < 0
    assert witness["left_rows"] != witness["right_rows"]
    assert not result["equivalence_proven"] and state(source) == before


def test_rule_probe_identical_is_not_equivalence_proof(source, recipe):
    result = probe_recipe_rules(str(source), recipe, recipe)
    assert result["status"] == "no_difference_found"
    assert result["witness"] is None and not result["equivalence_proven"]
    assert result["synthetic_cases_tested"] <= 128


@pytest.mark.parametrize("op,value", [("contains", 1), ("eq", "100"), ("gte", True)])
def test_rule_probe_retains_recipe_filter_type_guards(source, recipe, op, value):
    right = deepcopy(recipe)
    right["filters"] = [{"field": "Amount", "op": op, "value": value}]
    before = state(source)
    with pytest.raises(ValidationError):
        probe_recipe_rules(str(source), recipe, right)
    assert state(source) == before


def test_rule_probe_observed_witness_and_no_definition_mutation(source, recipe):
    other = deepcopy(recipe)
    other["filters"] = [{"field": "Region", "op": "eq", "value": "South"}]
    original = deepcopy(other)
    result = probe_recipe_rules(str(source), recipe, other)
    assert result["real_results_differ"] and result["witness"]["source_row"] == 2
    assert result["witness"]["kind"] == "observed_single_row" and other == original


@pytest.mark.parametrize("change", ["source", "metrics", "checks"])
def test_rule_probe_rejects_unsupported_definitions(source, recipe, change):
    other = deepcopy(recipe)
    other[change] = {"table_name": "Other"} if change == "source" else []
    if change == "checks":
        other[change] = [{"type": "unique_key", "fields": ["Id"]}]
    with pytest.raises(ValidationError):
        probe_recipe_rules(str(source), recipe, other)


def test_receipt_unchanged_and_source_refresh(source, recipe):
    receipt = receipt_for(source, recipe)
    before = state(source)
    result = check_report_receipt(str(source), receipt)
    assert result["status"] == "current"
    assert not any(
        result[k] for k in ("source_changed", "output_changed", "result_changed")
    )
    assert state(source) == before
    edit(source, lambda wb: setattr(wb["Data"]["C2"], "value", 125))
    before = state(source)
    result = check_report_receipt(str(source), receipt)
    assert (
        result["status"] == "refresh_needed"
        and result["source_changed"]
        and result["result_changed"]
    )
    assert not result["output_changed"] and state(source) == before


@pytest.mark.parametrize("source_also_changes", [False, True])
def test_receipt_manual_output_conflict(source, recipe, source_also_changes):
    receipt = receipt_for(source, recipe)

    def mutate(wb):
        wb["Report"]["E5"] = 109
        if source_also_changes:
            wb["Data"]["C2"] = 125

    edit(source, mutate)
    before = state(source)
    result = check_report_receipt(str(source), receipt)
    assert result["status"] == "conflict" and result["output_changed"]
    assert result["source_changed"] == source_also_changes
    assert result["comparison_sample"] == [
        {
            "group": {"Region": "North"},
            "previous": {"Revenue": 100},
            "computed": {"Revenue": 125 if source_also_changes else 100},
            "stored": {"Revenue": 109},
        }
    ]
    assert state(source) == before


def test_receipt_source_changes_same_aggregate_needs_no_write(source, recipe):
    recipe["group_by"] = []
    receipt = receipt_for(source, recipe)

    def mutate(wb):
        wb["Data"]["C2"], wb["Data"]["C3"] = 110, 110

    edit(source, mutate)
    result = check_report_receipt(str(source), receipt)
    assert result["source_changed"] and not result["result_changed"]
    assert result["status"] == "current"


def test_receipt_new_and_removed_groups_are_not_reported_as_zero(source, recipe):
    receipt = receipt_for(source, recipe)
    edit(source, lambda wb: setattr(wb["Data"]["B2"], "value", "West"))
    result = check_report_receipt(str(source), receipt)
    assert result["status"] == "refresh_needed"
    examples = {r["group"]["Region"]: r for r in result["comparison_sample"]}
    assert examples["West"]["previous"] is None
    assert examples["West"]["computed"] == {"Revenue": 100}
    assert examples["North"]["computed"] is None
    assert examples["North"]["previous"] == {"Revenue": 100}


def test_receipt_keyed_output_reorder_is_not_a_metric_conflict(source, recipe):
    receipt = receipt_for(source, recipe)

    def swap(wb):
        ws = wb["Report"]
        a, b = [c.value for c in ws[5][3:5]], [c.value for c in ws[6][3:5]]
        for col, (x, y) in enumerate(zip(a, b), 4):
            ws.cell(5, col).value, ws.cell(6, col).value = y, x

    edit(source, swap)
    result = check_report_receipt(str(source), receipt)
    assert result["output_changed"] and result["stored_result_matches_current"]
    assert result["status"] == "current"


@pytest.mark.parametrize(
    "coordinate,value",
    [("D4", "Changed"), ("F10", "stale"), ("E5", "=100"), ("D6", "North")],
)
def test_receipt_headers_padding_formulas_duplicate_groups_conflict(
    source, recipe, coordinate, value
):
    receipt = receipt_for(source, recipe)
    edit(source, lambda wb: setattr(wb["Report"][coordinate], "value", value))
    assert check_report_receipt(str(source), receipt)["status"] == "conflict"


def test_receipt_budget_failure_happens_before_any_write(source, recipe, monkeypatch):
    monkeypatch.setattr(report_receipts, "MAX_RECEIPT_CHARS", 20)
    before = state(source)
    with pytest.raises(ResponseTooLargeError):
        run_workbook_recipe(str(source), recipe, include_report_receipt=True)
    assert state(source) == before


def test_receipt_tampering_and_unsupported_version_fail(source, recipe):
    receipt = receipt_for(source, recipe)
    for key, value in [("version", 2), ("rows", []), ("receipt_sha256", "x")]:
        bad = {**receipt, key: value}
        before = state(source)
        with pytest.raises(ValidationError):
            check_report_receipt(str(source), bad)
        assert state(source) == before


def test_receipt_invalid_source_does_not_claim_current(source, recipe):
    receipt = receipt_for(source, recipe)
    edit(source, lambda wb: setattr(wb["Data"]["C2"], "value", "=100"))
    with pytest.raises(ValidationError):
        check_report_receipt(str(source), receipt)


def test_reporting_preflight_real_stdio(source, recipe, tmp_path):
    root = Path(__file__).resolve().parents[1]

    async def run():
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "excel_mcp", "stdio"],
            cwd=root,
            env={"PYTHONPATH": str(root / "src")},
        )
        with (tmp_path / "stderr.log").open("w") as log:
            async with stdio_client(params, errlog=log) as (read, write):
                async with ClientSession(read, write) as client:
                    await client.initialize()
                    inventory = (await client.list_tools()).tools
                    expected = {
                        t["name"]
                        for t in json.loads((root / "manifest.json").read_text())[
                            "tools"
                        ]
                    }
                    assert {t.name for t in inventory} == expected
                    for name in (
                        "rehearse_table_join",
                        "probe_recipe_rules",
                        "check_report_receipt",
                    ):
                        tool = next(t for t in inventory if t.name == name)
                        assert tool.annotations.readOnlyHint

                    async def call(name, args):
                        result = await client.call_tool(name, args)
                        assert not result.isError
                        payload = json.loads(result.content[0].text)
                        assert payload["ok"], payload
                        assert len(result.content[0].text) < 50_000
                        return payload["data"]

                    before = state(source)
                    assert (
                        await call(
                            "probe_recipe_rules",
                            {
                                "filepath": str(source),
                                "left_recipe": recipe,
                                "right_recipe": recipe,
                            },
                        )
                    )["status"] == "no_difference_found"
                    joined = await call(
                        "rehearse_table_join",
                        {
                            "left_filepath": str(source),
                            "right_filepath": str(source),
                            "left_source": {"table_name": "Sales"},
                            "right_source": {"table_name": "Sales"},
                            "left_key": "Id",
                            "right_key": "Id",
                            "measure_field": "Amount",
                        },
                    )
                    assert not joined["requires_decision"] and state(source) == before
                    args = {
                        "filepath": str(source),
                        "recipe": recipe,
                        "include_report_receipt": True,
                    }
                    preview = await call("run_workbook_recipe", args)
                    result = await call(
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
                    saved = state(source)
                    checked = await call(
                        "check_report_receipt",
                        {
                            "filepath": str(source),
                            "report_receipt": result["report_receipt"],
                        },
                    )
                    assert checked["status"] == "current" and state(source) == saved
                    assert (
                        checked["source_sha256"] == hashlib.sha256(saved[0]).hexdigest()
                    )

    asyncio.run(run())


def test_receipt_contract_rebind_after_column_reorder(source, recipe):
    contract = {
        "version": 1,
        "name": "Sales",
        "source": {"table_name": "Sales"},
        "row_grain": "one sale",
        "key_fields": ["id"],
        "fields": [
            {"name": "id", "headers": ["Id"], "type": "integer"},
            {"name": "region", "headers": ["Region"], "type": "string"},
            {"name": "amount", "headers": ["Amount"], "type": "number"},
        ],
    }
    recipe["group_by"] = ["region"]
    recipe["metrics"][0]["field"] = "amount"
    receipt = receipt_for(source, recipe, workbook_contract=contract)
    assert check_report_receipt(str(source), receipt)["status"] == "current"

    def swap(wb):
        ws = wb["Data"]
        for row in ws.iter_rows(min_row=1, max_row=3, max_col=3):
            row[0].value, row[2].value = row[2].value, row[0].value
        table = ws.tables["Sales"]
        table.tableColumns[0].name, table.tableColumns[2].name = "Amount", "Id"

    edit(source, swap)
    assert check_report_receipt(str(source), receipt)["status"] == "current"
    edit(source, lambda wb: setattr(wb["Data"]["C3"], "value", 1))
    with pytest.raises(Exception, match="contract"):
        check_report_receipt(str(source), receipt)


def test_failed_commit_cannot_return_receipt_or_mutate(source, recipe, monkeypatch):
    import excel_mcp.changeset as changeset

    before = state(source)
    preview = run_workbook_recipe(str(source), recipe, include_report_receipt=True)

    def fail(*args, **kwargs):
        raise ValidationError("simulated transaction refusal")

    monkeypatch.setattr(changeset, "apply_workbook_changeset", fail)
    with pytest.raises(ValidationError, match="transaction refusal"):
        run_workbook_recipe(
            str(source),
            recipe,
            mode="commit",
            include_report_receipt=True,
            recipe_token=preview["recipe_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert state(source) == before


def test_receipt_capacity_budget_stops_commit_before_persistence(
    source, recipe, monkeypatch
):
    preview = run_workbook_recipe(str(source), recipe, include_report_receipt=True)
    before = state(source)
    monkeypatch.setattr(report_receipts, "MAX_RECEIPT_CHARS", 20)
    with pytest.raises(ResponseTooLargeError):
        run_workbook_recipe(
            str(source),
            recipe,
            mode="commit",
            include_report_receipt=True,
            recipe_token=preview["recipe_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert state(source) == before


def test_join_totals_excluded_and_same_file_capture(source):
    def totals(wb):
        ws = wb["Data"]
        ws.append([None, "Total", "=SUM(C2:C3)"])
        ws.tables["Sales"].ref = "A1:C4"
        ws.tables["Sales"].totalsRowCount = 1

    edit(source, totals)
    before = state(source)
    result = rehearse_table_join(
        str(source),
        str(source),
        {"table_name": "Sales"},
        {"table_name": "Sales"},
        "Id",
        "Id",
        "Amount",
    )
    assert result["measure"]["after"] == "220" and result["left_rows"] == 2
    assert result["left_source"]["totals_rows_excluded"] == 1
    assert result["left_source_sha256"] == result["right_source_sha256"]
    assert state(source) == before


@pytest.mark.parametrize("value", ["=100", "#REF!", True, "100"])
def test_join_invalid_measure_never_coerces(source, value):
    edit(source, lambda wb: setattr(wb["Data"]["C2"], "value", value))
    before = state(source)
    with pytest.raises(ValidationError):
        rehearse_table_join(
            str(source),
            str(source),
            {"table_name": "Sales"},
            {"table_name": "Sales"},
            "Id",
            "Id",
            "Amount",
        )
    assert state(source) == before


@pytest.mark.parametrize("limit", [0, 6, True, "3"])
def test_all_evidence_tools_reject_invalid_sample_limits(source, recipe, limit):
    for call in (
        lambda: rehearse_table_join(
            str(source),
            str(source),
            {"table_name": "Sales"},
            {"table_name": "Sales"},
            "Id",
            "Id",
            "Amount",
            sample_limit=limit,
        ),
        lambda: probe_recipe_rules(str(source), recipe, recipe, limit),
        lambda: check_report_receipt(str(source), {}, limit),
    ):
        with pytest.raises(ValidationError):
            call()


def test_rule_group_identifiers_are_not_compared_with_metric_tolerance():
    from excel_mcp.rule_probe import _different

    assert _different(
        {"headers": ["Id", "Sum"], "rows": [[1e-11, 10]]},
        {"headers": ["Id", "Sum"], "rows": [[2e-11, 10]]},
        1,
    )


def test_receipt_all_null_result_is_not_confused_with_padding(source, recipe):
    edit(
        source,
        lambda wb: [setattr(wb["Data"].cell(r, 3), "value", None) for r in (2, 3)],
    )
    recipe["group_by"] = []
    recipe["metrics"][0]["op"] = "avg"
    before = state(source)
    assert run_workbook_recipe(str(source), recipe)["ready_to_commit"]
    with pytest.raises(ValidationError, match="entirely null"):
        run_workbook_recipe(str(source), recipe, include_report_receipt=True)
    assert state(source) == before


@pytest.mark.parametrize("kind", ["date", "duration", "array_formula", "boolean"])
def test_receipt_nonliteral_or_typed_output_changes_are_conflicts(source, recipe, kind):
    from datetime import datetime, timedelta
    from openpyxl.worksheet.formula import ArrayFormula

    receipt = receipt_for(source, recipe)
    values = {
        "date": datetime(2026, 11, 27),
        "duration": timedelta(days=1),
        "array_formula": ArrayFormula(ref="E5:E5", text="=100"),
        "boolean": True,
    }
    edit(source, lambda wb: setattr(wb["Report"]["E5"], "value", values[kind]))
    before = state(source)
    checked = check_report_receipt(str(source), receipt)
    assert (
        checked["status"] == "conflict" and not checked["stored_result_matches_current"]
    )
    assert state(source) == before


def test_receipt_post_replace_failure_rolls_back_without_issuing_receipt(
    source, recipe, monkeypatch
):
    from excel_mcp import workbook
    from excel_mcp.exceptions import WorkbookError

    preview = run_workbook_recipe(str(source), recipe, include_report_receipt=True)
    before = source.read_bytes()
    verify = workbook._verify_saved_workbook

    def fail(path):
        if Path(path).resolve() == source.resolve() and source.read_bytes() != before:
            raise WorkbookError("injected verification failure")
        return verify(path)

    monkeypatch.setattr(workbook, "_verify_saved_workbook", fail)
    with pytest.raises(WorkbookError, match="restored"):
        run_workbook_recipe(
            str(source),
            recipe,
            mode="commit",
            include_report_receipt=True,
            recipe_token=preview["recipe_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert source.read_bytes() == before


def test_failed_recipe_checks_never_expose_prospective_receipt(source, recipe):
    from excel_mcp.exceptions import PreconditionFailedError

    recipe["checks"] = [{"type": "row_count", "min": 99, "max": 100}]
    before = state(source)
    preview = run_workbook_recipe(str(source), recipe, include_report_receipt=True)
    assert not preview["ready_to_commit"] and "report_receipt" not in preview
    with pytest.raises(PreconditionFailedError) as exc:
        run_workbook_recipe(
            str(source),
            recipe,
            mode="commit",
            include_report_receipt=True,
            recipe_token=preview["recipe_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert "report_receipt" not in exc.value.details
    assert state(source) == before


def test_large_projected_join_is_counted_without_materializing_rows(tmp_path):
    left = book(tmp_path / "left.xlsx", "L", ["Id", "Amount"], [[1, 1]] * 300)
    right = book(tmp_path / "right.xlsx", "R", ["Id"], [[1]] * 300)
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "L"},
        {"table_name": "R"},
        "Id",
        "Id",
        "Amount",
        sample_limit=2,
    )
    assert result["projected_rows"] == 90_000
    assert result["measure"]["extra_signed"] == "89700"
    assert result["samples_truncated"]["multiplied"]
    assert len(result["samples"]["multiplied"]) == 2
    assert result["samples"]["multiplied"][0]["right_rows"] == [2, 3]
    assert len(json.dumps(result)) < 5000


@pytest.mark.parametrize("seed", range(12))
def test_join_independent_decimal_oracle_across_duplicates_refunds_and_nulls(
    tmp_path, seed
):
    from decimal import Decimal
    import random

    rng = random.Random(seed)
    left_rows = [
        [rng.choice([None, 1, 2, 3]), rng.choice([None, -1.25, 0, 3.5])]
        for _ in range(8)
    ]
    right_rows = [[rng.choice([None, 1, 2, 3])] for _ in range(6)]
    left = book(tmp_path / "l.xlsx", "L", ["Id", "Amount"], left_rows)
    right = book(tmp_path / "r.xlsx", "R", ["Id"], right_rows)
    for join_type in ("left", "inner"):
        output = []
        extra = lost = Decimal(0)
        for key, amount in left_rows:
            matches = [r for r in right_rows if key is not None and key == r[0]]
            copies = len(matches) if join_type == "inner" else max(1, len(matches))
            value = Decimal(str(amount)) if amount is not None else Decimal(0)
            output.extend([value] * copies)
            if copies > 1:
                extra += abs(value) * (copies - 1)
            if copies == 0:
                lost += abs(value)
        result = rehearse_table_join(
            str(left),
            str(right),
            {"table_name": "L"},
            {"table_name": "R"},
            "Id",
            "Id",
            "Amount",
            join_type,
        )
        assert result["projected_rows"] == len(output)
        assert Decimal(result["measure"]["after"]) == sum(output)
        assert Decimal(result["measure"]["extra_absolute"]) == extra
        assert Decimal(result["measure"]["lost_absolute"]) == lost
