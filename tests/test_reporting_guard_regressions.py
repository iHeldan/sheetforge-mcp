"""Regression controls for exact keys, Recipe guards and group comparisons."""

from copy import deepcopy
from fractions import Fraction
import hashlib
import io
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
import pytest

from excel_mcp import explanations, rebase, recipes, reporting_evidence, rule_probe
from excel_mcp.exceptions import PreconditionFailedError, ValidationError
from excel_mcp.join_rehearsal import rehearse_table_join
from excel_mcp.report_receipts import check_report_receipt
from excel_mcp.values import value_key
from tests.test_reporting_preflight import book, receipt_for, state
from tests.test_preflight_audit_regressions import legacy


def numeric_xml(path, entries):
    """Synthetic XLSX literals with exact numeric spellings not emitted by openpyxl."""
    payload = io.BytesIO()
    with ZipFile(path) as source, ZipFile(payload, "w") as target:
        for member in source.infolist():
            data = source.read(member)
            if member.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                for coordinate, text in entries.items():
                    cell = root.find(f".//{{{ns}}}c[@r='{coordinate}']")
                    cell.set("t", "n")
                    for child in list(cell):
                        cell.remove(child)
                    ET.SubElement(cell, f"{{{ns}}}v").text = text
                data = ET.tostring(root)
            target.writestr(member, data)
    path.write_bytes(payload.getvalue())


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


def drift_case(tmp_path, drift):
    base_row, current_row = [1, 10, "old"], [1, 10, "old"]
    column, index, expected = {
        "key": ("A2", 0, "missing_row"),
        "read_field": ("B2", 1, "stale_precondition"),
        "write_field": ("C2", 2, "write_conflict"),
    }[drift]
    current_row[index] = 1e23
    base = book(tmp_path / "base.xlsx", "Items", ["Id", "Premise", "Note"], [base_row])
    numeric_xml(base, {column: str(10**23)})
    current = book(
        tmp_path / "current.xlsx", "Items", ["Id", "Premise", "Note"], [current_row]
    )
    patch = {
        "version": 1,
        "source": {"table_name": "Items"},
        "key_fields": ["Id"],
        "read_fields": ["Premise"],
        "updates": [
            {"key": {"Id": 10**23 if drift == "key" else 1}, "values": {"Note": "new"}}
        ],
    }
    return base, current, patch, expected


@pytest.mark.parametrize("drift", ["key", "read_field", "write_field"])
def test_rebase_exact_numeric_drift_blocks_preview_and_commit(tmp_path, drift):
    base, current, patch, expected = drift_case(tmp_path, drift)
    before = state(base), state(current)
    args = {
        "base_filepath": str(base),
        "filepath": str(current),
        "patch": patch,
        "expected_base_sha256": hashlib.sha256(before[0][0]).hexdigest(),
    }
    preview = rebase.apply_table_patch(**args)
    assert preview["blocked"] and not preview["ready_to_commit"]
    assert preview["counts"][expected] == 1 and preview["counts"]["apply"] == 0
    with pytest.raises(PreconditionFailedError) as caught:
        rebase.apply_table_patch(
            **args,
            mode="commit",
            rebase_token=preview["rebase_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert caught.value.code == "rebase_conflict"
    assert (state(base), state(current)) == before
    assert {p.name for p in tmp_path.iterdir()} == {"base.xlsx", "current.xlsx"}


@pytest.mark.parametrize("drift", ["key", "read_field", "write_field"])
def test_rebase_exactly_equal_integer_float_values_allow_verified_commit(
    tmp_path, drift
):
    base, current, patch, _ = drift_case(tmp_path, drift)
    numeric_xml(
        base,
        {{"key": "A2", "read_field": "B2", "write_field": "C2"}[drift]: str(int(1e23))},
    )
    if drift == "key":
        patch["updates"][0]["key"]["Id"] = int(1e23)
    before_base, before_current = state(base), state(current)
    args = {
        "base_filepath": str(base),
        "filepath": str(current),
        "patch": patch,
        "expected_base_sha256": hashlib.sha256(before_base[0]).hexdigest(),
    }
    preview = rebase.apply_table_patch(**args)
    assert preview["ready_to_commit"] and preview["counts"]["apply"] == 1
    saved = rebase.apply_table_patch(
        **args,
        mode="commit",
        rebase_token=preview["rebase_token"],
        expected_workbook_sha256=preview["expected_workbook_sha256"],
    )
    assert saved["persisted"] and state(base) == before_base
    assert Path(saved["receipt"]["snapshot"]["path"]).read_bytes() == before_current[0]
    wb = load_workbook(current)
    try:
        assert wb["Data"]["C2"].value == "new"
    finally:
        wb.close()


@pytest.mark.parametrize("integer,matched", [(10**23, False), (int(1e23), True)])
def test_join_matches_exact_executed_numeric_keys(tmp_path, integer, matched):
    left = book(tmp_path / "left.xlsx", "Invoices", ["Key", "Amount"], [[1, 100]])
    numeric_xml(left, {"A2": str(integer)})
    right = book(tmp_path / "right.xlsx", "Shipments", ["Key"], [[1e23]])
    before = state(left), state(right)
    result = rehearse_table_join(
        str(left),
        str(right),
        {"table_name": "Invoices"},
        {"table_name": "Shipments"},
        "Key",
        "Key",
        "Amount",
        "inner",
    )
    assert result["projected_rows"] == int(matched)
    assert result["unmatched_left_rows"] == int(not matched)
    assert result["requires_decision"] is not matched
    assert result["measure"]["after"] == ("100" if matched else "0")
    assert (state(left), state(right)) == before


def test_probe_distinguishes_unequal_numeric_groups(tmp_path, rule):
    path = book(
        tmp_path / "groups.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, 1, 100], [2, 1e23, 100]],
    )
    numeric_xml(path, {"B2": str(10**23)})
    left, right = deepcopy(rule), deepcopy(rule)
    left["filters"] = [{"field": "Id", "op": "eq", "value": 1}]
    right["filters"] = [{"field": "Id", "op": "eq", "value": 2}]
    before = state(path)
    assert rule_probe.probe_recipe_rules(str(path), left, right)["real_results_differ"]
    assert state(path) == before


@pytest.mark.parametrize(
    "helper", [rebase._identity, explanations._identity, reporting_evidence.identity]
)
def test_numeric_identities_agree_with_query_and_preserve_decimal_arithmetic(helper):
    values = [
        None,
        True,
        False,
        "1",
        0,
        -0.0,
        1,
        1.0,
        0.1,
        1e-300,
        5e-324,
        2**53,
        2**53 + 1,
        10**20,
        1e20,
        10**23,
        1e23,
        int(1e23),
    ]
    for left in values:
        for right in values:
            assert (helper(left) == helper(right)) == (
                value_key(left) == value_key(right)
            )
    assert explanations._number(0.1) == Fraction(1, 10)
    assert reporting_evidence.text(reporting_evidence.number(0.1)) == "0.1"


def mixed_group_book(tmp_path):
    path = book(
        tmp_path / "order.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, 1, 100], [2, 10, 200], [3, 1, 100], [4, 10, 200]],
    )
    numeric_xml(path, {"B4": "1.0", "B5": "10.0"})
    return path


def test_probe_compares_keyed_aggregates_not_group_serialization_order(tmp_path, rule):
    path = mixed_group_book(tmp_path)
    left, right = deepcopy(rule), deepcopy(rule)
    left["filters"] = [{"field": "Id", "op": "lte", "value": 2}]
    right["filters"] = [{"field": "Id", "op": "gt", "value": 2}]
    before = state(path)
    result = rule_probe.probe_recipe_rules(str(path), left, right)
    assert not result["real_results_differ"]
    assert result["real_results"][0]["rows_sample"] == [[10, 200], [1, 100]]
    assert result["real_results"][1]["rows_sample"] == [[1.0, 100], [10.0, 200]]
    assert state(path) == before


def test_recipe_group_invariance_accepts_numeric_equivalence_and_preserves_receipt(
    tmp_path, rule
):
    path = mixed_group_book(tmp_path)
    receipt = receipt_for(path, rule)
    before = state(path)
    checked = check_report_receipt(str(path), receipt)
    assert checked["status"] == "current" and state(path) == before


@pytest.mark.parametrize(
    "op", ["contains", "CONTAINS", " Contains ", " STARTS_WITH ", "ENDS_WITH"]
)
def test_recipe_and_probe_reject_string_operators_on_numbers_for_all_spellings(
    tmp_path, rule, op
):
    path = book(
        tmp_path / "guard.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 12], [2, "North", 3]],
    )
    rule["filters"] = [{"field": "Amount", "op": op, "value": 1}]
    before = state(path)
    with pytest.raises(ValidationError, match="String filter requires a text field"):
        recipes.run_workbook_recipe(str(path), rule, include_report_receipt=True)
    with pytest.raises(ValidationError, match="String filter requires a text field"):
        rule_probe.probe_recipe_rules(str(path), rule, rule)
    assert state(path) == before


@pytest.mark.parametrize("op", ["is_blank", "IS_BLANK", " Not_Blank "])
@pytest.mark.parametrize("operand", ["value", "values"])
def test_recipe_blank_check_guards_apply_before_execution(tmp_path, rule, op, operand):
    path = book(
        tmp_path / "blank.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["filters"] = [
        {"field": "Amount", "op": op, operand: 5 if operand == "value" else [5]}
    ]
    before = state(path)
    with pytest.raises(ValidationError, match="Blank checks must not specify"):
        recipes.run_workbook_recipe(str(path), rule)
    assert state(path) == before


@pytest.mark.parametrize("op", ["CONTAINS", " Contains ", "STARTS_WITH", "ENDS_WITH"])
def test_valid_text_operator_spelling_remains_in_receipt_definition(tmp_path, rule, op):
    path = book(
        tmp_path / "text.xlsx", "Sales", ["Id", "Region", "Amount"], [[1, "North", 12]]
    )
    needle = "N" if "STARTS" in op else "h" if "ENDS" in op else "or"
    rule["filters"] = [{"field": "Region", "op": op, "value": needle}]
    original = deepcopy(rule)
    receipt = receipt_for(path, rule)
    assert receipt["recipe"]["filters"][0]["op"] == op and rule == original
    before = state(path)
    assert check_report_receipt(str(path), receipt)["status"] == "current"
    assert state(path) == before


@pytest.mark.parametrize("version", [1, 2])
def test_legacy_receipts_keep_valid_retained_operator_spelling(tmp_path, rule, version):
    path = book(
        tmp_path / "legacy.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 12]],
    )
    rule["filters"] = [{"field": "Region", "op": " Contains ", "value": "or"}]
    receipt = legacy(receipt_for(path, rule), path, version)
    before = state(path)
    assert check_report_receipt(str(path), receipt)["status"] == "current"
    assert (
        receipt["recipe"]["filters"][0]["op"] == " Contains " and state(path) == before
    )


@pytest.mark.parametrize("op", ["in", " IN ", "Not_In", "ne", " NE "])
def test_canonical_filter_guards_keep_valid_membership_and_alias_semantics(
    tmp_path, rule, op
):
    path = book(
        tmp_path / "membership.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", 12]],
    )
    is_membership = "in" in op.lower()
    operand = "North" if op.strip().lower() == "in" else "South"
    rule["filters"] = [
        {
            "field": "Region",
            "op": op,
            **({"values": [operand]} if is_membership else {"value": operand}),
        }
    ]
    before = state(path)
    preview = recipes.run_workbook_recipe(str(path), rule)
    assert preview["ready_to_commit"] and preview["rows_sample"] == [["North", 12]]
    assert state(path) == before


@pytest.mark.parametrize("op", ["is_blank", " IS_BLANK ", "Not_Blank"])
def test_blank_filters_without_operands_remain_valid(tmp_path, rule, op):
    path = book(
        tmp_path / "blank-ok.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [[1, "North", None]],
    )
    rule["filters"] = [{"field": "Amount", "op": op}]
    before = state(path)
    preview = recipes.run_workbook_recipe(str(path), rule)
    assert preview["ready_to_commit"]
    assert preview["rows_sample"] == (
        [["North", 0]] if op.strip().lower() == "is_blank" else []
    )
    assert state(path) == before


@pytest.mark.parametrize(
    "left,right,equal",
    [
        ([[1, 10], [10, 20]], [[10.0, 20], [1.0, 10]], True),
        ([[1, 10]], [[1, 10.0 + 1e-12]], True),
        ([[1, 10]], [[1, 11]], False),
        ([[1, 10]], [[1 + 1e-13, 10]], False),
        ([[True, 10]], [[1, 10]], False),
        ([[10**23, 10]], [[1e23, 10]], False),
        ([[int(1e23), 10]], [[1e23, 10]], True),
    ],
)
def test_group_comparisons_keep_exact_keys_and_metric_tolerances(left, right, equal):
    a, b = ({"headers": ["Group", "Value"], "rows": rows} for rows in (left, right))
    assert recipes._equal_grouped_result(a, b, 1) is equal
    assert rule_probe._different(a, b, 1) is not equal


@pytest.mark.parametrize(
    "rows,message",
    [
        ([[1, 10], [1.0, 20]], "duplicate group identities"),
        ([[1], [2]], "row width"),
    ],
)
def test_invalid_aggregate_shapes_do_not_certify_group_equivalence(rows, message):
    result = {"headers": ["Group", "Value"], "rows": rows}
    with pytest.raises(ValidationError, match=message):
        recipes._equal_grouped_result(result, result, 1)


@pytest.mark.parametrize("dimension", ["Id", "Region"])
def test_explanation_sibling_key_paths_keep_distinct_large_numbers(
    tmp_path, rule, dimension
):
    before = book(
        tmp_path / "before.xlsx",
        "Sales",
        ["Id", "Region", "Amount"],
        [
            [1, "North" if dimension == "Id" else 1, 10],
            [2, "North" if dimension == "Id" else 1e23, 20],
        ],
    )
    coordinate = "A2" if dimension == "Id" else "B2"
    numeric_xml(before, {coordinate: str(10**23)})
    if dimension == "Id":
        numeric_xml(before, {"A3": "1e23"})
    after = tmp_path / "after.xlsx"
    after.write_bytes(before.read_bytes())
    frozen = state(before), state(after)
    result = explanations.explain_recipe_delta(str(before), str(after), rule, ["Id"])
    assert (
        result["status"] == "verified" and result["changes"]["counts"]["unchanged"] == 2
    )
    assert result["metrics"][0]["delta"] == "0"
    assert (state(before), state(after)) == frozen
