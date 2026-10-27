from fractions import Fraction

import pytest

from excel_mcp.explanations import (
    _breakdown,
    _contract,
    _groups,
    _index,
    _reconcile,
    _text,
)
from excel_mcp.exceptions import PreconditionFailedError


@pytest.fixture
def recipe():
    return {
        "filters": [],
        "checks": [],
        "group_by": ["Region"],
        "metrics": [
            {"op": "sum", "field": "Amount", "as": "Revenue"},
            {"op": "count", "as": "Count"},
        ],
    }


def _source(rows):
    return {
        "headers": ["Id", "Region", "Amount"],
        "rows": rows,
        "data_start_row": 2,
        "sheet_name": "Data",
        "table_name": "Sales",
    }


def test_independent_reconciliation_catches_wrong_engine_result(recipe, monkeypatch):
    import excel_mcp.explanations as explanations

    source = _source([[1, "North", 10]])
    records = _index(source, recipe, ["Id"], ["Id", "Region", "Amount"])
    original = explanations.recipes._aggregate

    def incorrect(*args):
        result = original(*args)
        result["rows"][0][1] += 5
        return result

    monkeypatch.setattr(explanations.recipes, "_aggregate", incorrect)
    with pytest.raises(PreconditionFailedError) as error:
        _reconcile(source, records, _groups(records, recipe), recipe)
    assert error.value.code == "independent_reconciliation_failed"


def test_compensated_recipe_and_exact_decimal_comparison_agree(recipe):
    source = _source([[1, "North", 1e16], [2, "North", 1], [3, "North", -1e16]])
    records = _index(source, recipe, ["Id"], ["Id", "Region", "Amount"])
    observed, _ = _reconcile(source, records, _groups(records, recipe), recipe)
    assert next(iter(observed.values())) == [1, 3]


@pytest.mark.parametrize("op", ["avg", "count_distinct", "min", "max"])
def test_contract_rejects_nonadditive_metrics(recipe, op):
    from excel_mcp.exceptions import ValidationError

    recipe.update(
        version=1,
        source={"table_name": "Sales"},
        output={"sheet_name": "Report", "range_ref": "A1:C10"},
    )
    recipe["metrics"][0]["op"] = op
    with pytest.raises(ValidationError, match="sum and count"):
        _contract(recipe, ["Id"])


def test_exact_identity_decomposition(recipe):
    fields = ["Id", "Region", "Amount"]
    before = _index(
        _source([[1, "North", 60000], [2, "South", 35000], [3, "West", 5000]]),
        recipe,
        ["Id"],
        fields,
    )
    after = _index(
        _source([[2, "South", 32000], [1, "North", 60000], [4, "North", 20000]]),
        recipe,
        ["Id"],
        fields,
    )
    result = _breakdown(before, after, fields, 2)
    assert result["counts"] == {"added": 1, "removed": 1, "updated": 1, "unchanged": 1}
    assert result["totals"]["added"] == [20000, 1]
    assert result["totals"]["removed"] == [-5000, -1]
    assert result["totals"]["updated"] == [-3000, 0]
    assert sum(total[0] for total in result["totals"].values()) == 12000


def test_regrouping_cancels_globally_but_not_per_group(recipe):
    fields = ["Id", "Region", "Amount"]
    before = _index(_source([[1, "North", 0.1]]), recipe, ["Id"], fields)
    after = _index(_source([[1, "South", 0.1]]), recipe, ["Id"], fields)
    group = next(iter(_groups(before, recipe)))
    assert _breakdown(before, after, fields, 2)["totals"]["updated"] == [0, 0]
    assert _breakdown(before, after, fields, 2, group)["totals"]["updated"] == [
        -Fraction(1, 10),
        -1,
    ]


@pytest.mark.parametrize(
    "rows, code",
    [
        ([[1, "North", 1], [1, "South", 2]], "duplicate_source_key"),
        ([[None, "North", 1]], "blank_source_key"),
        ([["   ", "North", 1]], "blank_source_key"),
    ],
)
def test_keys_are_validated_even_when_all_rows_are_filtered_out(recipe, rows, code):
    recipe["filters"] = [{"field": "Amount", "op": "gt", "value": 99999}]
    with pytest.raises(PreconditionFailedError) as error:
        _index(_source(rows), recipe, ["Id"], ["Id", "Region", "Amount"])
    assert error.value.code == code


@pytest.mark.parametrize(
    "value", ["0.000000000000000000000001", "1e308", "-1e-308", "0.1", "-12.50"]
)
def test_exact_output_does_not_depend_on_decimal_context(value):
    from decimal import Decimal, localcontext

    with localcontext() as context:
        context.prec = 3
        assert Decimal(_text(Fraction(value))) == Decimal(value)
