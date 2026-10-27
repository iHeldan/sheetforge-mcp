"""Read-only, identity-aware explanations of bounded recipe result changes."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Context, Decimal, localcontext
from fractions import Fraction
from typing import Any

from . import changeset, query, recipes
from .exceptions import PreconditionFailedError, ValidationError

CATEGORIES = ("added", "removed", "updated", "unchanged")
ZERO = Fraction(0)
MAX_METRICS = 8
MAX_RESULT_CELLS = 10_000
REL_TOL = Fraction(str(recipes.NUMERIC_REL_TOL))
ABS_TOL = Fraction(str(recipes.NUMERIC_ABS_TOL))


def _contract(recipe: dict, key_fields: list[str]) -> tuple[dict, list[str], list[str]]:
    normalized = recipes._normalize_recipe(recipe)
    keys = recipes._field_list(key_fields, "key_fields")
    if not keys:
        raise ValidationError(
            "key_fields must contain at least one exact source header"
        )
    if len(normalized["metrics"]) > MAX_METRICS:
        raise ValidationError(f"Explanations v1 supports at most {MAX_METRICS} metrics")
    for item in normalized["metrics"]:
        if item["op"] not in {"sum", "count"}:
            raise ValidationError("Explanations v1 supports sum and count metrics only")
        if item["as"] != item["as"].strip():
            raise ValidationError("Metric aliases must not have surrounding whitespace")
    for spec in normalized["filters"]:
        op = spec["op"].strip().lower()
        spec["op"] = query.FILTER_OPERATOR_ALIASES.get(op, op)
    fields = list(dict.fromkeys([*recipes._source_fields(normalized), *keys]))
    return normalized, keys, fields


def _compatible(before: dict, after: dict, fields: list[str]) -> None:
    for field in fields:
        kinds = []
        for source in (before, after):
            index = source["headers"].index(field)
            kinds.append(
                {
                    recipes._kind(row[index])
                    for row in source["rows"]
                    if row[index] is not None
                }
            )
        if kinds[0] and kinds[1] and kinds[0] != kinds[1]:
            raise PreconditionFailedError(
                f"Referenced field '{field}' changed value type between sources",
                code="incompatible_field_types",
                details={"field": field},
            )


def _identity(value: Any) -> tuple:
    kind = recipes._kind(value)
    return kind, Fraction(str(value)) if kind == "number" else value


def _number(value: Any) -> Fraction:
    if value is None:
        return ZERO
    if recipes._kind(value) != "number":
        raise ValidationError("sum metrics require numeric or null source values")
    return Fraction(str(value))


def _text(value: Fraction) -> str:
    # All denominators originate from decimal literals, so this division
    # terminates. Explicit precision avoids the caller's ambient Decimal context.
    precision = abs(value.numerator).bit_length() + value.denominator.bit_length() + 5
    with localcontext(Context(prec=max(28, precision))):
        text = format(Decimal(value.numerator) / Decimal(value.denominator), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def _matches(values: dict, filters: list) -> bool:
    """Evaluate the narrow recipe filter contract independently of query execution."""

    def fold(value, case_sensitive):
        return (
            value.casefold() if isinstance(value, str) and not case_sensitive else value
        )

    for spec in filters:
        actual, op = values[spec["field"]], spec["op"]
        expected = spec.get("value")
        case = spec.get("case_sensitive", False)
        left, right = fold(actual, case), fold(expected, case)
        if op in {"is_blank", "not_blank"}:
            matched = actual is None or isinstance(actual, str) and not actual.strip()
            matched = matched if op == "is_blank" else not matched
        elif op in {"in", "not_in"}:
            choices = spec.get("values", [expected])
            matched = left in [fold(item, case) for item in choices]
            matched = matched if op == "in" else not matched
        elif op in {"eq", "neq"}:
            matched = left == right
            matched = matched if op == "eq" else not matched
        elif actual is None or expected is None:
            matched = False
        elif op == "contains":
            matched = right in left
        elif op == "starts_with":
            matched = left.startswith(right)
        elif op == "ends_with":
            matched = left.endswith(right)
        elif op == "gt":
            matched = actual > expected
        elif op == "gte":
            matched = actual >= expected
        elif op == "lt":
            matched = actual < expected
        elif op == "lte":
            matched = actual <= expected
        else:
            raise ValidationError(f"Unsupported recipe filter operator: {op}")
        if not matched:
            return False
    return True


@dataclass(frozen=True)
class _Record:
    row: int
    values: dict
    group: tuple
    included: bool
    contributions: tuple[Fraction, ...]


def _index(
    source: dict, recipe: dict, key_fields: list[str], fields: list[str]
) -> dict:
    indexes = {field: source["headers"].index(field) for field in fields}
    # Reuse syntax validation, not query's filter matcher or aggregation arithmetic.
    query._normalize_filters(
        recipe["filters"],
        source["headers"],
        query._build_schema(source["headers"], source["rows"]),
    )
    records = {}
    for offset, row in enumerate(source["rows"]):
        values = {field: row[index] for field, index in indexes.items()}
        key_values = [values[field] for field in key_fields]
        if any(
            value is None or isinstance(value, str) and not value.strip()
            for value in key_values
        ):
            raise PreconditionFailedError(
                "Source keys must be nonblank on every row", code="blank_source_key"
            )
        key = tuple(_identity(value) for value in key_values)
        if key in records:
            raise PreconditionFailedError(
                "Source keys must be unique even on excluded rows",
                code="duplicate_source_key",
            )
        included = _matches(values, recipe["filters"])
        amounts = tuple(
            Fraction(1) if item["op"] == "count" else _number(values[item["field"]])
            for item in recipe["metrics"]
        )
        records[key] = _Record(
            row=source["data_start_row"] + offset,
            values=values,
            group=tuple(_identity(values[field]) for field in recipe["group_by"]),
            included=included,
            contributions=amounts if included else (ZERO,) * len(amounts),
        )
    return records


def _groups(records: dict, recipe: dict) -> dict:
    groups = {} if recipe["group_by"] else {(): [ZERO] * len(recipe["metrics"])}
    for record in records.values():
        if record.included:
            total = groups.setdefault(record.group, [ZERO] * len(recipe["metrics"]))
            for index, value in enumerate(record.contributions):
                total[index] += value
    return groups


def _near(actual: Fraction, expected: Fraction, *, exact: bool) -> bool:
    return (
        actual == expected
        or not exact
        and abs(actual - expected)
        <= max(ABS_TOL, REL_TOL * max(abs(actual), abs(expected)))
    )


def _totals(groups: dict, size: int) -> list:
    return [
        sum((values[index] for values in groups.values()), ZERO)
        for index in range(size)
    ]


def _reconcile(
    source: dict, records: dict, oracle: dict, recipe: dict
) -> tuple[dict, dict]:
    result = recipes._aggregate(source, recipe)
    width = len(recipe["group_by"]) + len(recipe["metrics"])
    if (len(result["rows"]) + 1) * width > MAX_RESULT_CELLS:
        raise ValidationError(f"Recomputed result exceeds {MAX_RESULT_CELLS} cells")
    expected_headers = recipe["group_by"] + [item["as"] for item in recipe["metrics"]]
    if result["headers"] != expected_headers:
        raise PreconditionFailedError(
            "Recipe result headers disagree with the contract",
            code="independent_reconciliation_failed",
        )
    observed = {}
    group_size = len(recipe["group_by"])
    for row in result["rows"]:
        if len(row) != width or any(
            recipes._kind(value) != "number" for value in row[group_size:]
        ):
            raise PreconditionFailedError(
                "Recipe result has invalid metric values",
                code="independent_reconciliation_failed",
            )
        key = tuple(_identity(value) for value in row[:group_size])
        if key in observed:
            raise PreconditionFailedError(
                "Recipe result contains duplicate groups",
                code="independent_reconciliation_failed",
            )
        observed[key] = [_number(value) for value in row[group_size:]]
    if observed.keys() != oracle.keys() or result["matched_rows"] != sum(
        record.included for record in records.values()
    ):
        raise PreconditionFailedError(
            "Recipe group or filter coverage disagrees with the independent evaluator",
            code="independent_reconciliation_failed",
        )
    exact = [
        item["op"] == "count"
        or all(
            record.values[item["field"]] is None
            or type(record.values[item["field"]]) is int
            for record in records.values()
        )
        for item in recipe["metrics"]
    ]
    for actual, expected in [
        *((observed[key], oracle[key]) for key in oracle),
        (_totals(observed, len(exact)), _totals(oracle, len(exact))),
    ]:
        for index, (a, b) in enumerate(zip(actual, expected)):
            if not _near(a, b, exact=exact[index]):
                raise PreconditionFailedError(
                    "Recipe result does not reconcile with independent exact arithmetic",
                    code="independent_reconciliation_failed",
                    details={
                        "metric": recipe["metrics"][index]["as"],
                        "recipe_value": _text(a),
                        "independent_value": _text(b),
                    },
                )
    verification = recipes._verify_recipe(source, recipe, result, sample_limit=1)
    if not verification["passed"]:
        raise PreconditionFailedError(
            "Recipe checks or invariance probes failed on a source version",
            code="recipe_verification_failed",
            details=changeset._bound_response_value(
                {
                    "checks": verification["checks"],
                    "probes": verification["invariance_probes"],
                }
            ),
        )
    return observed, {
        "checks": verification["checks"],
        "invariance_probes": verification["invariance_probes"],
    }


def _delta(
    before: _Record | None, after: _Record | None, size: int, group: tuple | None
) -> list:
    old = (
        before.contributions
        if before and (group is None or before.group == group)
        else (ZERO,) * size
    )
    new = (
        after.contributions
        if after and (group is None or after.group == group)
        else (ZERO,) * size
    )
    return [right - left for left, right in zip(old, new)]


def _pairs(before: dict, after: dict, fields: list[str]):
    # Preserve source ordering for stable tie-breaking without sorting raw values
    # that might be null, text, or numeric composite keys.
    for key in dict.fromkeys([*before, *after]):
        old, new = before.get(key), after.get(key)
        changed = (
            []
            if old is None or new is None
            else [
                field
                for field in fields
                if _identity(old.values[field]) != _identity(new.values[field])
            ]
        )
        category = (
            "added"
            if old is None
            else "removed"
            if new is None
            else "updated"
            if changed
            else "unchanged"
        )
        yield category, old, new, changed


def _breakdown(
    before: dict, after: dict, fields: list[str], size: int, group: tuple | None = None
) -> dict:
    totals = {category: [ZERO] * size for category in CATEGORIES}
    counts = dict.fromkeys(CATEGORIES, 0)
    for category, old, new, _ in _pairs(before, after, fields):
        if group is not None and not any(
            record and record.included and record.group == group
            for record in (old, new)
        ):
            continue
        counts[category] += 1
        for index, change in enumerate(_delta(old, new, size, group)):
            totals[category][index] += change
    return {"counts": counts, "totals": totals}
