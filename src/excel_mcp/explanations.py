"""Read-only, identity-aware explanations of bounded recipe result changes."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Context, Decimal, localcontext
from fractions import Fraction
import hashlib
import json
from pathlib import Path
import re
from typing import Any

from . import changeset, query, recipes, workbook
from .exceptions import PreconditionFailedError, ResponseTooLargeError, ValidationError

CATEGORIES = ("added", "removed", "updated", "unchanged")
EXPLANATION_VERSION = 1
ZERO = Fraction(0)
MAX_METRICS = 8
MAX_RESULT_CELLS = 10_000
MAX_RESPONSE_CHARS = 40_000
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
    recipes._scalar(value, location="Numeric value")
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
    floating = {
        (record.group, index)
        for record in records.values()
        if record.included
        for index, item in enumerate(recipe["metrics"])
        if item["op"] == "sum" and type(record.values[item["field"]]) is float
    }
    size = len(recipe["metrics"])
    comparisons = [
        (observed[key], oracle[key], [(key, i) not in floating for i in range(size)])
        for key in oracle
    ]
    comparisons.append(
        (
            _totals(observed, size),
            _totals(oracle, size),
            [not any(index == i for _, index in floating) for i in range(size)],
        )
    )
    for actual, expected, exact in comparisons:
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
        "numeric_match": "exact"
        if all(a == b for a, b, _ in comparisons)
        else "within_declared_tolerance",
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


def _metric_report(
    item: dict,
    index: int,
    observed_before: list,
    observed_after: list,
    exact_before: list,
    exact_after: list,
    breakdown: dict,
) -> dict:
    components = {
        category: breakdown["totals"][category][index] for category in CATEGORIES[:3]
    }
    contribution_delta = sum(components.values(), ZERO)
    exact_delta = exact_after[index] - exact_before[index]
    if (
        contribution_delta != exact_delta
        or breakdown["totals"]["unchanged"][index] != ZERO
    ):
        raise PreconditionFailedError(
            "Identity contributions do not reconcile",
            code="independent_reconciliation_failed",
        )
    delta = observed_after[index] - observed_before[index]
    adjustment = delta - exact_delta
    return {
        "name": item["as"],
        "op": item["op"],
        "before": _text(observed_before[index]),
        "after": _text(observed_after[index]),
        "delta": _text(delta),
        "independent_before": _text(exact_before[index]),
        "independent_after": _text(exact_after[index]),
        "components": {name: _text(value) for name, value in components.items()},
        "contribution_delta": _text(contribution_delta),
        "numeric_adjustment": _text(adjustment),
        "residual": _text(delta - contribution_delta - adjustment),
    }


def _present(
    *,
    before: dict,
    after: dict,
    fields: list[str],
    keys: list[str],
    recipe: dict,
    before_exact: dict,
    after_exact: dict,
    before_actual: dict,
    after_actual: dict,
    metric: str | None,
    group: dict | None,
    sample_limit: int,
) -> dict:
    aliases = [item["as"] for item in recipe["metrics"]]
    if metric is not None and metric not in aliases:
        raise ValidationError("metric must be an exact recipe metric alias")
    indexes = (
        [aliases.index(metric)] if metric is not None else list(range(len(aliases)))
    )
    groups = dict.fromkeys([*before_exact, *after_exact])
    selected = None
    if group is not None:
        recipes._object(
            group,
            allowed=set(recipe["group_by"]),
            required=set(recipe["group_by"]),
            location="group",
        )
        selected = tuple(
            _identity(recipes._scalar(group[field], location="group"))
            for field in recipe["group_by"]
        )
        if selected not in groups:
            raise ValidationError(
                "Selected group is absent from both recomputed results"
            )
    size = len(aliases)
    zeros = [ZERO] * size

    def scoped(data):
        return (
            data.get(selected, zeros) if selected is not None else _totals(data, size)
        )

    breakdown = _breakdown(before, after, fields, size, selected)
    scoped_values = [
        scoped(data)
        for data in (before_actual, after_actual, before_exact, after_exact)
    ]
    metrics = [
        _metric_report(
            recipe["metrics"][index],
            index,
            *scoped_values,
            breakdown,
        )
        for index in indexes
    ]
    display = {(): {}}
    for record in [*before.values(), *after.values()]:
        display.setdefault(
            record.group, {field: record.values[field] for field in recipe["group_by"]}
        )
    changed_groups = [
        key
        for key in groups
        if (selected is None or key == selected)
        and (
            key not in before_exact
            or key not in after_exact
            or any(
                before_exact.get(key, zeros)[i] != after_exact.get(key, zeros)[i]
                or before_actual.get(key, zeros)[i] != after_actual.get(key, zeros)[i]
                for i in indexes
            )
        )
    ]
    changed_groups.sort(
        key=lambda key: (
            -max(
                abs(after_exact.get(key, zeros)[i] - before_exact.get(key, zeros)[i])
                for i in indexes
            ),
            repr(key),
        )
    )
    group_sample = [
        {
            "group": changeset._bound_response_value(display[key]),
            "existed_before": key in before_exact,
            "exists_after": key in after_exact,
            "metrics": {
                aliases[i]: {
                    "before": _text(before_actual.get(key, zeros)[i]),
                    "after": _text(after_actual.get(key, zeros)[i]),
                    "delta": _text(
                        after_actual.get(key, zeros)[i]
                        - before_actual.get(key, zeros)[i]
                    ),
                    "contribution_delta": _text(
                        after_exact.get(key, zeros)[i] - before_exact.get(key, zeros)[i]
                    ),
                }
                for i in indexes
            },
        }
        for key in changed_groups[:sample_limit]
    ]
    pairs = [
        pair
        for pair in _pairs(before, after, fields)
        if pair[0] != "unchanged"
        and (
            selected is None
            or any(
                record and record.included and record.group == selected
                for record in pair[1:3]
            )
        )
    ]
    ranked = [(pair, _delta(pair[1], pair[2], size, selected)) for pair in pairs]
    ranked.sort(key=lambda item: -max(abs(item[1][i]) for i in indexes))
    samples = []
    for (category, old, new, changed), deltas in ranked[:sample_limit]:
        record = new if new else old
        changes = {
            field: {"before": old.values[field], "after": new.values[field]}
            for field in changed
        }
        samples.append(
            {
                "key": changeset._bound_response_value(
                    {field: record.values[field] for field in keys}
                ),
                "category": category,
                "before_row": old.row if old else None,
                "after_row": new.row if new else None,
                "before_included": bool(old and old.included),
                "after_included": bool(new and new.included),
                "before_group": changeset._bound_response_value(display[old.group])
                if old
                else None,
                "after_group": changeset._bound_response_value(display[new.group])
                if new
                else None,
                "fields_changed": changeset._bound_response_value(changes),
                "contribution_delta": {aliases[i]: _text(deltas[i]) for i in indexes},
            }
        )
    return {
        "scope": {"metric_names": [aliases[i] for i in indexes], "group": group},
        "metrics": metrics,
        "changes": {
            "counts": breakdown["counts"],
            "sample": samples,
            "truncated": len(pairs) > sample_limit,
            "omitted_contribution_delta": {
                aliases[i]: _text(
                    sum(
                        (deltas[i] for _, deltas in ranked[sample_limit:]),
                        ZERO,
                    )
                )
                for i in indexes
            },
        },
        "groups": {
            "compared_count": len(groups) if selected is None else 1,
            "changed_count": len(changed_groups),
            "sample": group_sample,
            "truncated": len(changed_groups) > sample_limit,
            "omitted_recipe_delta": {
                aliases[i]: _text(
                    sum(
                        (
                            after_actual.get(key, zeros)[i]
                            - before_actual.get(key, zeros)[i]
                            for key in changed_groups[sample_limit:]
                        ),
                        ZERO,
                    )
                )
                for i in indexes
            },
            "omitted_contribution_delta": {
                aliases[i]: _text(
                    sum(
                        (
                            after_exact.get(key, zeros)[i]
                            - before_exact.get(key, zeros)[i]
                            for key in changed_groups[sample_limit:]
                        ),
                        ZERO,
                    )
                )
                for i in indexes
            },
        },
    }


def _digest(value: Any, name: str) -> str | None:
    if value is None:
        return None
    if (
        not isinstance(value, str)
        or re.fullmatch(r"[0-9a-fA-F]{64}", value.strip()) is None
    ):
        raise ValidationError(f"{name} must be a 64-character SHA-256 hex digest")
    return value.strip().lower()


def _assert_response_budget(payload: dict) -> None:
    size = len(
        json.dumps(payload, separators=(",", ":"), ensure_ascii=True, allow_nan=False)
    )
    if size > MAX_RESPONSE_CHARS:
        raise ResponseTooLargeError(
            "Explanation exceeds its response budget",
            estimated_size=size,
            limit=MAX_RESPONSE_CHARS,
            hints=["reduce sample_limit", "select one metric and/or one group"],
        )


def _results_digest(*versions: dict) -> str:
    digest = hashlib.sha256()
    for groups in versions:
        digest.update(len(groups).to_bytes(8, "big"))
        for key in sorted(groups, key=repr):
            row = [
                [
                    (kind, _text(value) if kind == "number" else value)
                    for kind, value in key
                ],
                [_text(value) for value in groups[key]],
            ]
            encoded = changeset._canonical_json(row).encode("utf-8")
            digest.update(len(encoded).to_bytes(8, "big"))
            digest.update(encoded)
    return digest.hexdigest()


def explain_recipe_delta(
    before_filepath: str,
    after_filepath: str,
    recipe: dict,
    key_fields: list[str],
    metric: str | None = None,
    group: dict | None = None,
    sample_limit: int = 5,
    expected_before_sha256: str | None = None,
    expected_after_sha256: str | None = None,
    expected_recipe_sha256: str | None = None,
) -> dict:
    """Explain recomputed source results, not historical saved report values."""
    if type(sample_limit) is not int or not 1 <= sample_limit <= 10:
        raise ValidationError("sample_limit must be an integer from 1 to 10")
    if metric is not None and not isinstance(metric, str):
        raise ValidationError("metric must be an exact recipe metric alias")
    normalized, keys, fields = _contract(recipe, key_fields)
    if group is not None:
        recipes._object(
            group,
            allowed=set(normalized["group_by"]),
            required=set(normalized["group_by"]),
            location="group",
        )
        group = changeset._json_copy(group, location="group")
    recipe_sha = hashlib.sha256(
        changeset._canonical_json(normalized).encode("utf-8")
    ).hexdigest()
    expected_recipe = _digest(expected_recipe_sha256, "expected_recipe_sha256")
    if expected_recipe is not None and expected_recipe != recipe_sha:
        raise PreconditionFailedError(
            "Recipe definition changed; use the intended recipe",
            code="recipe_definition_mismatch",
        )
    paths = [
        workbook._canonical_workbook_path(path, must_exist=True)
        for path in (before_filepath, after_filepath)
    ]
    if any(path.suffix.lower() != ".xlsx" for path in paths):
        raise ValidationError("Both source files must end with .xlsx")
    if any(not path.is_file() for path in paths):
        raise ValidationError("Both source versions must be regular .xlsx files")
    if paths[0].samefile(paths[1]):
        raise ValidationError(
            "Source versions must be two distinct files, not aliases of the same file"
        )
    expected = [
        _digest(expected_before_sha256, "expected_before_sha256"),
        _digest(expected_after_sha256, "expected_after_sha256"),
    ]
    sources, digests = [], []
    loading_recipe = {
        **normalized,
        "checks": [*normalized["checks"], {"type": "unique_key", "fields": keys}],
    }
    with recipes._recipe_workspace() as temp:
        for index, path in enumerate(paths):
            baseline, candidate = (
                Path(temp) / f"source-{index}.xlsx",
                Path(temp) / f"unused-{index}.xlsx",
            )
            with workbook._exclusive_workbook_lock(path):
                digest = changeset._stage_workbook(path, baseline, candidate)
            if expected[index] is not None and expected[index] != digest:
                raise PreconditionFailedError(
                    "Source version changed; inspect fresh snapshots",
                    code="stale_workbook",
                    details={"source": "before" if index == 0 else "after"},
                    suggested_next_tool="explain_recipe_delta",
                )
            with workbook.safe_workbook(str(baseline)) as wb:
                sources.append(recipes._load_table(wb, loading_recipe))
            digests.append(digest)
    _compatible(*sources, fields)
    records = [_index(source, normalized, keys, fields) for source in sources]
    oracles = [_groups(items, normalized) for items in records]
    executions = [
        _reconcile(source, indexed, oracle, normalized)
        for source, indexed, oracle in zip(sources, records, oracles)
    ]
    observed = [item[0] for item in executions]
    result = _present(
        before=records[0],
        after=records[1],
        fields=fields,
        keys=keys,
        recipe=normalized,
        before_exact=oracles[0],
        after_exact=oracles[1],
        before_actual=observed[0],
        after_actual=observed[1],
        metric=metric,
        group=group,
        sample_limit=sample_limit,
    )
    result_sha = _results_digest(*oracles, *observed)
    numeric_policy = {
        "relative": _text(REL_TOL),
        "absolute": _text(ABS_TOL),
        "integer_only_and_count": "exact",
    }
    identity = {
        "version": EXPLANATION_VERSION,
        "recipe_engine_version": recipes.RECIPE_ENGINE_VERSION,
        "numeric_policy": numeric_policy,
        "result_sha256": result_sha,
        "before_sha256": digests[0],
        "after_sha256": digests[1],
        "recipe_sha256": recipe_sha,
        "key_fields": keys,
        "scope": result["scope"],
    }
    result.update(
        {
            "message": "Recipe delta explained and independently reconciled",
            "explanation_version": EXPLANATION_VERSION,
            "recipe_engine_version": recipes.RECIPE_ENGINE_VERSION,
            "result_sha256": result_sha,
            "status": "verified",
            "recipe_sha256": recipe_sha,
            "explanation_token": f"explanation_v{EXPLANATION_VERSION}_"
            + hashlib.sha256(
                changeset._canonical_json(identity).encode("utf-8")
            ).hexdigest(),
            "key_fields": keys,
            "numeric_encoding": "exact_decimal_strings",
            "sources": {
                name: {
                    **{
                        key: source[key]
                        for key in (
                            "sheet_name",
                            "table_name",
                            "range_ref",
                            "bindings",
                            "totals_rows_excluded",
                        )
                    },
                    "sha256": digest,
                    "row_count": len(source["rows"]),
                    "matched_row_count": sum(
                        record.included for record in indexed.values()
                    ),
                }
                for name, source, digest, indexed in zip(
                    ("before", "after"), sources, digests, records
                )
            },
            "verification": {
                "passed": True,
                "coverage": "complete_bounded_referenced_data_and_results",
                "basis": "same_recipe_recomputed_on_both_sources",
                "saved_outputs_checked": False,
                "unreferenced_fields_compared": False,
                "business_causes_inferred": False,
                "snapshots": "independently_captured_not_cross_file_atomic",
                "arithmetic": "exact_rational_on_loaded_decimal_representations",
                "numeric_match": "exact"
                if all(item[1]["numeric_match"] == "exact" for item in executions)
                else "within_declared_tolerance",
                "recipe_numeric_tolerance": numeric_policy,
                "group_metric_checks": sum(len(oracle) for oracle in oracles)
                * len(normalized["metrics"]),
                "global_metric_checks": 2 * len(normalized["metrics"]),
                "recipe_checks": {
                    name: item[1] for name, item in zip(("before", "after"), executions)
                },
            },
        }
    )
    _assert_response_budget(result)
    return result
