"""Bounded counterexamples for two explicit recipe filter definitions."""

from itertools import islice, product
import math

from . import recipes, workbook
from . import reporting_evidence as evidence
from .exceptions import ValidationError

MAX_CASES = 128


def _different(left, right, group_count):
    if left["headers"] != right["headers"] or len(left["rows"]) != len(right["rows"]):
        return True
    for a, b in zip(left["rows"], right["rows"]):
        # Numeric group identities are exact; metric tolerance must never merge
        # two nearby numeric identifiers into the same group.
        if [evidence.identity(v) for v in a[:group_count]] != [
            evidence.identity(v) for v in b[:group_count]
        ]:
            return True
        if not recipes._equal_result([a[group_count:]], [b[group_count:]]):
            return True
    return False


def probe_recipe_rules(
    filepath: str, left_recipe: dict, right_recipe: dict, sample_limit: int = 3
) -> dict:
    evidence.sample_size(sample_limit)
    left, right = (recipes._normalize_recipe(r) for r in (left_recipe, right_recipe))
    if {k: v for k, v in left.items() if k != "filters"} != {
        k: v for k, v in right.items() if k != "filters"
    }:
        raise ValidationError("Rule Probe v1 requires identical recipes except filters")
    if (
        left["checks"]
        or any(m["op"] not in ("sum", "count") for m in left["metrics"])
        or len(left["metrics"]) > 8
    ):
        raise ValidationError(
            "Rule Probe v1 supports up to 8 sum/count metrics and no checks"
        )
    fields = list(
        dict.fromkeys(recipes._source_fields(left) + recipes._source_fields(right))
    )
    with evidence.capture([filepath]) as [(path, sha)]:
        with workbook.safe_workbook(str(path)) as wb:
            source = evidence.load_fields(wb, left["source"], fields)
        # Validate BOTH actual input definitions, even if an early witness exists.
        for candidate in (left, right):
            recipes._validate_filter_inputs(source, candidate)
        results = [recipes._aggregate(source, r) for r in (left, right)]
        real_different = _different(*results, len(left["group_by"]))
        fields_indexes = [source["headers"].index(f) for f in fields]
        witness = None
        observed = synthetic = invalid = 0

        def examine(row, kind, row_number=None):
            nonlocal witness
            small = {**source, "rows": [row]}
            for candidate in (left, right):
                recipes._validate_filter_inputs(small, candidate)
            a, b = [recipes._aggregate(small, r) for r in (left, right)]
            if _different(a, b, len(left["group_by"])):
                witness = {
                    "kind": kind,
                    "source_row": row_number,
                    "record": {f: row[i] for f, i in zip(fields, fields_indexes)},
                    "headers": a["headers"],
                    "left_rows": a["rows"],
                    "right_rows": b["rows"],
                }

        for offset, row in enumerate(source["rows"][:MAX_CASES]):
            observed += 1
            examine(row, "observed_single_row", source["data_start_row"] + offset)
            if witness:
                break
        # Synthetic examples are hypothetical, not claims about the source or its
        # business constraints. The search space is explicitly bounded, never a proof.
        space_size = 0
        values_truncated = False
        if not witness:
            if len(fields) <= 4:
                choices = []
                for field, index in zip(fields, fields_indexes):
                    values = [
                        row[index]
                        for row in source["rows"][:8]
                        if row[index] is not None
                    ]
                    for r in (left, right):
                        for spec in r["filters"]:
                            if spec["field"] == field:
                                values.extend(spec.get("values", [spec.get("value")]))
                    nonnull = [v for v in values if v is not None]
                    kind = recipes._kind(nonnull[0]) if nonnull else "number"
                    if any(recipes._kind(v) != kind for v in nonnull):
                        raise ValidationError(
                            "Synthetic probe operands have incompatible types"
                        )
                    if kind == "number":
                        values += [-1, 0, 1]
                        for value in nonnull[:8]:
                            values += [value - 1, value + 1]
                    elif kind == "str":
                        values += ["", "other", "posted", "cancelled"]
                    elif kind == "bool":
                        values += [False, True]
                    values.append(None)
                    unique = {}
                    for v in values:
                        if type(v) is float and not math.isfinite(v):
                            continue
                        unique.setdefault(evidence.identity(v), v)
                    values_truncated |= len(unique) > 16
                    choices.append(list(unique.values())[:16])
                space_size = math.prod(map(len, choices))
                for values in islice(product(*choices), MAX_CASES):
                    row = [None] * len(source["headers"])
                    for index, value in zip(fields_indexes, values):
                        row[index] = value
                    synthetic += 1
                    try:
                        examine(row, "hypothetical_single_row")
                    except ValidationError:
                        invalid += 1
                        continue
                    if witness:
                        break
        real_summary = [
            {
                "headers": r["headers"],
                "rows_sample": r["rows"][:sample_limit],
                "row_count": len(r["rows"]),
                "result_sha256": evidence.digest(r["rows"]),
                "sample_truncated": len(r["rows"]) > sample_limit,
            }
            for r in results
        ]
        return evidence.bounded(
            {
                "version": 1,
                "read_only": True,
                "source_sha256": sha,
                "status": "counterexample_found"
                if witness
                else "real_results_differ"
                if real_different
                else "no_difference_found",
                "real_results_differ": real_different,
                "real_results": real_summary,
                "witness": witness,
                "observed_cases_tested": observed,
                "synthetic_cases_tested": synthetic,
                "invalid_synthetic_cases": invalid,
                "search_truncated": len(source["rows"]) > observed
                or space_size > synthetic
                or len(fields) > 4
                or values_truncated,
                "synthetic_value_choices_truncated": values_truncated,
                "numeric_tolerance": {
                    "relative": recipes.NUMERIC_REL_TOL,
                    "absolute": recipes.NUMERIC_ABS_TOL,
                },
                "equivalence_proven": False,
                "next_action": "Ask the data owner which explicit rule is intended; retain an approved witness as a client-owned test case. No rule was selected",
                "limits": "128 observed singleton cases, then 128 hypothetical cases across at most 4 referenced fields; no business-constraint inference",
            }
        )
