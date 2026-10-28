"""Compare bounded keyed alternatives and commit only verified metric outcomes."""

import hashlib
import hmac
import json
import re

from . import changeset, explanations, rebase, recipes, workbook
from .exceptions import PreconditionFailedError, ResponseTooLargeError, ValidationError

SCENARIO_VERSION = 1
MAX_SPEC_BYTES = 256_000
MAX_COMPARISON_CHARS = 40_000
MAX_COMMIT_SUMMARY_CHARS = 8000


def _digest(value):
    try:
        return hashlib.sha256(
            changeset._canonical_json(value).encode("utf-8")
        ).hexdigest()
    except UnicodeError as exc:
        raise ValidationError("Scenario inputs must contain valid Unicode") from exc


def _analysis(recipe, key_fields, guards):
    recipe, keys, fields = explanations._contract(recipe, key_fields)
    guards = changeset._json_copy(guards, location="guards")
    if not isinstance(guards, list) or not 1 <= len(guards) <= 20:
        raise ValidationError("guards must contain 1-20 metric bounds")
    aliases = {item["as"] for item in recipe["metrics"]}
    for guard in guards:
        recipes._object(
            guard,
            allowed={"metric", "scope", "min", "max"},
            required={"metric", "scope"},
            location="guard",
        )
        recipes._name(guard["metric"], "guard.metric")
        if guard["metric"] not in aliases:
            raise ValidationError("guard.metric must be an exact metric alias")
        if not isinstance(guard["scope"], str) or guard["scope"] not in {
            "overall",
            "each_group",
        }:
            raise ValidationError("guard.scope must be overall or each_group")
        if guard["scope"] == "each_group" and not recipe["group_by"]:
            raise ValidationError("each_group requires nonempty group_by")
        if "min" not in guard and "max" not in guard:
            raise ValidationError("Every guard requires min and/or max")
        for bound in ("min", "max"):
            if bound in guard:
                recipes._scalar(guard[bound], location=f"guard.{bound}")
                if type(guard[bound]) not in (int, float):
                    raise ValidationError("Guard bounds must be finite JSON numbers")
        if (
            "min" in guard
            and "max" in guard
            and explanations._number(guard["min"]) > explanations._number(guard["max"])
        ):
            raise ValidationError("guard.min must not exceed guard.max")
    return recipe, keys, fields, guards


def _contract(spec):
    spec = changeset._json_copy(spec, location="scenario_set")
    try:
        size = len(changeset._canonical_json(spec).encode("utf-8"))
    except UnicodeError as exc:
        raise ValidationError("scenario_set must contain valid Unicode") from exc
    if size > MAX_SPEC_BYTES:
        raise ValidationError(f"scenario_set exceeds {MAX_SPEC_BYTES} bytes")
    recipes._object(
        spec,
        allowed={"version", "recipe", "key_fields", "guards", "scenarios"},
        required={"version", "recipe", "key_fields", "guards", "scenarios"},
        location="scenario_set",
    )
    if type(spec["version"]) is not int or spec["version"] != SCENARIO_VERSION:
        raise ValidationError("scenario_set.version must be integer 1")
    recipe, keys, _, guards = _analysis(
        spec["recipe"], spec["key_fields"], spec["guards"]
    )
    spec.update(recipe=recipe, key_fields=keys, guards=guards)
    alternatives = spec["scenarios"]
    if not isinstance(alternatives, list) or not 1 <= len(alternatives) <= 5:
        raise ValidationError("scenarios must contain 1-5 alternatives")
    seen, patches = set(), {}
    for item in alternatives:
        recipes._object(
            item,
            allowed={"id", "updates"},
            required={"id", "updates"},
            location="scenario",
        )
        identifier = item["id"]
        if (
            not isinstance(identifier, str)
            or re.fullmatch(r"[A-Za-z0-9_-]{1,64}", identifier) is None
        ):
            raise ValidationError("scenario.id must be 1-64 ASCII letters/digits/_/-")
        if identifier in seen:
            raise ValidationError("Duplicate scenario id")
        seen.add(identifier)
        patches[identifier] = rebase._contract(
            {
                "version": 1,
                "source": recipe["source"],
                "key_fields": keys,
                "read_fields": [],
                "updates": item["updates"],
            }
        )
    return spec, patches


def _evaluate(wb, recipe, keys):
    fields = list(dict.fromkeys([*recipes._source_fields(recipe), *keys]))
    # Also validate native-column metadata and calculated-column overrides.
    rebase._load(wb, {"source": recipe["source"], "key_fields": keys}, fields)
    loading = {
        **recipe,
        "checks": [*recipe["checks"], {"type": "unique_key", "fields": keys}],
    }
    source = recipes._load_table(wb, loading)
    records = explanations._index(source, recipe, keys, fields)
    exact = explanations._groups(records, recipe)
    observed, verification = explanations._reconcile(source, records, exact, recipe)
    # Numeric 1/1.0 may roundtrip to the same Excel scalar. Preserve bool/text kinds.
    rows = [
        [
            [recipes._kind(value), explanations._text(explanations._number(value))]
            if type(value) in (int, float)
            else [recipes._kind(value), value]
            for value in row
        ]
        for row in source["rows"]
    ]
    input_digest = _digest(
        {
            "headers": source["headers"],
            "rows": rows,
            "sheet_name": source["sheet_name"],
            "range_ref": source["range_ref"],
            "totals_rows_excluded": source["totals_rows_excluded"],
        }
    )
    return {
        "source": source,
        "records": records,
        "exact": exact,
        "observed": observed,
        "verification": verification,
        "input_digest": input_digest,
    }


def _guard_results(before, after, recipe, guards, sample_limit):
    aliases = [item["as"] for item in recipe["metrics"]]
    results = []
    for guard in guards:
        index = aliases.index(guard["metric"])
        samples, failed, evaluated = [], 0, 0
        if guard["scope"] == "overall":
            values = [(None, explanations._totals(after, len(aliases))[index])]
        else:
            # A vanished group cannot turn a minimum into a vacuously true check.
            union = list(dict.fromkeys([*before, *after]))
            values = [
                (group, after[group][index] if group in after else None)
                for group in union
            ]
        for group, value in values:
            evaluated += 1
            passed = (
                value is not None
                and ("min" not in guard or value >= explanations._number(guard["min"]))
                and ("max" not in guard or value <= explanations._number(guard["max"]))
            )
            if not passed:
                failed += 1
                if len(samples) < sample_limit:
                    # Typed identities preserve large numeric labels without float conversion.
                    samples.append(
                        {
                            "group": None
                            if group is None
                            else {
                                field: {
                                    "type": kind,
                                    "value": explanations._text(label)
                                    if kind == "number"
                                    else label,
                                }
                                for field, (kind, label) in zip(
                                    recipe["group_by"], group
                                )
                            },
                            "actual": None
                            if value is None
                            else explanations._text(value),
                            "reason": "missing_group"
                            if value is None
                            else "out_of_bounds",
                        }
                    )
        observed_values = [value for _, value in values if value is not None]
        results.append(
            {
                "metric": guard["metric"],
                "scope": guard["scope"],
                **{
                    b: explanations._text(explanations._number(guard[b]))
                    for b in ("min", "max")
                    if b in guard
                },
                "passed": evaluated > 0 and failed == 0,
                "evaluated_count": evaluated,
                "failed_count": failed,
                "empty_coverage": evaluated == 0,
                "failures_sample": samples,
                "truncated": failed > len(samples),
                "observed_min": explanations._text(min(observed_values))
                if observed_values
                else None,
                "observed_max": explanations._text(max(observed_values))
                if observed_values
                else None,
            }
        )
    return results


def _outcome(before, after, recipe, keys, guards, sample_limit):
    fields = list(dict.fromkeys([*recipes._source_fields(recipe), *keys]))
    explanations._compatible(before["source"], after["source"], fields)
    checks = _guard_results(
        before["exact"], after["exact"], recipe, guards, sample_limit
    )
    result_sha = _digest(
        {
            "version": SCENARIO_VERSION,
            "recipe_engine_version": recipes.RECIPE_ENGINE_VERSION,
            "explanation_version": explanations.EXPLANATION_VERSION,
            "recipe": recipe,
            "key_fields": keys,
            "guards": guards,
            "input_digests": [before["input_digest"], after["input_digest"]],
            "results": explanations._results_digest(
                before["exact"], after["exact"], before["observed"], after["observed"]
            ),
            "numeric_policy": [str(explanations.REL_TOL), str(explanations.ABS_TOL)],
        }
    )
    size = len(recipe["metrics"])
    before_totals = explanations._totals(before["exact"], size)
    after_totals = explanations._totals(after["exact"], size)
    return {
        "result_sha256": result_sha,
        "passes_guards": all(check["passed"] for check in checks),
        "guards": checks,
        "metrics": [
            {
                "name": item["as"],
                "before": explanations._text(a),
                "after": explanations._text(b),
                "delta": explanations._text(b - a),
            }
            for item, a, b in zip(recipe["metrics"], before_totals, after_totals)
        ],
        "group_count": {"before": len(before["exact"]), "after": len(after["exact"])},
        "numeric_match": "exact"
        if all(
            state["verification"]["numeric_match"] == "exact"
            for state in (before, after)
        )
        else "within_declared_tolerance",
    }


def normalize_outcome_assertion(assertion):
    recipe, keys, _, guards = _analysis(
        assertion["recipe"], assertion["key_fields"], assertion["guards"]
    )
    digest = changeset._normalize_sha256(
        assertion["expected_result_sha256"], required=True
    )
    if digest is None:
        raise ValidationError("expected_result_sha256 is required")
    return {
        "type": "table_scenario_matches",
        "recipe": recipe,
        "key_fields": keys,
        "guards": guards,
        "expected_result_sha256": digest,
    }


def assert_outcome(baseline_wb, candidate_wb, assertion, sample_limit):
    recipe, keys = assertion["recipe"], assertion["key_fields"]
    before = _evaluate(baseline_wb, recipe, keys)
    after = _evaluate(candidate_wb, recipe, keys)
    result = _outcome(
        before, after, recipe, keys, assertion["guards"], min(sample_limit, 5)
    )
    matches = hmac.compare_digest(
        result["result_sha256"], assertion["expected_result_sha256"]
    )
    return result["passes_guards"] and matches, {
        "result_matches": matches,
        "passes_guards": result["passes_guards"],
        "result_sha256": result["result_sha256"],
        "guard_count": len(result["guards"]),
        "failed_guards": [g for g in result["guards"] if not g["passed"]][
            :sample_limit
        ],
    }


def _budget(payload, limit):
    size = len(
        json.dumps(payload, ensure_ascii=True, allow_nan=False, separators=(",", ":"))
    )
    if size > limit:
        raise ResponseTooLargeError(
            "Scenario response exceeds pre-commit budget",
            estimated_size=size,
            limit=limit,
            hints=["reduce sample_limit, alternatives or guards"],
        )


def evaluate_workbook_scenarios(
    filepath: str,
    scenario_set: dict,
    mode: str = "compare",
    selected_scenario: str | None = None,
    expected_workbook_sha256: str | None = None,
    scenario_token: str | None = None,
    sample_limit: int = 3,
) -> dict:
    """Compare from one captured source; commit one exact guard-backed alternative."""
    if not isinstance(mode, str) or mode not in {"compare", "commit"}:
        raise ValidationError("mode must be compare or commit")
    if type(sample_limit) is not int or not 1 <= sample_limit <= 5:
        raise ValidationError("sample_limit must be an integer from 1 to 5")
    spec, patches = _contract(scenario_set)
    if mode == "compare" and (
        selected_scenario is not None or scenario_token is not None
    ):
        raise ValidationError("Selection and scenario_token are commit-only parameters")
    if mode == "commit" and (
        not isinstance(selected_scenario, str) or selected_scenario not in patches
    ):
        raise ValidationError("Commit requires a known selected_scenario id")
    expected = changeset._normalize_sha256(
        expected_workbook_sha256, required=mode == "commit"
    )
    if mode == "commit" and (
        not isinstance(scenario_token, str)
        or re.fullmatch(r"scenario_v1_[0-9a-f]{64}", scenario_token) is None
    ):
        raise ValidationError("Commit requires a scenario_v1_ token from comparison")
    source = workbook._canonical_workbook_path(filepath, must_exist=True)
    if source.suffix.lower() != ".xlsx" or not source.is_file():
        raise ValidationError("Scenario source must be a regular .xlsx file")
    recipe, keys, guards = spec["recipe"], spec["key_fields"], spec["guards"]
    compiled = []
    with recipes._recipe_workspace() as temp:
        baseline = temp / "baseline.xlsx"
        with workbook._exclusive_workbook_lock(source):
            source_sha = changeset._stage_workbook(
                source, baseline, temp / "unused.xlsx"
            )
        if expected is not None and expected != source_sha:
            raise PreconditionFailedError(
                "Workbook changed; compare scenarios again",
                code="stale_workbook",
                suggested_next_tool="evaluate_workbook_scenarios",
            )
        with workbook.safe_workbook(str(baseline)) as base_wb:
            before = _evaluate(base_wb, recipe, keys)
            for identifier, (patch, fields) in patches.items():
                if mode == "commit" and identifier != selected_scenario:
                    continue
                operations, assertions, patch_summary = rebase._compile(
                    base_wb, base_wb, patch, fields, sample_limit
                )
                if len(operations) > 49:
                    raise ValidationError(
                        "Scenarios support at most 49 contiguous column spans"
                    )
                with workbook.safe_workbook(str(baseline)) as candidate_wb:
                    for operation in operations:
                        args = operation["args"]
                        c1, r1, _, _ = changeset._range_bounds(
                            args["start_cell"], location="target"
                        )
                        for offset, row in enumerate(args["data"]):
                            candidate_wb[args["sheet_name"]].cell(
                                r1 + offset, c1
                            ).value = row[0]
                    after = _evaluate(candidate_wb, recipe, keys)
                    outcome = _outcome(
                        before, after, recipe, keys, guards, sample_limit
                    )
                assertions.append(
                    {
                        "type": "table_scenario_matches",
                        "recipe": recipe,
                        "key_fields": keys,
                        "guards": guards,
                        "expected_result_sha256": outcome["result_sha256"],
                    }
                )
                assertions = changeset._normalize_assertions(assertions)
                if operations:
                    operations = changeset._normalize_operations(operations)
                plan = changeset._plan_payload(
                    source=source,
                    source_sha256=source_sha,
                    operations=operations,
                    assertions=assertions,
                    create_snapshot=True,
                    snapshot_filepath=None,
                )
                cs_token = changeset._changeset_token(plan)
                token = "scenario_v1_" + _digest(
                    {
                        "version": SCENARIO_VERSION,
                        "scenario_set": spec,
                        "selected_scenario": identifier,
                        "changeset_token": cs_token,
                    }
                )
                compiled.append(
                    (
                        operations,
                        assertions,
                        cs_token,
                        {
                            "id": identifier,
                            "scenario_token": token,
                            **outcome,
                            "changed_cell_count": patch_summary["counts"]["apply"],
                            "write_sample": patch_summary["sample"],
                            "writes_truncated": patch_summary["truncated"],
                        },
                    )
                )
    common = {
        "scenario_version": SCENARIO_VERSION,
        "mode": mode,
        "expected_workbook_sha256": source_sha,
        "numeric_encoding": "exact_decimal_strings",
        "verification": {
            "coverage": "complete_bounded_referenced_data_and_results",
            "guard_basis": "exact_rational_on_loaded_decimal_values",
            "recipe_relative_tolerance": str(recipes.NUMERIC_REL_TOL),
            "recipe_absolute_tolerance": str(recipes.NUMERIC_ABS_TOL),
            "saved_report_cells_checked": False,
            "business_correctness_claimed": False,
        },
    }
    results = []
    for operations, assertions, cs_token, summary in compiled:
        summary = {
            **summary,
            "write_sample": changeset._bound_response_value(summary["write_sample"]),
            "guards": [
                {
                    **guard,
                    "failures_sample": changeset._bound_response_value(
                        guard["failures_sample"]
                    ),
                }
                for guard in summary["guards"]
            ],
        }
        if mode == "commit":
            if not hmac.compare_digest(scenario_token, summary["scenario_token"]):
                raise PreconditionFailedError(
                    "Scenario token does not match source, alternatives, recipe or guards",
                    code="scenario_token_mismatch",
                    suggested_next_tool="evaluate_workbook_scenarios",
                )
            _budget({**common, "scenario": summary}, MAX_COMMIT_SUMMARY_CHARS)
            if not summary["passes_guards"]:
                raise PreconditionFailedError(
                    "Selected scenario violates guards; no changes written",
                    code="scenario_guard_failed",
                    details={"scenario": summary},
                    suggested_next_tool="evaluate_workbook_scenarios",
                )
        ready = summary["passes_guards"]
        transaction = None
        if ready and operations:
            transaction = changeset.apply_workbook_changeset(
                str(source),
                operations,
                assertions,
                mode="preview" if mode == "compare" else "commit",
                expected_workbook_sha256=source_sha,
                changeset_token=cs_token,
                sample_limit=sample_limit,
            )
            ready = transaction.get("ready_to_commit", transaction["persisted"])
        if mode == "commit":
            if transaction:
                transaction.pop("changeset_token", None)
            return {
                **(transaction or {"persisted": False}),
                **common,
                "ready_to_commit": ready,
                "scenario": summary,
                "message": "Scenario committed and verified"
                if transaction
                else "Selected scenario already satisfied in captured source; no write performed",
            }
        results.append(
            {
                **summary,
                "ready_to_commit": ready,
                "snapshot": transaction["snapshot"]
                if transaction
                else {"enabled": False, "status": "not_created"},
            }
        )
    result = {
        **common,
        "persisted": False,
        "scenarios": results,
        "message": "Alternatives compared; no source changes or automatic selection",
    }
    _budget(result, MAX_COMPARISON_CHARS)
    return result
