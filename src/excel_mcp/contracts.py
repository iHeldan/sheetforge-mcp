"""Explicit native-table bindings and bounded, executable data contracts."""

from __future__ import annotations

import hashlib
import json
import math
from typing import Any

from . import changeset, query, recipes, workbook
from .exceptions import PreconditionFailedError, ResponseTooLargeError, ValidationError
from .workbook import safe_workbook

CONTRACT_ENGINE_VERSION = 1
MAX_CONTRACT_BYTES = 32_000
MAX_CONTRACT_FIELDS = 32
MAX_CONTRACT_SUMMARY = 12_000
MAX_RECOVERY_EXAMPLES = 5


def _digest(value: Any) -> str:
    return hashlib.sha256(changeset._canonical_json(value).encode("utf-8")).hexdigest()


def normalize_contract(value: Any) -> dict:
    contract = changeset._json_copy(value, location="workbook_contract")
    try:
        size = len(changeset._canonical_json(contract).encode("utf-8"))
    except UnicodeError as exc:
        raise ValidationError("Contract strings must contain valid Unicode") from exc
    if size > MAX_CONTRACT_BYTES:
        raise ValidationError("workbook_contract exceeds 32000 bytes")
    recipes._object(
        contract,
        allowed={
            "version",
            "name",
            "source",
            "row_grain",
            "key_fields",
            "fields",
            "filters",
            "min_rows",
        },
        required={"version", "name", "source", "row_grain", "key_fields", "fields"},
        location="workbook_contract",
    )
    if type(contract["version"]) is not int or contract["version"] != 1:
        raise ValidationError("workbook_contract.version must be 1")
    for name in ("name", "row_grain"):
        recipes._name(contract[name], f"workbook_contract.{name}")
    fields = contract["fields"]
    if not isinstance(fields, list) or not 1 <= len(fields) <= MAX_CONTRACT_FIELDS:
        raise ValidationError("contract.fields must contain 1 to 32 fields")
    names, aliases = set(), set()
    for field in fields:
        recipes._object(
            field,
            allowed={
                "name",
                "headers",
                "type",
                "nullable",
                "meaning",
                "allowed_values",
                "min",
                "max",
            },
            required={"name", "headers", "type"},
            location="contract field",
        )
        name = recipes._name(field["name"], "field.name")
        if name in names:
            raise ValidationError("Contract logical field names must be unique")
        names.add(name)
        headers = recipes._field_list(field["headers"], "field.headers")
        if not 1 <= len(headers) <= 8 or aliases.intersection(headers):
            raise ValidationError(
                "Each field needs 1-8 distinct headers, not shared with another field"
            )
        aliases.update(headers)
        if field["type"] not in ("string", "number", "integer", "boolean"):
            raise ValidationError(
                "field.type must be string, number, integer or boolean"
            )
        field.setdefault("nullable", False)
        if type(field["nullable"]) is not bool:
            raise ValidationError("field.nullable must be a boolean")
        if "meaning" in field:
            recipes._name(field["meaning"], "field.meaning")
        if "allowed_values" in field:
            values = field["allowed_values"]
            if not isinstance(values, list) or not 1 <= len(values) <= 32:
                raise ValidationError(
                    "allowed_values must contain 1-32 non-null typed values"
                )
            if any(v is None or not _matches_type(v, field["type"]) for v in values):
                raise ValidationError(
                    "allowed_values must match the field type and cannot contain null"
                )
        for bound in ("min", "max"):
            if bound in field:
                if field["type"] not in ("integer", "number") or not _matches_type(
                    field[bound], field["type"]
                ):
                    raise ValidationError(
                        "Bounds require finite values matching a numeric field type"
                    )
        if "min" in field and "max" in field and field["min"] > field["max"]:
            raise ValidationError("field.min must not exceed max")
    keys = recipes._field_list(contract["key_fields"], "contract.key_fields")
    if not 1 <= len(keys) <= 8 or set(keys) - names:
        raise ValidationError("key_fields must contain 1-8 declared logical fields")
    contract.setdefault("min_rows", 1)
    if (
        type(contract["min_rows"]) is not int
        or not 0 <= contract["min_rows"] <= recipes.MAX_SOURCE_CELLS
    ):
        raise ValidationError("min_rows must be an integer from 0 to 50000")
    contract.setdefault("filters", [])
    if not isinstance(contract["filters"], list) or len(contract["filters"]) > 20:
        raise ValidationError("contract.filters must contain at most 20 predicates")
    # Use the existing closed filter/source grammar rather than a second query DSL.
    probe = recipes._normalize_recipe(
        {
            "version": 1,
            "source": contract["source"],
            "filters": contract["filters"],
            "group_by": [],
            "metrics": [{"op": "count", "as": "Rows"}],
            "output": {"sheet_name": "ContractProbe", "range_ref": "A1"},
        }
    )
    contract["filters"] = probe["filters"]
    if any(item["field"] not in names for item in contract["filters"]):
        raise ValidationError("Contract filters must reference declared logical fields")
    validated = query._normalize_filters(
        contract["filters"], [f["name"] for f in fields], []
    )
    for predicate, normalized in zip(contract["filters"], validated):
        predicate["op"] = normalized["op"]
    return contract


def _matches_type(value: Any, kind: str) -> bool:
    if kind == "string":
        return type(value) is str
    if kind == "boolean":
        return type(value) is bool
    if kind == "integer":
        return type(value) is int
    return type(value) is int or (type(value) is float and math.isfinite(value))


def _physical_recipe(contract: dict, bindings: dict) -> dict:
    return {
        "version": 1,
        "source": contract["source"],
        "group_by": list(bindings.values()),
        "metrics": [{"op": "count", "as": "ContractRows"}],
        "filters": [
            {**item, "field": bindings[item["field"]]} for item in contract["filters"]
        ],
        "checks": [],
    }


def inspect_table_contract(wb: Any, contract: dict, sample_limit: int) -> dict:
    selector = {
        "source": contract["source"],
        "group_by": [],
        "metrics": [],
        "filters": [],
        "checks": [],
    }
    metadata = recipes._load_table(wb, selector)
    headers = metadata["headers"]
    bindings, unresolved = {}, []
    for field in contract["fields"]:
        candidates = [header for header in field["headers"] if header in headers]
        if len(candidates) == 1:
            bindings[field["name"]] = candidates[0]
        else:
            unresolved.append(
                {
                    "field": field["name"],
                    "reason": "missing" if not candidates else "ambiguous",
                    "candidates": candidates,
                }
            )
    declared = {
        "row_grain": contract["row_grain"],
        "meanings": {
            field["name"]: field["meaning"]
            for field in contract["fields"]
            if "meaning" in field
        },
    }
    source_info = {
        key: metadata[key]
        for key in ("sheet_name", "table_name", "range_ref", "totals_rows_excluded")
    }
    report = {
        "name": contract["name"],
        "contract_sha256": _digest(contract),
        "engine_version": CONTRACT_ENGINE_VERSION,
        "source": source_info,
        "declared": declared,
        "bindings": bindings,
        "recipe_input_fields": [field["name"] for field in contract["fields"]],
        "field_usage": "Recipe input references use logical names (binding keys), not physical Excel headers (binding values). Metric 'as' labels are output names and stay unchanged.",
        "unresolved": unresolved,
        "checked": [],
        "source_row_count": len(metadata["rows"]),
        "coverage": "bindings only; no value rules evaluated"
        if unresolved
        else "all bounded source rows before filters; examples only are sampled",
        "status": "unresolved" if unresolved else "valid",
        "passed": not unresolved,
    }
    if unresolved:
        report["next_action"] = {
            "action": "ask_data_owner",
            "reason": "resolve_missing_or_ambiguous_headers",
            "instruction": "Do not select an alias, change the approved contract or write a report. Reinspect after an approved source/definition change.",
        }
        return report
    source = recipes._load_table(wb, _physical_recipe(contract, bindings))
    checked = report["checked"]
    checked.append(
        {
            "check": "min_rows",
            "passed": len(source["rows"]) >= contract["min_rows"],
            "expected_min": contract["min_rows"],
            "actual": len(source["rows"]),
        }
    )
    for field in contract["fields"]:
        index = headers.index(bindings[field["name"]])
        failures, examples = 0, []
        for offset, row in enumerate(source["rows"]):
            value = row[index]
            valid = (
                field["nullable"]
                if value is None
                else _matches_type(value, field["type"])
            )
            if valid and value is not None:
                if "allowed_values" in field:
                    valid = value in field["allowed_values"]
                if "min" in field:
                    valid = valid and value >= field["min"]
                if "max" in field:
                    valid = valid and value <= field["max"]
            if not valid:
                failures += 1
                if len(examples) < sample_limit:
                    examples.append(source["data_start_row"] + offset)
        checked.append(
            {
                "check": "field_rules",
                "field": field["name"],
                "passed": failures == 0,
                "failed_rows": failures,
                "row_examples": examples,
            }
        )
    key_indexes = [headers.index(bindings[name]) for name in contract["key_fields"]]
    seen, failures, examples = set(), 0, []
    for offset, row in enumerate(source["rows"]):
        values = [row[i] for i in key_indexes]
        key = tuple((recipes._kind(v), v) for v in values)
        invalid = (
            any(v is None or (isinstance(v, str) and not v.strip()) for v in values)
            or key in seen
        )
        seen.add(key)
        if invalid:
            failures += 1
            if len(examples) < sample_limit:
                examples.append(source["data_start_row"] + offset)
    checked.append(
        {
            "check": "complete_unique_key",
            "passed": failures == 0,
            "failed_rows": failures,
            "row_examples": examples,
        }
    )
    report["passed"] = all(item["passed"] for item in checked)
    report["status"] = "valid" if report["passed"] else "violated"
    report["next_action"] = (
        {
            "tool": "run_workbook_recipe",
            "mode": "preview",
            "instruction": "Use recipe_input_fields in group_by, metrics.field, filters.field and source check fields; pass the identical contract as workbook_contract.",
        }
        if report["passed"]
        else {
            "action": "ask_data_owner",
            "reason": "source_rules_failed",
            "instruction": "Do not weaken the contract or filter away failed source rows. No report write is safe until the source or definition is explicitly approved and reinspected.",
        }
    )
    report["binding_sha256"] = _digest(
        {
            "engine": CONTRACT_ENGINE_VERSION,
            "contract": contract,
            "source": source_info,
            "headers": headers,
            "bindings": bindings,
            "cells": source["bindings"],
        }
    )
    return report


def _field_recovery(recipe: dict, bindings: dict, invalid: set[str]) -> dict:
    references = [
        (f"/recipe/group_by/{i}", field) for i, field in enumerate(recipe["group_by"])
    ]
    references += [
        (f"/recipe/metrics/{i}/field", item["field"])
        for i, item in enumerate(recipe["metrics"])
        if "field" in item
    ]
    references += [
        (f"/recipe/filters/{i}/field", item["field"])
        for i, item in enumerate(recipe["filters"])
    ]
    for i, check in enumerate(recipe["checks"]):
        if check["type"] == "unique_key":
            references += [
                (f"/recipe/checks/{i}/fields/{j}", field)
                for j, field in enumerate(check["fields"])
            ]
        elif check["type"] == "sum_matches":
            references.append(
                (f"/recipe/checks/{i}/source_field", check["source_field"])
            )
    reverse = {physical: logical for logical, physical in bindings.items()}
    failures = [(path, field) for path, field in references if field in invalid]
    examples = []
    for path, field in failures[:MAX_RECOVERY_EXAMPLES]:
        example = {"path": path, "provided": field}
        if field in reverse:
            example["replace_with"] = reverse[field]
        else:
            example["instruction"] = (
                "Choose an existing logical field only if it matches the intended meaning; no exact physical-header replacement is known."
            )
        examples.append(example)
    return _bounded_report(
        {
            "argument": "recipe",
            "allowed_input_fields": list(bindings),
            "invalid_reference_count": len(failures),
            "examples": examples,
            "examples_truncated": len(failures) > len(examples),
            "retry_mode": "preview",
            "instruction": "Correct only these Recipe input references; keep the approved workbook_contract, source selector, filter values/operators and output aliases unchanged, and request a fresh preview. Suggestions are not applied automatically.",
        }
    )


def compile_contract_recipe(
    wb: Any, recipe: dict, contract: dict, sample_limit: int
) -> tuple[dict, dict]:
    if recipe["source"] != contract["source"]:
        raise ValidationError(
            "Recipe and contract source selectors must be identical; copy workbook_contract.source to recipe.source without changing the approved contract",
            code="contract_recipe_source_mismatch",
            details=_bounded_report(
                {
                    "path": "/recipe/source",
                    "replace_with": contract["source"],
                    "retry_mode": "preview",
                }
            ),
            suggested_next_tool="run_workbook_recipe",
        )
    report = inspect_table_contract(wb, contract, sample_limit)
    if not report["passed"]:
        raise PreconditionFailedError(
            "Workbook contract is unresolved or violated; inspect it before running",
            code="workbook_contract_failed",
            details=_bounded_report(report),
            suggested_next_tool="inspect_workbook_contract",
        )
    bindings = report["bindings"]
    invalid = set(recipes._source_fields(recipe)) - set(bindings)
    if invalid:
        raise ValidationError(
            "Contract recipes must use declared logical fields only; use binding keys, not physical Excel header values. See details.examples for exact Recipe reference corrections",
            code="contract_recipe_fields_invalid",
            details=_field_recovery(recipe, bindings, invalid),
            suggested_next_tool="run_workbook_recipe",
        )
    output_headers = [*recipe["group_by"], *[m["as"] for m in recipe["metrics"]]]
    if len({name.strip().casefold() for name in output_headers}) != len(output_headers):
        raise ValidationError(
            "Contract recipe output headers must be unambiguous logical names"
        )
    physical = changeset._json_copy(recipe, location="recipe")
    physical["group_by"] = [bindings[name] for name in recipe["group_by"]]
    for metric in physical["metrics"]:
        if "field" in metric:
            metric["field"] = bindings[metric["field"]]
    physical["filters"] = [
        {**item, "field": bindings[item["field"]]}
        for item in [*contract["filters"], *recipe["filters"]]
    ]
    for check in physical["checks"]:
        if check["type"] == "unique_key":
            check["fields"] = [bindings[name] for name in check["fields"]]
        elif check["type"] == "sum_matches":
            check["source_field"] = bindings[check["source_field"]]
            if check["output_field"] in recipe["group_by"]:
                check["output_field"] = bindings[check["output_field"]]
    return recipes._normalize_recipe(physical), report


def _bounded_report(report: dict) -> dict:
    bounded = changeset._bound_response_value(report)
    size = len(
        json.dumps(bounded, ensure_ascii=True, separators=(",", ":"), allow_nan=False)
    )
    if size > MAX_CONTRACT_SUMMARY:
        raise ResponseTooLargeError(
            "Contract report exceeds 12000 characters; reduce fields, aliases or sample_limit",
            estimated_size=size,
            limit=MAX_CONTRACT_SUMMARY,
        )
    return bounded


def inspect_workbook_contract(
    filepath: str, contract: dict, sample_limit: int = 3
) -> dict:
    if type(sample_limit) is not int or not 1 <= sample_limit <= 5:
        raise ValidationError(
            "sample_limit must be an integer from 1 to 5; omit it for the default 3, or retry the same inspection with sample_limit=3. Recipe's 1-10 limit does not apply to inspection",
            code="contract_sample_limit_invalid",
            details={
                "argument": "sample_limit",
                "allowed": [1, 2, 3, 4, 5],
                "suggested_arguments": {"sample_limit": 3},
            },
            suggested_next_tool="inspect_workbook_contract",
        )
    normalized = normalize_contract(contract)
    path = workbook._canonical_workbook_path(filepath, must_exist=True)
    if path.suffix.lower() != ".xlsx":
        raise ValidationError("filepath must end with .xlsx")
    with recipes._recipe_workspace() as temp:
        baseline, candidate = temp / "baseline.xlsx", temp / "candidate.xlsx"
        with workbook._exclusive_workbook_lock(path):
            digest = changeset._stage_workbook(path, baseline, candidate)
        with safe_workbook(str(baseline)) as wb:
            report = inspect_table_contract(wb, normalized, sample_limit)
    return {"source_sha256": digest, "persisted": False, **_bounded_report(report)}


def normalize_contract_assertion(assertion: dict) -> dict:
    assertion["contract"] = normalize_contract(assertion["contract"])
    assertion["expected_binding_sha256"] = changeset._normalize_sha256(
        assertion["expected_binding_sha256"], required=True
    )
    return assertion


def assert_contract(wb: Any, assertion: dict, sample_limit: int) -> tuple[bool, dict]:
    report = inspect_table_contract(wb, assertion["contract"], min(sample_limit, 5))
    passed = (
        report["passed"]
        and report.get("binding_sha256") == assertion["expected_binding_sha256"]
    )
    return passed, {
        "status": report["status"],
        "contract_sha256": report["contract_sha256"],
        "binding_matches": report.get("binding_sha256")
        == assertion["expected_binding_sha256"],
        "passed": passed,
    }
