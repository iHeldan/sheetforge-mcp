"""Client-owned report fingerprints; no signatures, history store or write authority."""

import re

from . import changeset, recipes, workbook
from . import reporting_evidence as evidence
from .exceptions import PreconditionFailedError, ValidationError

VERSION = 1
MAX_RECEIPT_CHARS = 6_000


def _atom(value, data_type=None):
    if data_type in ("f", "e"):
        return [data_type, value if isinstance(value, str) else type(value).__name__]
    if type(value) in (int, float):
        return ["number", evidence.text(evidence.number(value))]
    if value is None or type(value) in (str, bool):
        return [type(value).__name__, value]
    # Receipts start from literal JSON values only. A later unsupported cell
    # must remain a mismatch rather than crash or masquerade as a literal.
    return ["unsupported", type(value).__name__]


def _row_fingerprints(rows, group_count):
    indexed = {}
    for row in rows:
        key = evidence.digest([_atom(v) for v in row[:group_count]])
        if key in indexed:
            raise ValidationError("Output has duplicate group identities")
        indexed[key] = [evidence.digest(_atom(v)) for v in row[group_count:]]
    return indexed


def _source_digest(source, recipe):
    fields = recipes._source_fields(recipe)
    indexes = [source["headers"].index(f) for f in fields]
    return evidence.digest(
        {
            "fields": fields,
            "rows": [[_atom(row[i]) for i in indexes] for row in source["rows"]],
        }
    )


def make_receipt(requested, contract, source, compiled, result, padded):
    if any(all(value is None for value in row) for row in result["rows"]):
        raise ValidationError(
            "Report receipts cannot distinguish an entirely null result row from blank padding"
        )
    receipt = {
        "version": VERSION,
        "recipe_engine_version": recipes.RECIPE_ENGINE_VERSION,
        "recipe": requested,
        "workbook_contract": contract,
        "source_content_sha256": _source_digest(source, compiled),
        "output_grid_sha256": evidence.digest(
            [[_atom(v) for v in row] for row in padded]
        ),
        "rows": result["rows"],
    }
    receipt["receipt_sha256"] = evidence.digest(receipt)
    return evidence.bounded(receipt, MAX_RECEIPT_CHARS)


def _normalize(receipt):
    receipt = changeset._json_copy(receipt, location="report_receipt")
    evidence.bounded(receipt, MAX_RECEIPT_CHARS)
    keys = {
        "version",
        "recipe_engine_version",
        "recipe",
        "workbook_contract",
        "source_content_sha256",
        "output_grid_sha256",
        "rows",
        "receipt_sha256",
    }
    recipes._object(receipt, allowed=keys, required=keys, location="report_receipt")
    if (
        type(receipt["version"]) is not int
        or receipt["version"] != VERSION
        or type(receipt["recipe_engine_version"]) is not int
        or receipt["recipe_engine_version"] != recipes.RECIPE_ENGINE_VERSION
    ):
        raise ValidationError("Unsupported report receipt/recipe engine version")
    for name in ("source_content_sha256", "output_grid_sha256", "receipt_sha256"):
        if not isinstance(receipt[name], str) or not re.fullmatch(
            r"[0-9a-f]{64}", receipt[name]
        ):
            raise ValidationError(f"Invalid receipt {name}")
    if (
        evidence.digest({k: v for k, v in receipt.items() if k != "receipt_sha256"})
        != receipt["receipt_sha256"]
    ):
        raise ValidationError("Report receipt integrity check failed")
    normalized = recipes._normalize_recipe(receipt["recipe"])
    if normalized != receipt["recipe"]:
        raise ValidationError("Receipt recipe is not normalized")
    rows = receipt["rows"]
    width = len(normalized["group_by"]) + len(normalized["metrics"])
    if not isinstance(rows, list) or any(
        not isinstance(row, list)
        or len(row) != width
        or any(
            value is not None and type(value) not in (str, bool, int, float)
            for value in row
        )
        or all(value is None for value in row)
        for row in rows
    ):
        raise ValidationError("Invalid receipt result rows")
    _row_fingerprints(rows, len(normalized["group_by"]))
    return receipt, normalized


def check_report_receipt(
    filepath: str, report_receipt: dict, sample_limit: int = 3
) -> dict:
    evidence.sample_size(sample_limit)
    receipt, requested = _normalize(report_receipt)
    compiled = requested
    contract = receipt["workbook_contract"]
    with evidence.capture([filepath]) as [(path, sha)]:
        with workbook.safe_workbook(str(path)) as wb:
            if contract is not None:
                from .contracts import compile_contract_recipe, normalize_contract

                contract = normalize_contract(contract)
                compiled, _ = compile_contract_recipe(
                    wb, requested, contract, sample_limit
                )
            source = recipes._load_table(wb, compiled)
            result = recipes._aggregate(source, compiled)
            verification = recipes._verify_recipe(
                source, compiled, result, sample_limit
            )
            if not verification["passed"]:
                raise PreconditionFailedError(
                    "Current source fails the receipt's recipe checks",
                    code="receipt_source_invalid",
                )
            current = _row_fingerprints(result["rows"], len(requested["group_by"]))
            output = requested["output"]
            c1, r1, c2, r2 = changeset._range_bounds(
                output["range_ref"], location="receipt output"
            )
            expected_headers = [
                *requested["group_by"],
                *[m["as"] for m in requested["metrics"]],
            ]
            width = len(expected_headers)
            stored = None
            stored_rows = None
            grid_sha = None
            if output["sheet_name"] in wb.sheetnames:
                ws = workbook.require_worksheet(
                    wb,
                    output["sheet_name"],
                    error_cls=ValidationError,
                    operation="receipt output",
                )
                grid = list(
                    ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2)
                )
                grid_sha = evidence.digest(
                    [[_atom(c.value, c.data_type) for c in row] for row in grid]
                )
                if [c.value for c in grid[0][:width]] == expected_headers:
                    rows = []
                    valid = True
                    for cells in grid[1:]:
                        if any(c.value is not None for c in cells[width:]):
                            valid = False
                        used = cells[:width]
                        if all(c.value is None for c in used):
                            continue
                        if any(
                            c.data_type in ("f", "e")
                            or not (
                                c.value is None
                                or type(c.value) in (str, bool, int, float)
                            )
                            for c in used
                        ):
                            valid = False
                            continue
                        rows.append([c.value for c in used])
                    if any(c.value is not None for c in grid[0][width:]):
                        valid = False
                    if valid:
                        try:
                            stored = _row_fingerprints(rows, len(requested["group_by"]))
                            stored_rows = rows
                        except ValidationError:
                            stored = None
        group_count = len(requested["group_by"])
        old = _row_fingerprints(receipt["rows"], group_count)
        source_changed = (
            _source_digest(source, compiled) != receipt["source_content_sha256"]
        )
        output_changed = grid_sha != receipt["output_grid_sha256"]
        result_changed = current != old
        stored_matches = stored == current
        status = (
            "current"
            if stored_matches
            else "conflict"
            if output_changed or not result_changed
            else "refresh_needed"
        )
        changed = [
            key
            for key in sorted(set(old) | set(current))
            if old.get(key) != current.get(key)
        ]

        def row_map(rows):
            return {
                evidence.digest([_atom(v) for v in row[:group_count]]): row
                for row in rows
            }

        previous_values = row_map(receipt["rows"])
        current_values = row_map(result["rows"])
        stored_values = row_map(stored_rows or [])
        relevant = [
            key
            for key in sorted(set(old) | set(current) | set(stored or {}))
            if old.get(key) != current.get(key)
            or stored is not None
            and stored.get(key) != current.get(key)
        ]
        examples = []
        aliases = [m["as"] for m in requested["metrics"]]
        for key in relevant[:sample_limit]:
            label_row = current_values.get(
                key, previous_values.get(key, stored_values.get(key))
            )
            item = {"group": dict(zip(requested["group_by"], label_row[:group_count]))}
            for name, values in (
                ("previous", previous_values),
                ("computed", current_values),
                ("stored", stored_values),
            ):
                item[name] = (
                    dict(zip(aliases, values[key][group_count:]))
                    if key in values
                    else None
                )
            examples.append(item)
        return evidence.bounded(
            {
                "version": VERSION,
                "read_only": True,
                "status": status,
                "source_sha256": sha,
                "source_changed": source_changed,
                "source_change_scope": "referenced fields and all source rows in worksheet order, excluding declared totals",
                "output_changed": output_changed,
                "result_changed": result_changed,
                "stored_result_matches_current": stored_matches,
                "changed_group_count": len(changed),
                "changed_group_hashes": changed[:sample_limit],
                "groups_truncated": len(changed) > sample_limit,
                "comparison_sample": examples,
                "comparison_group_count": len(relevant),
                "comparison_truncated": len(relevant) > sample_limit,
                "stored_output_parseable": stored is not None,
                "next_action": "No refresh needed for these computed values"
                if status == "current"
                else "Obtain an owner decision before overwriting changed report values"
                if status == "conflict"
                else "Request a fresh recipe preview; this receipt is not write permission",
                "scope": "value fingerprints under the receipt's exact definition, not styles, authorship, business truth or authentication",
            }
        )
