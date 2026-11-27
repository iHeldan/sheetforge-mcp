"""Client-owned report fingerprints; no signatures, history store or write authority."""

import re

from . import changeset, recipes, workbook
from . import reporting_evidence as evidence
from .exceptions import PreconditionFailedError, ResponseTooLargeError, ValidationError

VERSION = 3
MAX_RECEIPT_CHARS = 6_000
# Checking accepts bounded JSON notation growth; issuance/write budgets stay fixed.
MAX_RECEIPT_INPUT_CHARS = 24_000


def _atom(value, data_type=None):
    if data_type in ("f", "e"):
        return [data_type, value if isinstance(value, str) else type(value).__name__]
    if type(value) in (int, float):
        recipes._scalar(value, location="receipt number")
        numerator, denominator = (
            (value, 1) if type(value) is int else value.as_integer_ratio()
        )
        return ["number", str(numerator), str(denominator)]
    if value is None or type(value) in (str, bool):
        return [type(value).__name__, value]
    # Receipts start from literal JSON values only. A later unsupported cell
    # must remain a mismatch rather than crash or masquerade as a literal.
    return ["unsupported", type(value).__name__]


def _legacy_atom(value, data_type=None):
    if data_type not in ("f", "e") and type(value) in (int, float):
        return ["number", evidence.text(evidence.number(value))]
    return _atom(value, data_type)


def _check_legacy_numbers(value):
    if isinstance(value, dict):
        for item in value.values():
            _check_legacy_numbers(item)
    elif isinstance(value, list):
        for item in value:
            _check_legacy_numbers(item)
    elif type(value) in (int, float):
        # Old decimal-string hashes cannot distinguish an inexact large integer
        # from the float whose shortest decimal spelling denotes that integer.
        if type(value) is int:
            try:
                ambiguous = float(value) != value
            except OverflowError:
                ambiguous = True
        else:
            decimal = evidence.number(value)
            ambiguous = decimal.denominator == 1 and value != decimal.numerator
        if ambiguous:
            raise ValidationError(
                "Legacy report receipt has ambiguous numeric identity; obtain a new "
                "receipt through an authorized Recipe preview and verified commit",
                code="legacy_receipt_numeric_ambiguity",
                details={"requires_new_receipt": True, "automatic_upgrade": False},
                suggested_next_tool="run_workbook_recipe",
            )


def _checked_legacy_atom(value, data_type=None):
    if data_type not in ("f", "e"):
        _check_legacy_numbers(value)
    return _legacy_atom(value, data_type)


def _row_fingerprints(rows, group_count):
    indexed = {}
    for row in rows:
        key = evidence.digest([_atom(v) for v in row[:group_count]])
        if key in indexed:
            raise ValidationError("Output has duplicate group identities")
        indexed[key] = [evidence.digest(_atom(v)) for v in row[group_count:]]
    return indexed


def _receipt_digest(value, *, atom=_atom):
    # Tag every node, not just numbers: a JSON array ["number", "1"] must
    # never collide with the numeric scalar 1. Do not alter other tool tokens.
    def canonical(node):
        if isinstance(node, dict):
            return ["object", [[key, canonical(node[key])] for key in sorted(node)]]
        if isinstance(node, list):
            return ["array", [canonical(item) for item in node]]
        return atom(node)

    return evidence.digest(canonical(value))


def _output_headers(requested, contract, result):
    if contract is not None:
        return [*requested["group_by"], *[m["as"] for m in requested["metrics"]]]
    return result["headers"]


def _source_digest(source, recipe, *, atom=_atom):
    fields = recipes._source_fields(recipe)
    indexes = [source["headers"].index(f) for f in fields]
    return evidence.digest(
        {
            "fields": fields,
            "rows": [[atom(row[i]) for i in indexes] for row in source["rows"]],
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
        "output_headers": result["headers"],
    }
    receipt["receipt_sha256"] = _receipt_digest(receipt)
    return evidence.bounded(receipt, MAX_RECEIPT_CHARS)


def _normalize(receipt):
    receipt = changeset._json_copy(receipt, location="report_receipt")
    try:
        evidence.bounded(receipt, MAX_RECEIPT_INPUT_CHARS)
    except ResponseTooLargeError as exc:
        raise ResponseTooLargeError(
            "Report receipt exceeds its checking input budget",
            estimated_size=exc.estimated_size,
            limit=MAX_RECEIPT_INPUT_CHARS,
            hints=[
                "Use the original issued receipt JSON without adding data",
                "If unavailable, obtain a new receipt through an authorized Recipe preview and verified commit",
            ],
        ) from exc
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
    if (
        isinstance(receipt, dict)
        and type(receipt.get("version")) is int
        and receipt["version"] in (2, VERSION)
    ):
        keys.add("output_headers")
    recipes._object(receipt, allowed=keys, required=keys, location="report_receipt")
    if (
        type(receipt["version"]) is not int
        or receipt["version"] not in (1, 2, VERSION)
        or type(receipt["recipe_engine_version"]) is not int
        or receipt["recipe_engine_version"] != recipes.RECIPE_ENGINE_VERSION
    ):
        raise ValidationError("Unsupported report receipt/recipe engine version")
    for name in ("source_content_sha256", "output_grid_sha256", "receipt_sha256"):
        if not isinstance(receipt[name], str) or not re.fullmatch(
            r"[0-9a-f]{64}", receipt[name]
        ):
            raise ValidationError(f"Invalid receipt {name}")

    def checksum(value):
        if receipt["version"] == 1:
            return evidence.digest(value)
        return _receipt_digest(
            value, atom=_legacy_atom if receipt["version"] == 2 else _atom
        )

    if (
        checksum({k: v for k, v in receipt.items() if k != "receipt_sha256"})
        != receipt["receipt_sha256"]
    ):
        raise ValidationError("Report receipt integrity check failed")
    if receipt["version"] < VERSION:
        _check_legacy_numbers(receipt)
    normalized = recipes._normalize_recipe(receipt["recipe"])
    if normalized != receipt["recipe"]:
        raise ValidationError("Receipt recipe is not normalized")
    rows = receipt["rows"]
    width = len(normalized["group_by"]) + len(normalized["metrics"])
    if receipt["version"] >= 2:
        headers = receipt["output_headers"]
        if not isinstance(headers, list) or len(headers) != width:
            raise ValidationError("Invalid receipt output headers")
        for header in headers:
            recipes._name(header, "receipt output header")
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
    fingerprint_atom = _checked_legacy_atom if receipt["version"] < VERSION else _atom
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
            source_sha = _source_digest(source, compiled, atom=fingerprint_atom)
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
            expected_headers = _output_headers(requested, contract, result)
            if (
                receipt["version"] >= 2
                and receipt["output_headers"] != expected_headers
            ):
                raise ValidationError("Receipt output headers do not match its recipe")
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
                    [
                        [fingerprint_atom(c.value, c.data_type) for c in row]
                        for row in grid
                    ]
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
        source_changed = source_sha != receipt["source_content_sha256"]
        output_changed = grid_sha != receipt["output_grid_sha256"]
        result_changed = current != old
        stored_matches = stored == current
        values_changed = stored != old if stored is not None else None
        positions_only = bool(output_changed and stored is not None and stored == old)
        status = (
            "current"
            if stored_matches
            else "conflict"
            if output_changed or not result_changed
            else "refresh_needed"
        )
        conflict_reason = None
        if status == "conflict":
            conflict_reason = (
                "output_structure_changed"
                if stored is None
                else "output_values_changed"
                if values_changed
                else "output_positions_changed"
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
        aliases = expected_headers[group_count:]
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
                "version": receipt["version"],
                "read_only": True,
                "status": status,
                "source_sha256": sha,
                "source_changed": source_changed,
                "source_change_scope": "referenced fields and all source rows in worksheet order, excluding declared totals",
                "output_changed": output_changed,
                "output_values_changed": values_changed,
                "output_positions_only_changed": positions_only,
                "conflict_reason": conflict_reason,
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
