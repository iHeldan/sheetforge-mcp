import logging
import re
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from .exceptions import WorkbookError

logger = logging.getLogger(__name__)

_STRUCTURED_REFERENCE_FLAG_RANGE_RE = re.compile(
    r"^\[\[(?P<flag>#[^,\]]+)\],\[(?P<start>[^\]]+)\]:\[(?P<end>[^\]]+)\]\]$"
)
_STRUCTURED_REFERENCE_FLAG_COLUMN_RE = re.compile(
    r"^\[\[(?P<flag>#[^,\]]+)\],\[(?P<column>[^\]]+)\]\]$"
)
_STRUCTURED_REFERENCE_COLUMN_RANGE_RE = re.compile(
    r"^\[\[(?P<start>[^\]]+)\]:\[(?P<end>[^\]]+)\]\]$"
)
_STRUCTURED_REFERENCE_FLAG_ONLY_RE = re.compile(r"^\[(?P<flag>#[^\]]+)\]$")
_STRUCTURED_REFERENCE_THIS_ROW_COLUMN_RE = re.compile(r"^\[@(?P<column>[^\]]+)\]$")
_STRUCTURED_REFERENCE_COLUMN_RE = re.compile(r"^\[(?P<column>[^\]]+)\]$")


def _get_sheet_usage(ws) -> tuple[int, int, str | None, bool]:
    """Return user-facing sheet dimensions instead of openpyxl's default 1x1 empty sheet."""
    is_empty = ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None
    if is_empty:
        return 0, 0, None, True
    return ws.max_row, ws.max_column, f"A-{get_column_letter(ws.max_column)}", False


def _get_used_range(ws) -> str | None:
    rows, columns, _, is_empty = _get_sheet_usage(ws)
    if is_empty or rows == 0 or columns == 0:
        return None
    return f"A1:{get_column_letter(columns)}{rows}"


def _bounds_to_range(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def _bounds_intersect(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> bool:
    return not (
        first[2] < second[0]
        or second[2] < first[0]
        or first[3] < second[1]
        or second[3] < first[1]
    )


def _intersection_bounds(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> tuple[int, int, int, int] | None:
    if not _bounds_intersect(first, second):
        return None
    return (
        max(first[0], second[0]),
        max(first[1], second[1]),
        min(first[2], second[2]),
        min(first[3], second[3]),
    )


def _parse_range_reference(
    range_ref: str,
    *,
    worksheet: Worksheet | None = None,
    expected_sheet: str | None = None,
    error_cls: type[Exception] = WorkbookError,
) -> tuple[tuple[int, int, int, int], str]:
    cleaned = str(range_ref).strip()
    if not cleaned:
        raise error_cls("Range reference is required")

    if "!" in cleaned:
        range_sheet, cleaned = cleaned.rsplit("!", 1)
        normalized_sheet = range_sheet.strip().strip("'")
        if expected_sheet is not None and normalized_sheet != expected_sheet:
            raise error_cls(
                f"Range '{range_ref}' refers to sheet '{normalized_sheet}', expected '{expected_sheet}'"
            )

    cleaned = cleaned.replace("$", "")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except ValueError as exc:
        raise error_cls(f"Invalid range reference: {range_ref}") from exc

    if None in (min_col, min_row, max_col, max_row):
        if worksheet is None:
            raise error_cls(f"Range reference requires worksheet bounds: {range_ref}")
        min_col = 1 if min_col is None else min_col
        min_row = 1 if min_row is None else min_row
        max_col = worksheet.max_column if max_col is None else max_col
        max_row = worksheet.max_row if max_row is None else max_row

    return (min_row, min_col, max_row, max_col), _bounds_to_range(min_row, min_col, max_row, max_col)


def _iter_range_references(
    range_ref: Any,
    *,
    worksheet: Worksheet | None = None,
    expected_sheet: str | None = None,
) -> list[tuple[tuple[int, int, int, int], str]]:
    if not range_ref:
        return []

    if hasattr(range_ref, "ranges"):
        parts = [str(item) for item in range_ref.ranges]
    elif isinstance(range_ref, (list, tuple)):
        parts = range_ref
    else:
        parts = str(range_ref).split(",")
    references: list[tuple[tuple[int, int, int, int], str]] = []
    for part in parts:
        cleaned = str(part).strip()
        if not cleaned:
            continue

        if "!" in cleaned:
            range_sheet, local_range = cleaned.rsplit("!", 1)
            normalized_sheet = range_sheet.strip().strip("'")
            if expected_sheet is not None and normalized_sheet != expected_sheet:
                continue
        else:
            local_range = cleaned

        local_range = local_range.replace("$", "")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(local_range)
        except ValueError:
            continue
        if None in (min_col, min_row, max_col, max_row):
            if worksheet is None:
                continue
            min_col = 1 if min_col is None else min_col
            min_row = 1 if min_row is None else min_row
            max_col = worksheet.max_column if max_col is None else max_col
            max_row = worksheet.max_row if max_row is None else max_row
        references.append(
            ((min_row, min_col, max_row, max_col), _bounds_to_range(min_row, min_col, max_row, max_col))
        )
    return references


def _cell_is_within_bounds(
    sheet_name: str,
    row_index: int,
    column_index: int,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> bool:
    return (
        sheet_name == target_sheet
        and target_bounds[0] <= row_index <= target_bounds[2]
        and target_bounds[1] <= column_index <= target_bounds[3]
    )


def _defined_name_local_sheet(
    wb: Any,
    defined_name: Any,
) -> str | None:
    local_sheet_id = getattr(defined_name, "localSheetId", None)
    if local_sheet_id is None:
        return None
    try:
        return wb.sheetnames[local_sheet_id]
    except Exception:
        return None


def _iter_defined_name_entries(wb: Any):
    seen: set[tuple[str, str | None]] = set()

    for name, defined_name in wb.defined_names.items():
        local_sheet = _defined_name_local_sheet(wb, defined_name)
        entry_key = (name, local_sheet)
        if entry_key in seen:
            continue
        seen.add(entry_key)
        yield name, defined_name, local_sheet

    for ws in getattr(wb, "worksheets", []):
        for name, defined_name in ws.defined_names.items():
            entry_key = (name, ws.title)
            if entry_key in seen:
                continue
            seen.add(entry_key)
            yield name, defined_name, ws.title


def _resolve_defined_name(
    wb: Any,
    *,
    name: str,
    formula_sheet_name: str,
    scope_sheet_name: str | None = None,
) -> tuple[Any, str | None] | None:
    if scope_sheet_name is not None and scope_sheet_name in wb.sheetnames:
        scoped_ws = wb[scope_sheet_name]
        scoped_name = scoped_ws.defined_names.get(name)
        if scoped_name is not None:
            return scoped_name, scope_sheet_name

    if scope_sheet_name is not None:
        scoped_name = wb.defined_names.get(name)
        if scoped_name is None:
            return None
        local_sheet = _defined_name_local_sheet(wb, scoped_name)
        if local_sheet != scope_sheet_name:
            return None
        return scoped_name, local_sheet

    if formula_sheet_name in wb.sheetnames:
        local_ws = wb[formula_sheet_name]
        local_name = local_ws.defined_names.get(name)
        if local_name is not None:
            return local_name, formula_sheet_name

    workbook_name = wb.defined_names.get(name)
    if workbook_name is None:
        return None

    local_sheet = _defined_name_local_sheet(wb, workbook_name)
    if local_sheet is not None and local_sheet != formula_sheet_name:
        return None

    return workbook_name, local_sheet


def _resolve_named_range_references(
    wb: Any,
    *,
    name: str,
    formula_sheet_name: str,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
    scope_sheet_name: str | None = None,
) -> list[dict[str, Any]]:
    resolved = _resolve_defined_name(
        wb,
        name=name,
        formula_sheet_name=formula_sheet_name,
        scope_sheet_name=scope_sheet_name,
    )
    if resolved is None:
        return []

    defined_name, _ = resolved

    matches: list[dict[str, Any]] = []
    try:
        destinations = list(defined_name.destinations)
    except Exception:
        destinations = []

    for destination_sheet_name, destination_range in destinations:
        if destination_sheet_name != target_sheet or destination_sheet_name not in wb.sheetnames:
            continue

        destination_ws = wb[destination_sheet_name]
        if _sheet_type(destination_ws) == "chartsheet":
            continue

        for destination_bounds, destination_ref in _iter_range_references(
            destination_range,
            worksheet=destination_ws,
            expected_sheet=target_sheet,
        ):
            intersection = _intersection_bounds(target_bounds, destination_bounds)
            if intersection is None:
                continue
            matches.append(
                {
                    "reference": f"{destination_sheet_name}!{destination_ref}",
                    "intersection_range": _bounds_to_range(*intersection),
                    "via_named_range": defined_name.name,
                }
            )

    return matches


def _find_table_reference(
    wb: Any,
    *,
    table_name: str,
    sheet_name: str | None = None,
) -> tuple[str, Worksheet, Any] | None:
    sheet_names = [sheet_name] if sheet_name is not None else list(wb.sheetnames)
    for current_sheet_name in sheet_names:
        if current_sheet_name not in wb.sheetnames:
            continue
        ws = wb[current_sheet_name]
        if _sheet_type(ws) == "chartsheet":
            continue
        for table in ws.tables.values():
            if table.displayName == table_name:
                return current_sheet_name, ws, table
    return None


def _table_column_lookup(ws: Worksheet, table: Any) -> dict[str, int]:
    min_col, min_row, max_col, _ = range_boundaries(table.ref)
    column_names = list(getattr(table, "column_names", []) or [])
    if len(column_names) == (max_col - min_col + 1):
        return {
            str(column_name): min_col + offset
            for offset, column_name in enumerate(column_names)
        }

    return {
        str(ws.cell(row=min_row, column=column_index).value): column_index
        for column_index in range(min_col, max_col + 1)
    }


def _normalize_structured_flag(flag: str | None) -> str:
    if not flag:
        return "data"
    normalized = flag.strip().lower()
    if normalized.startswith("#"):
        normalized = normalized[1:]
    return normalized.replace(" ", "")


def _parse_structured_reference(
    local_reference: str,
) -> tuple[str, dict[str, Any]] | None:
    if "[" not in local_reference or not local_reference.endswith("]"):
        return None

    table_name, spec = local_reference.split("[", 1)
    table_name = table_name.strip()
    if not table_name:
        return None
    spec = f"[{spec}"

    match = _STRUCTURED_REFERENCE_FLAG_RANGE_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": _normalize_structured_flag(match.group("flag")),
            "start_column": match.group("start"),
            "end_column": match.group("end"),
        }

    match = _STRUCTURED_REFERENCE_FLAG_COLUMN_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": _normalize_structured_flag(match.group("flag")),
            "start_column": match.group("column"),
            "end_column": match.group("column"),
        }

    match = _STRUCTURED_REFERENCE_COLUMN_RANGE_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": "data",
            "start_column": match.group("start"),
            "end_column": match.group("end"),
        }

    match = _STRUCTURED_REFERENCE_FLAG_ONLY_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": _normalize_structured_flag(match.group("flag")),
            "start_column": None,
            "end_column": None,
        }

    match = _STRUCTURED_REFERENCE_THIS_ROW_COLUMN_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": "thisrow",
            "start_column": match.group("column"),
            "end_column": match.group("column"),
        }

    match = _STRUCTURED_REFERENCE_COLUMN_RE.fullmatch(spec)
    if match:
        return table_name, {
            "row_selector": "data",
            "start_column": match.group("column"),
            "end_column": match.group("column"),
        }

    return None


def _resolve_table_structured_reference(
    wb: Any,
    *,
    local_reference: str,
    formula_sheet_name: str,
    formula_row: int,
    scope_sheet_name: str | None,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    parsed_reference = _parse_structured_reference(local_reference)
    if parsed_reference is None:
        return []

    table_name, reference_parts = parsed_reference
    table_match = _find_table_reference(wb, table_name=table_name, sheet_name=scope_sheet_name)
    if table_match is None:
        return []

    table_sheet_name, table_ws, table = table_match
    if table_sheet_name != target_sheet:
        return []

    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    header_row_count = int(table.headerRowCount or 0)
    totals_row_count = int(table.totalsRowCount or 0)
    data_start_row = min_row + header_row_count
    data_end_row = max_row - totals_row_count

    row_selector = reference_parts["row_selector"]
    if row_selector == "all":
        row_start, row_end = min_row, max_row
    elif row_selector == "data":
        row_start, row_end = data_start_row, data_end_row
    elif row_selector == "headers":
        if header_row_count < 1:
            return []
        row_start, row_end = min_row, min_row + header_row_count - 1
    elif row_selector == "totals":
        if totals_row_count < 1:
            return []
        row_start, row_end = max_row - totals_row_count + 1, max_row
    elif row_selector == "thisrow":
        if table_sheet_name != formula_sheet_name or not (data_start_row <= formula_row <= data_end_row):
            return []
        row_start = row_end = formula_row
    else:
        return []

    if row_start > row_end:
        return []

    start_column_name = reference_parts["start_column"]
    end_column_name = reference_parts["end_column"]
    if start_column_name is None:
        column_start, column_end = min_col, max_col
    else:
        column_lookup = _table_column_lookup(table_ws, table)
        if start_column_name not in column_lookup or end_column_name not in column_lookup:
            return []
        column_start = column_lookup[start_column_name]
        column_end = column_lookup[end_column_name]
        if column_start > column_end:
            column_start, column_end = column_end, column_start

    reference_bounds = (row_start, column_start, row_end, column_end)
    intersection = _intersection_bounds(target_bounds, reference_bounds)
    if intersection is None:
        return []

    return [
        {
            "reference": f"{table_sheet_name}!{_bounds_to_range(*reference_bounds)}",
            "intersection_range": _bounds_to_range(*intersection),
            "via_table": table.displayName,
            "structured_reference": local_reference,
        }
    ]


def _resolve_formula_token_references(
    wb: Any,
    *,
    token_value: str,
    formula_sheet_name: str,
    formula_row: int,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    reference_scope_sheet = None
    local_reference = token_value
    if "!" in token_value:
        range_sheet, local_reference = token_value.rsplit("!", 1)
        reference_scope_sheet = range_sheet.strip().strip("'")

    table_matches = _resolve_table_structured_reference(
        wb,
        local_reference=local_reference,
        formula_sheet_name=formula_sheet_name,
        formula_row=formula_row,
        scope_sheet_name=reference_scope_sheet,
        target_sheet=target_sheet,
        target_bounds=target_bounds,
    )
    if table_matches:
        return table_matches

    reference_sheet_name = formula_sheet_name if reference_scope_sheet is None else reference_scope_sheet
    if reference_sheet_name == target_sheet and reference_sheet_name in wb.sheetnames:
        reference_ws = wb[reference_sheet_name]
        if _sheet_type(reference_ws) != "chartsheet":
            try:
                reference_bounds, normalized_reference = _parse_range_reference(
                    local_reference,
                    worksheet=reference_ws,
                    error_cls=WorkbookError,
                )
            except WorkbookError:
                pass
            else:
                intersection = _intersection_bounds(target_bounds, reference_bounds)
                if intersection is not None:
                    return [
                        {
                            "reference": f"{reference_sheet_name}!{normalized_reference}",
                            "intersection_range": _bounds_to_range(*intersection),
                        }
                    ]

    if "[" in local_reference:
        return []

    return _resolve_named_range_references(
        wb,
        name=local_reference,
        formula_sheet_name=formula_sheet_name,
        target_sheet=target_sheet,
        target_bounds=target_bounds,
        scope_sheet_name=reference_scope_sheet,
    )


def _extract_formula_dependencies(
    wb: Any,
    *,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    def _dedupe_reference_matches(references: list[dict[str, Any]]) -> list[dict[str, Any]]:
        deduped: list[dict[str, Any]] = []
        seen: set[tuple[tuple[str, str], ...]] = set()
        for reference in references:
            key = tuple(sorted((str(item_key), str(item_value)) for item_key, item_value in reference.items()))
            if key in seen:
                continue
            seen.add(key)
            deduped.append(reference)
        return deduped

    def _dedupe_formula_cells(cells: list[dict[str, Any]]) -> list[dict[str, Any]]:
        deduped: list[dict[str, Any]] = []
        seen: set[tuple[str, str]] = set()
        for cell in cells:
            key = (cell["sheet_name"], cell["cell"])
            if key in seen:
                continue
            seen.add(key)
            deduped.append(cell)
        return deduped

    formula_entries: list[dict[str, Any]] = []

    for formula_sheet_name in wb.sheetnames:
        formula_ws = wb[formula_sheet_name]
        if _sheet_type(formula_ws) == "chartsheet":
            continue

        for row in formula_ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                if _cell_is_within_bounds(
                    formula_sheet_name,
                    cell.row,
                    cell.column,
                    target_sheet,
                    target_bounds,
                ):
                    continue

                matched_references: list[dict[str, Any]] = []
                try:
                    tokenizer = Tokenizer(cell.value)
                except Exception:
                    continue

                token_values: list[str] = []
                for token in tokenizer.items:
                    if token.type != "OPERAND" or token.subtype != "RANGE":
                        continue

                    token_value = str(token.value).strip()
                    if not token_value:
                        continue

                    token_values.append(token_value)
                    matched_references.extend(
                        _resolve_formula_token_references(
                            wb,
                            token_value=token_value,
                            formula_sheet_name=formula_sheet_name,
                            formula_row=cell.row,
                            target_sheet=target_sheet,
                            target_bounds=target_bounds,
                        )
                    )

                formula_entries.append(
                    {
                        "sheet_name": formula_sheet_name,
                        "cell": cell.coordinate,
                        "row": cell.row,
                        "column": cell.column,
                        "formula": cell.value,
                        "token_values": token_values,
                        "direct_references": _dedupe_reference_matches(matched_references),
                    }
                )

    dependencies: list[dict[str, Any]] = []
    formula_entry_lookup = {
        (entry["sheet_name"], entry["cell"]): entry for entry in formula_entries
    }

    direct_frontier: set[tuple[str, str]] = set()
    for entry in formula_entries:
        if not entry["direct_references"]:
            continue
        direct_frontier.add((entry["sheet_name"], entry["cell"]))
        dependencies.append(
            {
                "sheet_name": entry["sheet_name"],
                "cell": entry["cell"],
                "formula": entry["formula"],
                "references": entry["direct_references"],
                "dependency_depth": 1,
                "dependency_type": "direct",
            }
        )

    discovered = set(direct_frontier)
    frontier = set(direct_frontier)
    depth = 2

    while frontier:
        frontier_entries = [formula_entry_lookup[key] for key in frontier]
        next_frontier: set[tuple[str, str]] = set()

        for entry in formula_entries:
            entry_key = (entry["sheet_name"], entry["cell"])
            if entry_key in discovered:
                continue

            matched_references: list[dict[str, Any]] = []
            transitive_via: list[dict[str, Any]] = []
            for predecessor in frontier_entries:
                predecessor_bounds = (
                    predecessor["row"],
                    predecessor["column"],
                    predecessor["row"],
                    predecessor["column"],
                )
                predecessor_matches: list[dict[str, Any]] = []
                for token_value in entry["token_values"]:
                    predecessor_matches.extend(
                        _resolve_formula_token_references(
                            wb,
                            token_value=token_value,
                            formula_sheet_name=entry["sheet_name"],
                            formula_row=entry["row"],
                            target_sheet=predecessor["sheet_name"],
                            target_bounds=predecessor_bounds,
                        )
                    )

                if predecessor_matches:
                    matched_references.extend(predecessor_matches)
                    transitive_via.append(
                        {
                            "sheet_name": predecessor["sheet_name"],
                            "cell": predecessor["cell"],
                        }
                    )

            matched_references = _dedupe_reference_matches(matched_references)
            transitive_via = _dedupe_formula_cells(transitive_via)
            if not matched_references:
                continue

            dependencies.append(
                {
                    "sheet_name": entry["sheet_name"],
                    "cell": entry["cell"],
                    "formula": entry["formula"],
                    "references": matched_references,
                    "dependency_depth": depth,
                    "dependency_type": "transitive",
                    "transitive_via": transitive_via,
                }
            )
            discovered.add(entry_key)
            next_frontier.add(entry_key)

        frontier = next_frontier
        depth += 1

    return dependencies


def _extract_reference_matches_from_formula_text(
    wb: Any,
    *,
    formula_text: Any,
    formula_sheet_name: str,
    formula_row: int,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    if formula_text is None:
        return []

    normalized_formula = str(formula_text).strip()
    if not normalized_formula:
        return []
    if not normalized_formula.startswith("="):
        normalized_formula = f"={normalized_formula}"

    try:
        tokenizer = Tokenizer(normalized_formula)
    except Exception:
        return []

    matched_references: list[dict[str, Any]] = []
    for token in tokenizer.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue

        token_value = str(token.value).strip()
        if not token_value:
            continue
        matched_references.extend(
            _resolve_formula_token_references(
                wb,
                token_value=token_value,
                formula_sheet_name=formula_sheet_name,
                formula_row=formula_row,
                target_sheet=target_sheet,
                target_bounds=target_bounds,
            )
        )

    return matched_references


def _extract_validation_overlaps(
    ws: Worksheet,
    *,
    sheet_name: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    overlaps: list[dict[str, Any]] = []

    for validation in getattr(getattr(ws, "data_validations", None), "dataValidation", []):
        intersection_ranges: list[str] = []
        for validation_bounds, validation_ref in _iter_range_references(
            validation.sqref,
            worksheet=ws,
            expected_sheet=sheet_name,
        ):
            intersection = _intersection_bounds(target_bounds, validation_bounds)
            if intersection is None:
                continue
            intersection_ranges.append(_bounds_to_range(*intersection))

        if not intersection_ranges:
            continue

        overlaps.append(
            {
                "applies_to": str(validation.sqref),
                "intersection_ranges": intersection_ranges,
                "validation_type": validation.type,
                "operator": validation.operator or None,
                "formula1": validation.formula1 or None,
                "formula2": validation.formula2 or None,
            }
        )

    return overlaps


def _extract_validation_dependencies(
    wb: Any,
    *,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    dependencies: list[dict[str, Any]] = []

    for validation_sheet_name in wb.sheetnames:
        validation_ws = wb[validation_sheet_name]
        if _sheet_type(validation_ws) == "chartsheet":
            continue

        for validation in getattr(getattr(validation_ws, "data_validations", None), "dataValidation", []):
            applies_to_refs = _iter_range_references(
                validation.sqref,
                worksheet=validation_ws,
                expected_sheet=validation_sheet_name,
            )
            formula_row = applies_to_refs[0][0][0] if applies_to_refs else 1

            matched_references: list[dict[str, Any]] = []
            matched_references.extend(
                _extract_reference_matches_from_formula_text(
                    wb,
                    formula_text=validation.formula1,
                    formula_sheet_name=validation_sheet_name,
                    formula_row=formula_row,
                    target_sheet=target_sheet,
                    target_bounds=target_bounds,
                )
            )
            matched_references.extend(
                _extract_reference_matches_from_formula_text(
                    wb,
                    formula_text=validation.formula2,
                    formula_sheet_name=validation_sheet_name,
                    formula_row=formula_row,
                    target_sheet=target_sheet,
                    target_bounds=target_bounds,
                )
            )

            if not matched_references:
                continue

            dependencies.append(
                {
                    "sheet_name": validation_sheet_name,
                    "applies_to": str(validation.sqref),
                    "validation_type": validation.type,
                    "operator": validation.operator or None,
                    "formula1": validation.formula1 or None,
                    "formula2": validation.formula2 or None,
                    "references": matched_references,
                }
            )

    return dependencies


def _extract_conditional_format_overlaps(
    ws: Worksheet,
    *,
    sheet_name: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    overlaps: list[dict[str, Any]] = []

    for conditional_format, rules in getattr(ws.conditional_formatting, "_cf_rules", {}).items():
        intersection_ranges: list[str] = []
        for format_bounds, _ in _iter_range_references(
            conditional_format.sqref,
            worksheet=ws,
            expected_sheet=sheet_name,
        ):
            intersection = _intersection_bounds(target_bounds, format_bounds)
            if intersection is None:
                continue
            intersection_ranges.append(_bounds_to_range(*intersection))

        if not intersection_ranges:
            continue

        for rule in rules:
            overlaps.append(
                {
                    "applies_to": str(conditional_format.sqref),
                    "intersection_ranges": intersection_ranges,
                    "rule_type": getattr(rule, "type", None),
                    "operator": getattr(rule, "operator", None),
                    "formula": list(getattr(rule, "formula", None) or []),
                }
            )

    return overlaps


def _extract_conditional_format_dependencies(
    wb: Any,
    *,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    dependencies: list[dict[str, Any]] = []

    for format_sheet_name in wb.sheetnames:
        format_ws = wb[format_sheet_name]
        if _sheet_type(format_ws) == "chartsheet":
            continue

        for conditional_format, rules in getattr(format_ws.conditional_formatting, "_cf_rules", {}).items():
            applies_to_refs = _iter_range_references(
                conditional_format.sqref,
                worksheet=format_ws,
                expected_sheet=format_sheet_name,
            )
            formula_row = applies_to_refs[0][0][0] if applies_to_refs else 1

            for rule in rules:
                formulas = list(getattr(rule, "formula", None) or [])
                matched_references: list[dict[str, Any]] = []
                for formula_text in formulas:
                    matched_references.extend(
                        _extract_reference_matches_from_formula_text(
                            wb,
                            formula_text=formula_text,
                            formula_sheet_name=format_sheet_name,
                            formula_row=formula_row,
                            target_sheet=target_sheet,
                            target_bounds=target_bounds,
                        )
                    )

                if not matched_references:
                    continue

                dependencies.append(
                    {
                        "sheet_name": format_sheet_name,
                        "applies_to": str(conditional_format.sqref),
                        "rule_type": getattr(rule, "type", None),
                        "operator": getattr(rule, "operator", None),
                        "formula": formulas,
                        "references": matched_references,
                    }
                )

    return dependencies


def _freeze_panes_value(ws) -> str | None:
    value = ws.freeze_panes
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return getattr(value, "coordinate", None)


def _sheet_type(ws: Any) -> str:
    return "chartsheet" if ws.__class__.__name__ == "Chartsheet" else "worksheet"


def require_worksheet(
    wb: Any,
    sheet_name: str,
    *,
    error_cls: type[Exception] = WorkbookError,
    operation: str = "worksheet operations",
) -> Worksheet:
    if sheet_name not in wb.sheetnames:
        raise error_cls(f"Sheet '{sheet_name}' not found")

    ws = wb[sheet_name]
    if not isinstance(ws, Worksheet):
        raise error_cls(
            f"Sheet '{sheet_name}' is a chartsheet and cannot be used for {operation}"
        )
    return ws


def first_worksheet(
    wb: Any,
    *,
    error_cls: type[Exception] = WorkbookError,
) -> tuple[str, Worksheet]:
    if not wb.sheetnames:
        raise error_cls("Workbook contains no sheets")

    worksheets = list(getattr(wb, "worksheets", []))
    if not worksheets:
        raise error_cls("Workbook contains no worksheets")

    worksheet = worksheets[0]
    return worksheet.title, worksheet


def _serialize_named_ranges(wb: Any) -> list[dict[str, Any]]:
    ranges = []
    for name, defined_name, local_sheet in _iter_defined_name_entries(wb):
        destinations = []
        try:
            destinations = [
                {
                    "sheet_name": sheet_name,
                    "range": cell_range,
                }
                for sheet_name, cell_range in defined_name.destinations
            ]
        except Exception:
            destinations = []

        ranges.append(
            {
                "name": name,
                "type": defined_name.type,
                "value": defined_name.value,
                "destinations": destinations,
                "local_sheet": local_sheet,
                "hidden": bool(getattr(defined_name, "hidden", False)),
            }
        )

    return sorted(ranges, key=lambda item: item["name"].lower())


@contextmanager
def safe_workbook(filepath: str, save: bool = False, read_only: bool = False):
    """Context manager that ensures workbook is always closed.

    Args:
        filepath: Path to Excel file.
        save: If True, save the workbook before closing.
        read_only: If True, open in read-only mode.
    """
    wb = load_workbook(filepath, read_only=read_only)
    try:
        yield wb
    except Exception:
        raise
    else:
        if save and not read_only:
            wb.save(filepath)
    finally:
        wb.close()

def create_workbook(filepath: str, sheet_name: str = "Sheet1") -> dict[str, Any]:
    """Create a new Excel workbook with optional custom sheet name"""
    try:
        wb = Workbook()
        # Rename default sheet
        if "Sheet" in wb.sheetnames:
            sheet = wb["Sheet"]
            sheet.title = sheet_name
        else:
            wb.create_sheet(sheet_name)

        path = Path(filepath)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(path))
        return {
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
            "workbook": wb
        }
    except Exception as e:
        logger.error(f"Failed to create workbook: {e}")
        raise WorkbookError(f"Failed to create workbook: {e!s}")

def get_or_create_workbook(filepath: str) -> Workbook:
    """Load an existing workbook. Raises FileNotFoundError if it doesn't exist."""
    return load_workbook(filepath)

def create_sheet(filepath: str, sheet_name: str) -> dict:
    """Create a new worksheet in the workbook if it doesn't exist."""
    try:
        with safe_workbook(filepath, save=True) as wb:
            # Check if sheet already exists
            if sheet_name in wb.sheetnames:
                raise WorkbookError(f"Sheet {sheet_name} already exists")

            # Create new sheet
            wb.create_sheet(sheet_name)
        return {"message": f"Sheet {sheet_name} created successfully"}
    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create sheet: {e}")
        raise WorkbookError(str(e))

def list_sheets(filepath: str) -> list[dict[str, Any]]:
    """List all sheets with basic metadata."""
    try:
        with safe_workbook(filepath) as wb:
            sheets = []
            for name in wb.sheetnames:
                ws = wb[name]
                if _sheet_type(ws) == "chartsheet":
                    sheets.append(
                        {
                            "name": name,
                            "sheet_type": "chartsheet",
                            "rows": 0,
                            "columns": 0,
                            "column_range": None,
                            "is_empty": len(getattr(ws, "_charts", [])) == 0,
                        }
                    )
                    continue

                rows, columns, column_range, is_empty = _get_sheet_usage(ws)
                sheets.append({
                    "name": name,
                    "sheet_type": "worksheet",
                    "rows": rows,
                    "columns": columns,
                    "column_range": column_range,
                    "is_empty": is_empty,
                })
            return sheets
    except Exception as e:
        logger.error(f"Failed to list sheets: {e}")
        raise WorkbookError(str(e))


def list_named_ranges(filepath: str) -> list[dict[str, Any]]:
    """List workbook-level and local defined names."""
    try:
        with safe_workbook(filepath) as wb:
            return _serialize_named_ranges(wb)
    except Exception as e:
        logger.error(f"Failed to list named ranges: {e}")
        raise WorkbookError(str(e))

def get_workbook_info(filepath: str, include_ranges: bool = False) -> dict[str, Any]:
    """Get metadata about workbook including sheets, ranges, etc."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            info = {
                "filename": path.name,
                "sheets": wb.sheetnames,
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime
            }

            if include_ranges:
                # Add used ranges for each sheet
                ranges = {}
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    if _sheet_type(ws) == "chartsheet":
                        continue

                    used_range = _get_used_range(ws)
                    if used_range is not None:
                        ranges[sheet_name] = used_range
                info["used_ranges"] = ranges

        return info

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get workbook info: {e}")
        raise WorkbookError(str(e))


def profile_workbook(filepath: str) -> dict[str, Any]:
    """Return a workbook-level inventory tuned for agent orientation."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            from .chart import (
                _chart_occupied_range,
                _chart_type_name,
                _extract_chart_anchor,
                _extract_chart_dimensions,
                _extract_title_text,
            )
            from .sheet import _sheet_protection_state
            from .tables import _build_table_metadata

            named_ranges = _serialize_named_ranges(wb)
            sheets: list[dict[str, Any]] = []
            total_tables = 0
            total_charts = 0

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                charts = []
                for chart_index, chart in enumerate(getattr(ws, "_charts", []), start=1):
                    width, height = _extract_chart_dimensions(chart)
                    anchor = _extract_chart_anchor(chart)
                    series = getattr(chart, "ser", None) or list(getattr(chart, "series", []))
                    chart_info = {
                        "chart_index": chart_index,
                        "chart_type": _chart_type_name(chart),
                        "anchor": anchor,
                        "title": _extract_title_text(getattr(chart, "title", None)),
                        "width": width,
                        "height": height,
                        "series_count": len(series),
                    }
                    if anchor and width and height:
                        chart_info["occupied_range"] = _chart_occupied_range(
                            ws,
                            anchor,
                            width=width,
                            height=height,
                        )
                    charts.append(chart_info)

                if _sheet_type(ws) == "chartsheet":
                    total_charts += len(charts)
                    sheets.append(
                        {
                            "name": sheet_name,
                            "sheet_type": "chartsheet",
                            "visibility": ws.sheet_state,
                            "table_count": 0,
                            "chart_count": len(charts),
                            "tables": [],
                            "charts": charts,
                        }
                    )
                    continue

                rows, columns, column_range, is_empty = _get_sheet_usage(ws)
                tables = [
                    _build_table_metadata(sheet_name, ws, table)
                    for table in ws.tables.values()
                ]

                total_tables += len(tables)
                total_charts += len(charts)

                sheets.append(
                    {
                        "name": sheet_name,
                        "sheet_type": "worksheet",
                        "rows": rows,
                        "columns": columns,
                        "column_range": column_range,
                        "used_range": _get_used_range(ws),
                        "is_empty": is_empty,
                        "visibility": ws.sheet_state,
                        "freeze_panes": _freeze_panes_value(ws),
                        "has_autofilter": bool(ws.auto_filter.ref),
                        "autofilter_range": ws.auto_filter.ref or None,
                        "print_area": ws.print_area or None,
                        "print_title_rows": ws.print_title_rows or None,
                        "print_title_columns": ws.print_title_cols or None,
                        "merged_range_count": len(ws.merged_cells.ranges),
                        "table_count": len(tables),
                        "chart_count": len(charts),
                        "protection": {
                            "enabled": _sheet_protection_state(ws)["enabled"],
                            "password_protected": _sheet_protection_state(ws)["password_protected"],
                        },
                        "tables": tables,
                        "charts": charts,
                    }
                )

            return {
                "filename": path.name,
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime,
                "sheet_count": len(wb.sheetnames),
                "named_range_count": len(named_ranges),
                "table_count": total_tables,
                "chart_count": total_charts,
                "sheets": sheets,
                "named_ranges": named_ranges,
            }

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to profile workbook: {e}")
        raise WorkbookError(str(e))


def analyze_range_impact(filepath: str, sheet_name: str, range_ref: str) -> dict[str, Any]:
    """Inspect workbook structures that overlap a worksheet range before mutation."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            ws = require_worksheet(
                wb,
                sheet_name,
                error_cls=WorkbookError,
                operation="range impact analysis",
            )
            target_bounds, normalized_range = _parse_range_reference(
                range_ref,
                worksheet=ws,
                expected_sheet=sheet_name,
                error_cls=WorkbookError,
            )

            from .chart import (
                DEFAULT_CHART_HEIGHT,
                DEFAULT_CHART_WIDTH,
                _chart_occupied_range,
                _chart_type_name,
                _extract_chart_anchor,
                _extract_chart_dimensions,
                _extract_title_text,
            )
            from .tables import _build_table_metadata

            tables: list[dict[str, Any]] = []
            for table in ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                table_bounds = (min_row, min_col, max_row, max_col)
                intersection = _intersection_bounds(target_bounds, table_bounds)
                if intersection is None:
                    continue

                metadata = _build_table_metadata(sheet_name, ws, table)
                header_row_count = metadata["header_row_count"]
                totals_row_count = metadata["totals_row_count"]
                header_bounds = (
                    (min_row, min_col, min_row + header_row_count - 1, max_col)
                    if header_row_count > 0
                    else None
                )
                data_start_row = min_row + header_row_count
                data_end_row = max_row - totals_row_count
                data_bounds = (
                    (data_start_row, min_col, data_end_row, max_col)
                    if data_start_row <= data_end_row
                    else None
                )
                totals_bounds = (
                    (max_row - totals_row_count + 1, min_col, max_row, max_col)
                    if totals_row_count > 0
                    else None
                )
                tables.append(
                    {
                        "table_name": table.displayName,
                        "range": table.ref,
                        "intersection_range": _bounds_to_range(*intersection),
                        "covers_header": bool(
                            header_bounds and _bounds_intersect(target_bounds, header_bounds)
                        ),
                        "covers_data": bool(
                            data_bounds and _bounds_intersect(target_bounds, data_bounds)
                        ),
                        "covers_totals_row": bool(
                            totals_bounds and _bounds_intersect(target_bounds, totals_bounds)
                        ),
                    }
                )

            charts: list[dict[str, Any]] = []
            for chart_index, chart in enumerate(getattr(ws, "_charts", []), start=1):
                anchor = _extract_chart_anchor(chart)
                if not anchor:
                    continue
                width, height = _extract_chart_dimensions(chart)
                occupied_range = _chart_occupied_range(
                    ws,
                    anchor,
                    width=width or DEFAULT_CHART_WIDTH,
                    height=height or DEFAULT_CHART_HEIGHT,
                )
                chart_bounds, _ = _parse_range_reference(occupied_range, error_cls=WorkbookError)
                intersection = _intersection_bounds(target_bounds, chart_bounds)
                if intersection is None:
                    continue
                charts.append(
                    {
                        "chart_index": chart_index,
                        "chart_type": _chart_type_name(chart),
                        "title": _extract_title_text(getattr(chart, "title", None)),
                        "anchor": anchor,
                        "occupied_range": occupied_range,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            merged_ranges: list[dict[str, Any]] = []
            for merged_range in ws.merged_cells.ranges:
                merged_bounds, merged_ref = _parse_range_reference(
                    str(merged_range),
                    worksheet=ws,
                    error_cls=WorkbookError,
                )
                intersection = _intersection_bounds(target_bounds, merged_bounds)
                if intersection is None:
                    continue
                merged_ranges.append(
                    {
                        "range": merged_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            named_ranges: list[dict[str, Any]] = []
            for named_range in _serialize_named_ranges(wb):
                matching_destinations = []
                for destination in named_range["destinations"]:
                    if destination["sheet_name"] != sheet_name:
                        continue
                    for destination_bounds, destination_ref in _iter_range_references(
                        destination["range"],
                        worksheet=ws,
                        expected_sheet=sheet_name,
                    ):
                        intersection = _intersection_bounds(target_bounds, destination_bounds)
                        if intersection is None:
                            continue
                        matching_destinations.append(
                            {
                                "range": destination_ref,
                                "intersection_range": _bounds_to_range(*intersection),
                            }
                        )

                if matching_destinations:
                    named_ranges.append(
                        {
                            "name": named_range["name"],
                            "local_sheet": named_range["local_sheet"],
                            "hidden": named_range["hidden"],
                            "destinations": matching_destinations,
                        }
                    )

            autofilter = None
            if ws.auto_filter.ref:
                autofilter_bounds, autofilter_ref = _parse_range_reference(
                    ws.auto_filter.ref,
                    worksheet=ws,
                    error_cls=WorkbookError,
                )
                intersection = _intersection_bounds(target_bounds, autofilter_bounds)
                if intersection is not None:
                    autofilter = {
                        "range": autofilter_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }

            print_area_matches = []
            for print_bounds, print_ref in _iter_range_references(
                ws.print_area,
                worksheet=ws,
                expected_sheet=sheet_name,
            ):
                intersection = _intersection_bounds(target_bounds, print_bounds)
                if intersection is None:
                    continue
                print_area_matches.append(
                    {
                        "range": print_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            formula_cells = []
            for row in ws.iter_rows(
                min_row=target_bounds[0],
                max_row=target_bounds[2],
                min_col=target_bounds[1],
                max_col=target_bounds[3],
            ):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_cells.append(cell.coordinate)

            dependent_formulas = _extract_formula_dependencies(
                wb,
                target_sheet=sheet_name,
                target_bounds=target_bounds,
            )
            direct_formula_count = sum(
                1 for dependency in dependent_formulas if dependency.get("dependency_depth", 1) == 1
            )
            transitive_formula_count = len(dependent_formulas) - direct_formula_count
            data_validations = _extract_validation_overlaps(
                ws,
                sheet_name=sheet_name,
                target_bounds=target_bounds,
            )
            dependent_validations = _extract_validation_dependencies(
                wb,
                target_sheet=sheet_name,
                target_bounds=target_bounds,
            )
            conditional_formats = _extract_conditional_format_overlaps(
                ws,
                sheet_name=sheet_name,
                target_bounds=target_bounds,
            )
            dependent_conditional_formats = _extract_conditional_format_dependencies(
                wb,
                target_sheet=sheet_name,
                target_bounds=target_bounds,
            )

            impact_score = (
                len(tables) * 3
                + len(charts) * 3
                + len(merged_ranges) * 2
                + len(named_ranges) * 2
                + len(data_validations) * 2
                + len(conditional_formats) * 2
                + (1 if formula_cells else 0)
                + len(dependent_formulas) * 3
                + len(dependent_validations) * 2
                + len(dependent_conditional_formats) * 2
                + (1 if autofilter else 0)
                + (1 if print_area_matches else 0)
            )
            if impact_score >= 3:
                risk_level = "high"
            elif impact_score >= 1:
                risk_level = "medium"
            else:
                risk_level = "low"

            hints: list[str] = []
            if tables:
                hints.append("Selected range overlaps native Excel tables.")
                if any(table["covers_header"] for table in tables):
                    hints.append("Table header rows are inside the selected range.")
                if any(table["covers_totals_row"] for table in tables):
                    hints.append("Table totals rows are inside the selected range.")
            if charts:
                hints.append("Selected range overlaps embedded chart footprints.")
            if merged_ranges:
                hints.append("Selected range touches merged cells that may need to be unmerged first.")
            if named_ranges:
                hints.append("Named ranges point into the selected range.")
            if data_validations:
                hints.append("Selected range overlaps worksheet data validation rules.")
            if conditional_formats:
                hints.append("Selected range overlaps conditional formatting rules.")
            if autofilter:
                hints.append("Selected range overlaps the worksheet autofilter.")
            if print_area_matches:
                hints.append("Selected range overlaps the worksheet print area.")
            if formula_cells:
                hints.append("Selected range contains formula cells that may recalculate or break.")
            if dependent_formulas:
                hints.append("Formulas elsewhere in the workbook reference the selected range.")
            if dependent_validations:
                hints.append("Validation rules elsewhere in the workbook reference the selected range.")
            if dependent_conditional_formats:
                hints.append("Conditional formatting rules reference the selected range.")
            if not hints:
                hints.append("No overlapping workbook structures detected for this range.")

            return {
                "sheet_name": sheet_name,
                "range": normalized_range,
                "used_range": _get_used_range(ws),
                "summary": {
                    "risk_level": risk_level,
                    "table_count": len(tables),
                    "chart_count": len(charts),
                    "merged_range_count": len(merged_ranges),
                    "named_range_count": len(named_ranges),
                    "data_validation_count": len(data_validations),
                    "conditional_format_count": len(conditional_formats),
                    "formula_cell_count": len(formula_cells),
                    "dependent_formula_count": len(dependent_formulas),
                    "direct_formula_count": direct_formula_count,
                    "transitive_formula_count": transitive_formula_count,
                    "dependent_validation_count": len(dependent_validations),
                    "dependent_conditional_format_count": len(dependent_conditional_formats),
                    "autofilter_overlap": autofilter is not None,
                    "print_area_overlap": bool(print_area_matches),
                },
                "tables": tables,
                "charts": charts,
                "merged_ranges": merged_ranges,
                "named_ranges": named_ranges,
                "data_validations": {
                    "count": len(data_validations),
                    "sample": data_validations[:10],
                },
                "conditional_formats": {
                    "count": len(conditional_formats),
                    "sample": conditional_formats[:10],
                },
                "formula_cells": {
                    "count": len(formula_cells),
                    "sample": formula_cells[:10],
                },
                "dependent_formulas": {
                    "count": len(dependent_formulas),
                    "direct_count": direct_formula_count,
                    "transitive_count": transitive_formula_count,
                    "sample": dependent_formulas[:10],
                },
                "dependent_validations": {
                    "count": len(dependent_validations),
                    "sample": dependent_validations[:10],
                },
                "dependent_conditional_formats": {
                    "count": len(dependent_conditional_formats),
                    "sample": dependent_conditional_formats[:10],
                },
                "worksheet_features": {
                    "autofilter": autofilter,
                    "print_area": print_area_matches,
                },
                "hints": hints,
            }

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to analyze range impact: {e}")
        raise WorkbookError(str(e))
