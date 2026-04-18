import logging
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional

from .data import (
    _build_schema,
    _stringify_value_for_uniqueness,
    _validate_row_mode,
    augment_tabular_payload,
    quick_read,
    read_as_table,
)
from .exceptions import DataError
from .tables import read_excel_table

logger = logging.getLogger(__name__)

FILTER_OPERATORS = {
    "eq",
    "neq",
    "gt",
    "gte",
    "lt",
    "lte",
    "contains",
    "starts_with",
    "ends_with",
    "in",
    "not_in",
    "is_blank",
    "not_blank",
}

AGGREGATE_OPERATORS = {
    "count",
    "count_non_null",
    "count_distinct",
    "sum",
    "avg",
    "min",
    "max",
}
SCHEMA_MODES = {"strict", "intersect", "union"}


def _is_numeric(value: Any) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(value, bool)


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _validate_positive_integer(value: Optional[int], *, argument_name: str) -> None:
    if value is None:
        return
    if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
        raise DataError(f"{argument_name} must be a positive integer")


def _validate_schema_mode(schema_mode: str) -> None:
    if schema_mode not in SCHEMA_MODES:
        raise DataError("schema_mode must be 'strict', 'intersect', or 'union'")


def _validate_filepaths(filepaths: List[str]) -> None:
    if not isinstance(filepaths, list) or not filepaths:
        raise DataError("filepaths must be a non-empty list of workbook paths")
    for index, filepath in enumerate(filepaths, start=1):
        if not isinstance(filepath, str) or not filepath.strip():
            raise DataError(f"filepaths[{index}] must be a non-empty string")


def _load_source_dataset(
    filepath: str,
    *,
    sheet_name: Optional[str] = None,
    table_name: Optional[str] = None,
    header_row: int = 1,
) -> Dict[str, Any]:
    if table_name is not None:
        result = read_excel_table(
            filepath,
            table_name,
            sheet_name=sheet_name,
            row_mode="arrays",
            infer_schema=True,
        )
        return {
            "target_kind": "excel_table",
            "sheet_name": result["sheet_name"],
            "table_name": result["table_name"],
            "headers": result["headers"],
            "rows": result["rows"],
            "schema": result.get("schema", []),
            "total_rows": result["total_rows"],
            "range": result["range"],
        }

    if sheet_name is not None:
        result = read_as_table(
            filepath,
            sheet_name=sheet_name,
            header_row=header_row,
            row_mode="arrays",
            infer_schema=True,
        )
    else:
        result = quick_read(
            filepath,
            header_row=header_row,
            row_mode="arrays",
            infer_schema=True,
        )

    return {
        "target_kind": "worksheet",
        "sheet_name": result["sheet_name"],
        "table_name": None,
        "headers": result["headers"],
        "rows": result["rows"],
        "schema": result.get("schema", []),
        "total_rows": result["total_rows"],
        "range": None,
        "auto_selected_sheet": result.get("auto_selected_sheet", False),
    }


def _field_keys(schema: List[Dict[str, Any]]) -> set[str]:
    return {str(column["field"]).casefold() for column in schema}


def _collect_aggregate_input_refs(
    *,
    filters: Optional[List[Dict[str, Any]]],
    group_by: Optional[List[str]],
    metrics: Optional[List[Dict[str, Any]]],
) -> List[str]:
    refs: List[str] = []
    seen_refs: set[str] = set()

    def add_ref(value: Any) -> None:
        if not isinstance(value, str) or not value.strip():
            return
        normalized = value.strip()
        ref_key = normalized.casefold()
        if ref_key in seen_refs:
            return
        seen_refs.add(ref_key)
        refs.append(normalized)

    if isinstance(filters, list):
        for filter_spec in filters:
            if isinstance(filter_spec, dict):
                add_ref(filter_spec.get("field"))

    if isinstance(group_by, list):
        for field_ref in group_by:
            add_ref(field_ref)

    if isinstance(metrics, list):
        for metric in metrics:
            if not isinstance(metric, dict):
                continue
            operator = str(metric.get("op", "")).strip().lower()
            if operator == "count":
                continue
            add_ref(metric.get("field"))

    return refs


def _column_lookup(headers: List[Any], schema: List[Dict[str, Any]]) -> Dict[str, int]:
    lookup: Dict[str, int] = {}
    casefold_headers: Dict[str, List[int]] = defaultdict(list)
    casefold_fields: Dict[str, List[int]] = defaultdict(list)

    for index, header in enumerate(headers):
        if header is not None:
            header_text = str(header)
            lookup.setdefault(header_text, index)
            casefold_headers[header_text.casefold()].append(index)

    for index, column in enumerate(schema):
        field_name = column["field"]
        lookup.setdefault(field_name, index)
        casefold_fields[field_name.casefold()].append(index)

    for casefolded, indexes in casefold_headers.items():
        if len(indexes) == 1:
            header_index = indexes[0]
            header_text = headers[header_index]
            if header_text is not None:
                lookup.setdefault(casefolded, header_index)

    for casefolded, indexes in casefold_fields.items():
        if len(indexes) == 1:
            lookup.setdefault(casefolded, indexes[0])

    return lookup


def _resolve_column(
    column_ref: str,
    headers: List[Any],
    schema: List[Dict[str, Any]],
    *,
    argument_name: str,
) -> tuple[int, str]:
    if not isinstance(column_ref, str) or not column_ref.strip():
        raise DataError(f"{argument_name} must be a non-empty string")

    lookup = _column_lookup(headers, schema)
    normalized_ref = column_ref.strip()
    candidate_keys = [normalized_ref, normalized_ref.casefold()]
    for key in candidate_keys:
        if key in lookup:
            index = lookup[key]
            header = headers[index]
            return index, "" if header is None else str(header)

    raise DataError(f"{argument_name} '{column_ref}' was not found in the selected dataset")


def _resolve_column_or_none(
    column_ref: str,
    headers: List[Any],
    schema: List[Dict[str, Any]],
) -> Optional[tuple[int, str]]:
    try:
        return _resolve_column(
            column_ref,
            headers,
            schema,
            argument_name="column",
        )
    except DataError:
        return None


def _normalize_string(value: Any, *, case_sensitive: bool) -> Optional[str]:
    if value is None:
        return None
    text = str(value)
    return text if case_sensitive else text.casefold()


def _normalize_membership_value(value: Any, *, case_sensitive: bool) -> Any:
    if isinstance(value, str):
        return _normalize_string(value, case_sensitive=case_sensitive)
    return value


def _matches_filter(
    cell_value: Any,
    *,
    operator: str,
    value: Any = None,
    values: Optional[List[Any]] = None,
    case_sensitive: bool = False,
    field_name: str,
) -> bool:
    if operator == "is_blank":
        return _is_blank(cell_value)
    if operator == "not_blank":
        return not _is_blank(cell_value)

    if operator in {"contains", "starts_with", "ends_with"}:
        if cell_value is None:
            return False
        haystack = _normalize_string(cell_value, case_sensitive=case_sensitive)
        needle = _normalize_string(value, case_sensitive=case_sensitive)
        if haystack is None or needle is None:
            return False
        if operator == "contains":
            return needle in haystack
        if operator == "starts_with":
            return haystack.startswith(needle)
        return haystack.endswith(needle)

    if operator in {"in", "not_in"}:
        if values is None:
            raise DataError(f"Filter on '{field_name}' with operator '{operator}' requires values")
        normalized_values = {
            _normalize_membership_value(item, case_sensitive=case_sensitive) for item in values
        }
        normalized_cell_value = _normalize_membership_value(
            cell_value,
            case_sensitive=case_sensitive,
        )
        match = normalized_cell_value in normalized_values
        return match if operator == "in" else not match

    if operator in {"eq", "neq"} and isinstance(cell_value, str) and isinstance(value, str):
        left = _normalize_string(cell_value, case_sensitive=case_sensitive)
        right = _normalize_string(value, case_sensitive=case_sensitive)
        match = left == right
        return match if operator == "eq" else not match

    if operator == "eq":
        return cell_value == value
    if operator == "neq":
        return cell_value != value

    if cell_value is None or value is None:
        return False

    try:
        if operator == "gt":
            return cell_value > value
        if operator == "gte":
            return cell_value >= value
        if operator == "lt":
            return cell_value < value
        if operator == "lte":
            return cell_value <= value
    except TypeError as exc:
        raise DataError(
            f"Cannot compare values for field '{field_name}' with operator '{operator}'"
        ) from exc

    raise DataError(f"Unsupported filter operator '{operator}'")


def _normalize_filters(
    filters: Optional[List[Dict[str, Any]]],
    headers: List[Any],
    schema: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if filters is None:
        return []
    if not isinstance(filters, list):
        raise DataError("filters must be a list of filter objects")

    normalized_filters: List[Dict[str, Any]] = []
    for index, filter_spec in enumerate(filters, start=1):
        if not isinstance(filter_spec, dict):
            raise DataError(f"filters[{index}] must be an object")

        field_ref = filter_spec.get("field")
        operator = str(filter_spec.get("op", "eq")).strip().lower()
        if operator not in FILTER_OPERATORS:
            raise DataError(
                f"filters[{index}] uses unsupported operator '{operator}'"
            )
        column_index, header_name = _resolve_column(
            field_ref,
            headers,
            schema,
            argument_name=f"filters[{index}].field",
        )

        normalized_filter = {
            "field": field_ref,
            "resolved_header": header_name,
            "column_index": column_index,
            "op": operator,
            "case_sensitive": bool(filter_spec.get("case_sensitive", False)),
        }

        if operator in {"in", "not_in"}:
            values = filter_spec.get("values")
            if not isinstance(values, list):
                raise DataError(
                    f"filters[{index}] with operator '{operator}' requires a list in values"
                )
            normalized_filter["values"] = values
        elif operator not in {"is_blank", "not_blank"}:
            if "value" not in filter_spec:
                raise DataError(
                    f"filters[{index}] with operator '{operator}' requires a value"
                )
            normalized_filter["value"] = filter_spec["value"]

        normalized_filters.append(normalized_filter)

    return normalized_filters


def _apply_filters(
    rows: List[List[Any]],
    normalized_filters: List[Dict[str, Any]],
) -> List[List[Any]]:
    if not normalized_filters:
        return list(rows)

    matched_rows: List[List[Any]] = []
    for row in rows:
        keep_row = True
        for filter_spec in normalized_filters:
            column_index = filter_spec["column_index"]
            cell_value = row[column_index] if column_index < len(row) else None
            if not _matches_filter(
                cell_value,
                operator=filter_spec["op"],
                value=filter_spec.get("value"),
                values=filter_spec.get("values"),
                case_sensitive=filter_spec["case_sensitive"],
                field_name=filter_spec["resolved_header"],
            ):
                keep_row = False
                break
        if keep_row:
            matched_rows.append(row)

    return matched_rows


def _sort_rows(
    rows: List[List[Any]],
    column_index: int,
    *,
    sort_desc: bool,
    field_name: str,
) -> List[List[Any]]:
    non_null_rows = [row for row in rows if column_index < len(row) and row[column_index] is not None]
    null_rows = [row for row in rows if column_index >= len(row) or row[column_index] is None]

    try:
        sorted_non_null = sorted(
            non_null_rows,
            key=lambda row: row[column_index],
            reverse=sort_desc,
        )
    except TypeError as exc:
        raise DataError(f"Cannot sort by field '{field_name}' because values are not comparable") from exc

    return sorted_non_null + null_rows


def _select_columns(
    headers: List[Any],
    rows: List[List[Any]],
    schema: List[Dict[str, Any]],
    select: Optional[List[str]],
) -> tuple[List[Any], List[List[Any]], List[Dict[str, Any]]]:
    if select is None:
        return headers, rows, schema
    if not isinstance(select, list) or not select:
        raise DataError("select must be a non-empty list of column references")

    resolved_columns: List[tuple[int, str]] = []
    seen_indexes: set[int] = set()
    for index, column_ref in enumerate(select, start=1):
        column_index, header_name = _resolve_column(
            column_ref,
            headers,
            schema,
            argument_name=f"select[{index}]",
        )
        if column_index in seen_indexes:
            continue
        seen_indexes.add(column_index)
        resolved_columns.append((column_index, header_name))

    selected_headers = [headers[column_index] for column_index, _ in resolved_columns]
    selected_rows = [
        [row[column_index] if column_index < len(row) else None for column_index, _ in resolved_columns]
        for row in rows
    ]
    selected_schema = [schema[column_index] for column_index, _ in resolved_columns]
    return selected_headers, selected_rows, selected_schema


def _normalize_metrics(
    metrics: Optional[List[Dict[str, Any]]],
    headers: List[Any],
    schema: List[Dict[str, Any]],
    *,
    reserved_aliases: Optional[set[str]] = None,
) -> List[Dict[str, Any]]:
    if not metrics:
        reserved = reserved_aliases or set()
        if "row_count".casefold() in reserved:
            raise DataError("Metric alias 'row_count' conflicts with an existing output column")
        return [{"op": "count", "as": "row_count", "field": None, "column_index": None}]
    if not isinstance(metrics, list):
        raise DataError("metrics must be a list of metric objects")

    normalized_metrics: List[Dict[str, Any]] = []
    seen_aliases: set[str] = set(reserved_aliases or set())
    for index, metric in enumerate(metrics, start=1):
        if not isinstance(metric, dict):
            raise DataError(f"metrics[{index}] must be an object")

        operator = str(metric.get("op", "")).strip().lower()
        if operator not in AGGREGATE_OPERATORS:
            raise DataError(f"metrics[{index}] uses unsupported operator '{operator}'")

        field_ref = metric.get("field")
        column_index = None
        header_name = None
        if operator != "count":
            column_index, header_name = _resolve_column(
                field_ref,
                headers,
                schema,
                argument_name=f"metrics[{index}].field",
            )

        alias = metric.get("as")
        if alias is None:
            if operator == "count":
                alias = "row_count"
            else:
                alias = f"{operator}_{schema[column_index]['field']}"
        alias = str(alias).strip()
        if not alias:
            raise DataError(f"metrics[{index}].as must not be blank")
        alias_key = alias.casefold()
        if alias_key in seen_aliases:
            raise DataError(f"Duplicate metric alias '{alias}' is not allowed")
        seen_aliases.add(alias_key)

        normalized_metrics.append(
            {
                "op": operator,
                "field": field_ref,
                "resolved_header": header_name,
                "column_index": column_index,
                "as": alias,
            }
        )

    return normalized_metrics


def _group_by_indexes(
    group_by: Optional[List[str]],
    headers: List[Any],
    schema: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if group_by is None:
        return []
    if not isinstance(group_by, list):
        raise DataError("group_by must be a list of column references")

    groups: List[Dict[str, Any]] = []
    seen_indexes: set[int] = set()
    for index, field_ref in enumerate(group_by, start=1):
        column_index, header_name = _resolve_column(
            field_ref,
            headers,
            schema,
            argument_name=f"group_by[{index}]",
        )
        if column_index in seen_indexes:
            continue
        seen_indexes.add(column_index)
        groups.append(
            {
                "field": field_ref,
                "resolved_header": header_name,
                "column_index": column_index,
            }
        )
    return groups


def _compute_metric(metric: Dict[str, Any], rows: List[List[Any]]) -> Any:
    operator = metric["op"]
    column_index = metric["column_index"]

    if operator == "count":
        return len(rows)

    values = [
        row[column_index] if column_index is not None and column_index < len(row) else None
        for row in rows
    ]
    non_blank_values = [value for value in values if value is not None]

    if operator == "count_non_null":
        return len(non_blank_values)
    if operator == "count_distinct":
        return len({_stringify_value_for_uniqueness(value) for value in non_blank_values})

    if operator in {"sum", "avg"}:
        numeric_values = [value for value in non_blank_values if _is_numeric(value)]
        if len(numeric_values) != len(non_blank_values):
            raise DataError(
                f"Metric '{metric['as']}' requires numeric values in field '{metric['resolved_header']}'"
            )
        if operator == "sum":
            return sum(numeric_values)
        return sum(numeric_values) / len(numeric_values) if numeric_values else None

    comparable_values = list(non_blank_values)
    if not comparable_values:
        return None

    try:
        if operator == "min":
            return min(comparable_values)
        if operator == "max":
            return max(comparable_values)
    except TypeError as exc:
        raise DataError(
            f"Metric '{metric['as']}' cannot compare mixed values in field '{metric['resolved_header']}'"
        ) from exc

    raise DataError(f"Unsupported aggregate operator '{operator}'")


def _finalize_tabular_result(
    *,
    payload: Dict[str, Any],
    headers: List[Any],
    rows: List[List[Any]],
    row_mode: str,
    infer_schema: bool,
) -> Dict[str, Any]:
    return augment_tabular_payload(
        payload,
        headers=headers,
        rows=rows,
        row_mode=row_mode,
        infer_schema=infer_schema,
    )


def _aggregate_dataset(
    *,
    headers: List[Any],
    rows: List[List[Any]],
    target_kind: str,
    sheet_name: Optional[str],
    table_name: Optional[str],
    auto_selected_sheet: bool,
    filters: Optional[List[Dict[str, Any]]],
    group_by: Optional[List[str]],
    metrics: Optional[List[Dict[str, Any]]],
    sort_by: Optional[str],
    sort_desc: bool,
    limit: Optional[int],
    row_mode: str,
    infer_schema: bool,
    extra_payload: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    schema = _build_schema(headers, rows)
    normalized_filters = _normalize_filters(filters, headers, schema)
    matched_rows = _apply_filters(rows, normalized_filters)
    normalized_groups = _group_by_indexes(group_by, headers, schema)
    normalized_metrics = _normalize_metrics(
        metrics,
        headers,
        schema,
        reserved_aliases={
            group["resolved_header"].casefold()
            for group in normalized_groups
            if group["resolved_header"]
        },
    )

    grouped_rows: Dict[tuple[Any, ...], List[List[Any]]] = defaultdict(list)
    if normalized_groups:
        for row in matched_rows:
            key = tuple(
                row[group["column_index"]] if group["column_index"] < len(row) else None
                for group in normalized_groups
            )
            grouped_rows[key].append(row)
    else:
        grouped_rows[tuple()] = list(matched_rows)

    result_headers = [group["resolved_header"] for group in normalized_groups] + [
        metric["as"] for metric in normalized_metrics
    ]
    result_rows: List[List[Any]] = []
    for key, grouped in grouped_rows.items():
        output_row = list(key)
        for metric in normalized_metrics:
            output_row.append(_compute_metric(metric, grouped))
        result_rows.append(output_row)

    resolved_sort_by = None
    if sort_by is not None:
        aggregate_schema = _build_schema(result_headers, result_rows)
        sort_index, resolved_sort_by = _resolve_column(
            sort_by,
            result_headers,
            aggregate_schema,
            argument_name="sort_by",
        )
        result_rows = _sort_rows(
            result_rows,
            sort_index,
            sort_desc=sort_desc,
            field_name=resolved_sort_by,
        )

    total_groups = len(result_rows)
    if limit is not None:
        result_rows = result_rows[:limit]

    result: Dict[str, Any] = {
        "target_kind": target_kind,
        "sheet_name": sheet_name,
        "table_name": table_name,
        "auto_selected_sheet": auto_selected_sheet,
        "headers": result_headers,
        "rows": result_rows,
        "group_by": [group["resolved_header"] for group in normalized_groups],
        "metrics": [
            {
                "op": metric["op"],
                "field": metric["resolved_header"],
                "as": metric["as"],
            }
            for metric in normalized_metrics
        ],
        "group_count": total_groups,
        "returned_groups": len(result_rows),
        "source_row_count": len(rows),
        "matched_rows": len(matched_rows),
        "truncated": limit is not None and total_groups > len(result_rows),
        "filters": [
            {
                "field": item["resolved_header"],
                "op": item["op"],
                **({"value": item["value"]} if "value" in item else {}),
                **({"values": item["values"]} if "values" in item else {}),
                **({"case_sensitive": True} if item["case_sensitive"] else {}),
            }
            for item in normalized_filters
        ],
        "sort_by": resolved_sort_by,
        "sort_desc": sort_desc,
    }
    if extra_payload:
        result.update(extra_payload)

    return _finalize_tabular_result(
        payload=result,
        headers=result_headers,
        rows=result_rows,
        row_mode=row_mode,
        infer_schema=infer_schema,
    )


def query_table(
    filepath: str,
    *,
    sheet_name: Optional[str] = None,
    table_name: Optional[str] = None,
    header_row: int = 1,
    select: Optional[List[str]] = None,
    filters: Optional[List[Dict[str, Any]]] = None,
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    limit: Optional[int] = None,
    row_mode: str = "arrays",
    infer_schema: bool = False,
) -> Dict[str, Any]:
    """Query worksheet-shaped data or a native Excel table with declarative filters."""
    try:
        _validate_row_mode(row_mode)
        _validate_positive_integer(header_row, argument_name="header_row")
        _validate_positive_integer(limit, argument_name="limit")

        source = _load_source_dataset(
            filepath,
            sheet_name=sheet_name,
            table_name=table_name,
            header_row=header_row,
        )
        headers = source["headers"]
        rows = source["rows"]
        schema = source["schema"]

        normalized_filters = _normalize_filters(filters, headers, schema)
        matched_rows = _apply_filters(rows, normalized_filters)

        resolved_sort_by = None
        if sort_by is not None:
            sort_index, resolved_sort_by = _resolve_column(
                sort_by,
                headers,
                schema,
                argument_name="sort_by",
            )
            matched_rows = _sort_rows(
                matched_rows,
                sort_index,
                sort_desc=sort_desc,
                field_name=resolved_sort_by,
            )

        matched_row_count = len(matched_rows)
        if limit is not None:
            matched_rows = matched_rows[:limit]

        selected_headers, selected_rows, selected_schema = _select_columns(
            headers,
            matched_rows,
            schema,
            select,
        )

        result = {
            "target_kind": source["target_kind"],
            "sheet_name": source["sheet_name"],
            "table_name": source["table_name"],
            "auto_selected_sheet": source.get("auto_selected_sheet", False),
            "headers": selected_headers,
            "rows": selected_rows,
            "matched_rows": matched_row_count,
            "returned_rows": len(selected_rows),
            "source_row_count": source["total_rows"],
            "truncated": limit is not None and matched_row_count > len(selected_rows),
            "filters": [
                {
                    "field": item["resolved_header"],
                    "op": item["op"],
                    **({"value": item["value"]} if "value" in item else {}),
                    **({"values": item["values"]} if "values" in item else {}),
                    **({"case_sensitive": True} if item["case_sensitive"] else {}),
                }
                for item in normalized_filters
            ],
            "select": select,
            "sort_by": resolved_sort_by,
            "sort_desc": sort_desc,
        }

        return _finalize_tabular_result(
            payload=result,
            headers=selected_headers,
            rows=selected_rows,
            row_mode=row_mode,
            infer_schema=infer_schema,
        )
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to query table data: {e}")
        raise DataError(str(e))


def aggregate_table(
    filepath: str,
    *,
    sheet_name: Optional[str] = None,
    table_name: Optional[str] = None,
    header_row: int = 1,
    filters: Optional[List[Dict[str, Any]]] = None,
    group_by: Optional[List[str]] = None,
    metrics: Optional[List[Dict[str, Any]]] = None,
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    limit: Optional[int] = None,
    row_mode: str = "arrays",
    infer_schema: bool = False,
) -> Dict[str, Any]:
    """Aggregate worksheet-shaped data or a native Excel table with declarative metrics."""
    try:
        _validate_row_mode(row_mode)
        _validate_positive_integer(header_row, argument_name="header_row")
        _validate_positive_integer(limit, argument_name="limit")

        source = _load_source_dataset(
            filepath,
            sheet_name=sheet_name,
            table_name=table_name,
            header_row=header_row,
        )
        return _aggregate_dataset(
            headers=source["headers"],
            rows=source["rows"],
            target_kind=source["target_kind"],
            sheet_name=source["sheet_name"],
            table_name=source["table_name"],
            auto_selected_sheet=source.get("auto_selected_sheet", False),
            filters=filters,
            group_by=group_by,
            metrics=metrics,
            sort_by=sort_by,
            sort_desc=sort_desc,
            limit=limit,
            row_mode=row_mode,
            infer_schema=infer_schema,
            extra_payload={
                "source_row_count": source["total_rows"],
            },
        )
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to aggregate table data: {e}")
        raise DataError(str(e))


def bulk_aggregate_workbooks(
    filepaths: List[str],
    *,
    sheet_name: Optional[str] = None,
    table_name: Optional[str] = None,
    header_row: int = 1,
    filters: Optional[List[Dict[str, Any]]] = None,
    group_by: Optional[List[str]] = None,
    metrics: Optional[List[Dict[str, Any]]] = None,
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    limit: Optional[int] = None,
    schema_mode: str = "strict",
    source_sample_limit: int = 10,
    row_mode: str = "arrays",
    infer_schema: bool = False,
) -> Dict[str, Any]:
    """Aggregate comparable worksheet or table data across multiple workbooks."""
    try:
        _validate_filepaths(filepaths)
        _validate_row_mode(row_mode)
        _validate_positive_integer(header_row, argument_name="header_row")
        _validate_positive_integer(limit, argument_name="limit")
        _validate_positive_integer(source_sample_limit, argument_name="source_sample_limit")
        _validate_schema_mode(schema_mode)

        sources: List[Dict[str, Any]] = []
        for filepath in filepaths:
            source = _load_source_dataset(
                filepath,
                sheet_name=sheet_name,
                table_name=table_name,
                header_row=header_row,
            )
            source["filepath"] = filepath
            source["file_name"] = Path(filepath).name
            sources.append(source)

        schema_key_sets = [_field_keys(source["schema"]) for source in sources]
        shared_field_keys = set.intersection(*schema_key_sets) if schema_key_sets else set()
        union_field_keys = set().union(*schema_key_sets)
        strict_compatible = all(
            key_set == schema_key_sets[0]
            for key_set in schema_key_sets[1:]
        ) if schema_key_sets else True
        if schema_mode == "strict" and not strict_compatible:
            baseline = sources[0]["file_name"]
            for source, key_set in zip(sources[1:], schema_key_sets[1:]):
                if key_set != schema_key_sets[0]:
                    raise DataError(
                        "schema_mode 'strict' requires identical columns across workbooks; "
                        f"'{source['file_name']}' differs from '{baseline}'"
                    )

        input_refs = _collect_aggregate_input_refs(
            filters=filters,
            group_by=group_by,
            metrics=metrics,
        )

        combined_headers: List[Any] = []
        for ref in input_refs:
            resolved_header = None
            for source in sources:
                resolved = _resolve_column_or_none(ref, source["headers"], source["schema"])
                if resolved is not None:
                    _, resolved_header = resolved
                    break
            if resolved_header is None:
                raise DataError(f"Column '{ref}' was not found in any selected workbook")
            combined_headers.append(resolved_header)

        combined_rows: List[List[Any]] = []
        source_rows_by_file: Dict[str, List[List[Any]]] = {}
        for source in sources:
            resolved_indexes: List[Optional[int]] = []
            for ref in input_refs:
                resolved = _resolve_column_or_none(ref, source["headers"], source["schema"])
                if resolved is None:
                    if schema_mode == "union":
                        resolved_indexes.append(None)
                        continue
                    raise DataError(
                        f"schema_mode '{schema_mode}' requires column '{ref}' in workbook "
                        f"'{source['file_name']}'"
                    )
                resolved_indexes.append(resolved[0])

            unified_rows: List[List[Any]] = []
            for row in source["rows"]:
                unified_row = [
                    row[column_index] if column_index is not None and column_index < len(row) else None
                    for column_index in resolved_indexes
                ]
                unified_rows.append(unified_row)
            source_rows_by_file[source["filepath"]] = unified_rows
            combined_rows.extend(unified_rows)

        combined_schema = _build_schema(combined_headers, combined_rows)
        normalized_filters = _normalize_filters(filters, combined_headers, combined_schema)
        source_workbook_sample = []
        for source in sources[:source_sample_limit]:
            unified_rows = source_rows_by_file[source["filepath"]]
            matched_rows = _apply_filters(unified_rows, normalized_filters)
            source_workbook_sample.append(
                {
                    "filepath": source["filepath"],
                    "file_name": source["file_name"],
                    "target_kind": source["target_kind"],
                    "sheet_name": source["sheet_name"],
                    "table_name": source["table_name"],
                    "source_row_count": source["total_rows"],
                    "matched_rows": len(matched_rows),
                    "auto_selected_sheet": source.get("auto_selected_sheet", False),
                }
            )

        return _aggregate_dataset(
            headers=combined_headers,
            rows=combined_rows,
            target_kind="multi_workbook",
            sheet_name=sheet_name,
            table_name=table_name,
            auto_selected_sheet=False,
            filters=filters,
            group_by=group_by,
            metrics=metrics,
            sort_by=sort_by,
            sort_desc=sort_desc,
            limit=limit,
            row_mode=row_mode,
            infer_schema=infer_schema,
            extra_payload={
                "workbook_count": len(sources),
                "processed_workbooks": len(sources),
                "schema_mode": schema_mode,
                "schema_summary": {
                    "strict_compatible": strict_compatible,
                    "shared_field_count": len(shared_field_keys),
                    "union_field_count": len(union_field_keys),
                },
                "source_workbooks": {
                    "count": len(sources),
                    "sample": source_workbook_sample,
                    "truncated": len(sources) > source_sample_limit,
                },
            },
        )
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to aggregate multiple workbooks: {e}")
        raise DataError(str(e))
