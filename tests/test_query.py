import json

import pytest

from excel_mcp.exceptions import DataError
from excel_mcp.query import (
    aggregate_table as aggregate_table_impl,
    query_table as query_table_impl,
)
from excel_mcp.server import (
    aggregate_table as aggregate_table_tool,
    query_table as query_table_tool,
)


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    assert "operation" in payload
    assert "message" in payload
    return payload


def test_query_table_filters_selects_and_returns_records(tmp_workbook):
    result = query_table_impl(
        tmp_workbook,
        sheet_name="Sheet1",
        filters=[{"field": "age", "op": "gte", "value": 30}],
        select=["name", "AGE"],
        sort_by="AGE",
        sort_desc=True,
        limit=2,
        row_mode="objects",
        infer_schema=True,
    )

    assert result["headers"] == ["Name", "Age"]
    assert result["records"] == [
        {"name": "Carol", "age": 35},
        {"name": "Eve", "age": 32},
    ]
    assert result["schema"] == [
        {"field": "name", "header": "Name", "type": "string", "nullable": False},
        {"field": "age", "header": "Age", "type": "integer", "nullable": False},
    ]
    assert result["row_mode"] == "objects"
    assert "rows" not in result
    assert result["matched_rows"] == 3
    assert result["returned_rows"] == 2
    assert result["source_row_count"] == 5
    assert result["truncated"] is True
    assert result["sort_by"] == "Age"
    assert result["sort_desc"] is True
    assert result["filters"] == [{"field": "Age", "op": "gte", "value": 30}]


def test_query_table_can_read_native_excel_table(complex_workbook):
    result = query_table_impl(
        complex_workbook,
        table_name="SalesData",
        filters=[{"field": "region", "op": "eq", "value": "north"}],
        select=["Product", "sales"],
        sort_by="sales",
        sort_desc=True,
    )

    assert result["target_kind"] == "excel_table"
    assert result["sheet_name"] == "Data"
    assert result["table_name"] == "SalesData"
    assert result["headers"] == ["Product", "Sales"]
    assert result["rows"] == [
        ["Tool", 16],
        ["Widget", 12],
    ]
    assert result["matched_rows"] == 2
    assert result["returned_rows"] == 2
    assert result["source_row_count"] == 5
    assert result["truncated"] is False


def test_query_table_rejects_explicit_chart_sheet(complex_workbook):
    with pytest.raises(DataError, match="Sheet 'Charts' is a chartsheet"):
        query_table_impl(complex_workbook, sheet_name="Charts")


def test_query_table_rejects_non_positive_header_row(tmp_workbook):
    with pytest.raises(DataError, match="header_row must be a positive integer"):
        query_table_impl(tmp_workbook, sheet_name="Sheet1", header_row=0)


def test_aggregate_table_groups_metrics_and_returns_records(complex_workbook):
    result = aggregate_table_impl(
        complex_workbook,
        table_name="SalesData",
        group_by=["region"],
        metrics=[
            {"op": "sum", "field": "sales", "as": "total_sales"},
            {"op": "count_non_null", "field": "Target", "as": "target_rows"},
        ],
        sort_by="total_sales",
        sort_desc=True,
        row_mode="objects",
        infer_schema=True,
    )

    assert result["headers"] == ["Region", "total_sales", "target_rows"]
    assert result["records"] == [
        {"region": "East", "total_sales": 30, "target_rows": 1},
        {"region": "North", "total_sales": 28, "target_rows": 2},
        {"region": "South", "total_sales": 24, "target_rows": 1},
        {"region": "West", "total_sales": 18, "target_rows": 1},
    ]
    assert result["schema"] == [
        {"field": "region", "header": "Region", "type": "string", "nullable": False},
        {"field": "total_sales", "header": "total_sales", "type": "integer", "nullable": False},
        {"field": "target_rows", "header": "target_rows", "type": "integer", "nullable": False},
    ]
    assert result["row_mode"] == "objects"
    assert "rows" not in result
    assert result["group_by"] == ["Region"]
    assert result["metrics"] == [
        {"op": "sum", "field": "Sales", "as": "total_sales"},
        {"op": "count_non_null", "field": "Target", "as": "target_rows"},
    ]
    assert result["group_count"] == 4
    assert result["returned_groups"] == 4
    assert result["matched_rows"] == 5
    assert result["source_row_count"] == 5
    assert result["truncated"] is False


def test_aggregate_table_rejects_metric_alias_collision_with_group_name(complex_workbook):
    with pytest.raises(DataError, match="Duplicate metric alias 'Region' is not allowed"):
        aggregate_table_impl(
            complex_workbook,
            table_name="SalesData",
            group_by=["Region"],
            metrics=[{"op": "sum", "field": "Sales", "as": "Region"}],
        )


def test_aggregate_table_rejects_boolean_limit(complex_workbook):
    with pytest.raises(DataError, match="limit must be a positive integer"):
        aggregate_table_impl(complex_workbook, table_name="SalesData", limit=True)


def test_query_table_tool_returns_json_envelope(tmp_workbook):
    payload = _load_tool_payload(
        query_table_tool(
            tmp_workbook,
            sheet_name="Sheet1",
            filters=[{"field": "Age", "op": "gt", "value": 28}],
            limit=1,
        )
    )

    assert payload["operation"] == "query_table"
    assert payload["data"]["returned_rows"] == 1
    assert payload["data"]["rows"] == [["Alice", 30, "Helsinki"]]


def test_aggregate_table_tool_returns_json_envelope(complex_workbook):
    payload = _load_tool_payload(
        aggregate_table_tool(
            complex_workbook,
            table_name="SalesData",
            group_by=["Region"],
            metrics=[{"op": "sum", "field": "Sales", "as": "total_sales"}],
            sort_by="total_sales",
            sort_desc=True,
            limit=2,
        )
    )

    assert payload["operation"] == "aggregate_table"
    assert payload["data"]["returned_groups"] == 2
    assert payload["data"]["rows"] == [
        ["East", 30],
        ["North", 28],
    ]
