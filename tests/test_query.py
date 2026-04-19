import json

import pytest
from openpyxl import Workbook

from excel_mcp.exceptions import DataError
from excel_mcp.query import (
    aggregate_table as aggregate_table_impl,
    bulk_aggregate_workbooks as bulk_aggregate_workbooks_impl,
    bulk_filter_workbooks as bulk_filter_workbooks_impl,
    query_table as query_table_impl,
    union_tables as union_tables_impl,
)
from excel_mcp.server import (
    aggregate_table as aggregate_table_tool,
    bulk_aggregate_workbooks as bulk_aggregate_workbooks_tool,
    bulk_filter_workbooks as bulk_filter_workbooks_tool,
    query_table as query_table_tool,
    union_tables as union_tables_tool,
)


def _load_tool_payload(raw: str) -> dict:
    payload = json.loads(raw)
    assert payload["ok"] is True
    assert "operation" in payload
    assert "message" in payload
    return payload


def _create_query_workbook(tmp_path, filename: str, headers: list[str], rows: list[tuple]) -> str:
    filepath = str(tmp_path / filename)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet.append(headers)
    for row in rows:
        worksheet.append(list(row))
    workbook.save(filepath)
    workbook.close()
    return filepath


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


def test_bulk_aggregate_workbooks_combines_matching_workbooks(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Region", "Sales", "Channel"],
        [
            ("North", 10, "Paid"),
            ("South", 5, "Organic"),
        ],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Region", "Sales", "Channel"],
        [
            ("North", 7, "Paid"),
            ("East", 3, "Paid"),
        ],
    )

    result = bulk_aggregate_workbooks_impl(
        [january, february],
        sheet_name="Sheet1",
        group_by=["region"],
        metrics=[{"op": "sum", "field": "sales", "as": "total_sales"}],
        sort_by="total_sales",
        sort_desc=True,
        row_mode="objects",
        infer_schema=True,
    )

    assert result["target_kind"] == "multi_workbook"
    assert result["schema_mode"] == "strict"
    assert result["workbook_count"] == 2
    assert result["source_row_count"] == 4
    assert result["matched_rows"] == 4
    assert result["records"] == [
        {"region": "North", "total_sales": 17},
        {"region": "South", "total_sales": 5},
        {"region": "East", "total_sales": 3},
    ]
    assert result["source_workbooks"]["sample"] == [
        {
            "filepath": january,
            "file_name": "january.xlsx",
            "target_kind": "worksheet",
            "sheet_name": "Sheet1",
            "table_name": None,
            "source_row_count": 2,
            "matched_rows": 2,
            "auto_selected_sheet": False,
        },
        {
            "filepath": february,
            "file_name": "february.xlsx",
            "target_kind": "worksheet",
            "sheet_name": "Sheet1",
            "table_name": None,
            "source_row_count": 2,
            "matched_rows": 2,
            "auto_selected_sheet": False,
        },
    ]


def test_bulk_aggregate_workbooks_union_mode_treats_missing_columns_as_blank(tmp_path):
    primary = _create_query_workbook(
        tmp_path,
        "primary.xlsx",
        ["Region", "Sales", "Bonus"],
        [
            ("North", 10, 1),
            ("South", 4, None),
        ],
    )
    secondary = _create_query_workbook(
        tmp_path,
        "secondary.xlsx",
        ["Region", "Sales"],
        [
            ("North", 5),
            ("South", 6),
        ],
    )

    result = bulk_aggregate_workbooks_impl(
        [primary, secondary],
        sheet_name="Sheet1",
        group_by=["Region"],
        metrics=[
            {"op": "sum", "field": "Sales", "as": "total_sales"},
            {"op": "count_non_null", "field": "Bonus", "as": "bonus_rows"},
        ],
        sort_by="Region",
        schema_mode="union",
    )

    assert result["schema_mode"] == "union"
    assert result["rows"] == [
        ["North", 15, 1],
        ["South", 10, 0],
    ]
    assert result["schema_summary"]["strict_compatible"] is False


def test_bulk_aggregate_workbooks_intersect_rejects_missing_referenced_columns(tmp_path):
    primary = _create_query_workbook(
        tmp_path,
        "primary.xlsx",
        ["Region", "Sales", "Bonus"],
        [("North", 10, 1)],
    )
    secondary = _create_query_workbook(
        tmp_path,
        "secondary.xlsx",
        ["Region", "Sales"],
        [("North", 5)],
    )

    with pytest.raises(DataError, match="schema_mode 'intersect' requires column 'Bonus'"):
        bulk_aggregate_workbooks_impl(
            [primary, secondary],
            sheet_name="Sheet1",
            metrics=[{"op": "count_non_null", "field": "Bonus", "as": "bonus_rows"}],
            schema_mode="intersect",
        )


def test_bulk_aggregate_workbooks_strict_rejects_schema_drift(tmp_path):
    baseline = _create_query_workbook(
        tmp_path,
        "baseline.xlsx",
        ["Region", "Sales"],
        [("North", 10)],
    )
    drifted = _create_query_workbook(
        tmp_path,
        "drifted.xlsx",
        ["Region", "Sales", "Channel"],
        [("North", 5, "Paid")],
    )

    with pytest.raises(DataError, match="schema_mode 'strict' requires identical columns across workbooks"):
        bulk_aggregate_workbooks_impl(
            [baseline, drifted],
            sheet_name="Sheet1",
            metrics=[{"op": "sum", "field": "Sales", "as": "total_sales"}],
        )


def test_bulk_aggregate_workbooks_tool_returns_json_envelope(tmp_path):
    north = _create_query_workbook(
        tmp_path,
        "north.xlsx",
        ["Region", "Sales"],
        [("North", 10)],
    )
    south = _create_query_workbook(
        tmp_path,
        "south.xlsx",
        ["Region", "Sales"],
        [("South", 8)],
    )

    payload = _load_tool_payload(
        bulk_aggregate_workbooks_tool(
            [north, south],
            sheet_name="Sheet1",
            group_by=["Region"],
            metrics=[{"op": "sum", "field": "Sales", "as": "total_sales"}],
            sort_by="total_sales",
            sort_desc=True,
        )
    )

    assert payload["operation"] == "bulk_aggregate_workbooks"
    assert payload["data"]["workbook_count"] == 2
    assert payload["data"]["rows"] == [
        ["North", 10],
        ["South", 8],
    ]


def test_bulk_filter_workbooks_returns_records_with_source_columns(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age", "City"],
        [
            ("Alice", 30, "Helsinki"),
            ("Bob", 25, "Tampere"),
        ],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age", "City"],
        [
            ("Carol", 35, "Turku"),
            ("Dave", 28, "Oulu"),
        ],
    )

    result = bulk_filter_workbooks_impl(
        [january, february],
        sheet_name="Sheet1",
        filters=[{"field": "Age", "op": "gte", "value": 28}],
        select=["Name", "Age"],
        sort_by="Age",
        sort_desc=True,
        row_mode="objects",
        infer_schema=True,
    )

    assert result["target_kind"] == "multi_workbook"
    assert result["schema_mode"] == "strict"
    assert result["workbook_count"] == 2
    assert result["matched_rows"] == 3
    assert result["returned_rows"] == 3
    assert result["source_columns"] == ["_source_file", "_source_sheet", "_source_table"]
    assert result["records"] == [
        {
            "source_file": "february.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Carol",
            "age": 35,
        },
        {
            "source_file": "january.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Alice",
            "age": 30,
        },
        {
            "source_file": "february.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Dave",
            "age": 28,
        },
    ]


def test_bulk_filter_workbooks_can_omit_source_columns(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age"],
        [("Bob", 25)],
    )

    result = bulk_filter_workbooks_impl(
        [january, february],
        sheet_name="Sheet1",
        include_source_columns=False,
    )

    assert result["source_columns"] == []
    assert result["headers"] == ["Name", "Age"]
    assert result["rows"] == [
        ["Alice", 30],
        ["Bob", 25],
    ]


def test_bulk_filter_workbooks_union_mode_fills_missing_values_with_none(tmp_path):
    primary = _create_query_workbook(
        tmp_path,
        "primary.xlsx",
        ["Region", "Sales", "Bonus"],
        [("North", 10, 1)],
    )
    secondary = _create_query_workbook(
        tmp_path,
        "secondary.xlsx",
        ["Region", "Sales"],
        [("South", 6)],
    )

    result = bulk_filter_workbooks_impl(
        [primary, secondary],
        sheet_name="Sheet1",
        select=["Region", "Bonus"],
        schema_mode="union",
        include_source_columns=False,
        sort_by="Region",
    )

    assert result["schema_mode"] == "union"
    assert result["rows"] == [
        ["North", 1],
        ["South", None],
    ]


def test_bulk_filter_workbooks_rejects_source_column_filters(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age"],
        [("Bob", 25)],
    )

    with pytest.raises(DataError, match="cannot reference multi-workbook source columns directly"):
        bulk_filter_workbooks_impl(
            [january, february],
            sheet_name="Sheet1",
            filters=[{"field": "_source_file", "op": "eq", "value": "january.xlsx"}],
        )


def test_bulk_filter_workbooks_strict_rejects_schema_drift(tmp_path):
    baseline = _create_query_workbook(
        tmp_path,
        "baseline.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    drifted = _create_query_workbook(
        tmp_path,
        "drifted.xlsx",
        ["Name", "Age", "City"],
        [("Bob", 25, "Turku")],
    )

    with pytest.raises(DataError, match="schema_mode 'strict' requires identical columns across workbooks"):
        bulk_filter_workbooks_impl(
            [baseline, drifted],
            sheet_name="Sheet1",
        )


def test_bulk_filter_workbooks_tool_returns_json_envelope(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age"],
        [("Bob", 25)],
    )

    payload = _load_tool_payload(
        bulk_filter_workbooks_tool(
            [january, february],
            sheet_name="Sheet1",
            filters=[{"field": "Age", "op": "gte", "value": 26}],
            select=["Name"],
        )
    )

    assert payload["operation"] == "bulk_filter_workbooks"
    assert payload["data"]["matched_rows"] == 1
    assert payload["data"]["rows"] == [["january.xlsx", "Sheet1", None, "Alice"]]


def test_union_tables_returns_records_with_source_columns(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age", "City"],
        [
            ("Alice", 30, "Helsinki"),
            ("Bob", 25, "Tampere"),
        ],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age", "City"],
        [
            ("Carol", 35, "Turku"),
            ("Dave", 28, "Oulu"),
        ],
    )

    result = union_tables_impl(
        [january, february],
        sheet_name="Sheet1",
        select=["Name", "Age"],
        sort_by="Age",
        sort_desc=True,
        row_mode="objects",
        infer_schema=True,
    )

    assert result["target_kind"] == "multi_workbook"
    assert result["schema_mode"] == "strict"
    assert result["source_row_count"] == 4
    assert result["union_row_count"] == 4
    assert result["returned_rows"] == 4
    assert result["duplicates_removed"] == 0
    assert result["source_columns"] == ["_source_file", "_source_sheet", "_source_table"]
    assert result["records"] == [
        {
            "source_file": "february.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Carol",
            "age": 35,
        },
        {
            "source_file": "january.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Alice",
            "age": 30,
        },
        {
            "source_file": "february.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Dave",
            "age": 28,
        },
        {
            "source_file": "january.xlsx",
            "source_sheet": "Sheet1",
            "source_table": None,
            "name": "Bob",
            "age": 25,
        },
    ]


def test_union_tables_dedupes_after_sorting(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Id", "Status", "Updated"],
        [
            (1, "draft", 20260401),
            (2, "open", 20260402),
        ],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Id", "Status", "Updated"],
        [
            (1, "published", 20260405),
            (3, "open", 20260403),
        ],
    )

    result = union_tables_impl(
        [january, february],
        sheet_name="Sheet1",
        select=["Id", "Status", "Updated"],
        sort_by="Updated",
        sort_desc=True,
        dedupe_on=["Id"],
        include_source_columns=False,
        row_mode="objects",
    )

    assert result["dedupe_on"] == ["Id"]
    assert result["duplicates_removed"] == 1
    assert result["union_row_count"] == 3
    assert result["records"] == [
        {"id": 1, "status": "published", "updated": 20260405},
        {"id": 3, "status": "open", "updated": 20260403},
        {"id": 2, "status": "open", "updated": 20260402},
    ]


def test_union_tables_union_mode_fills_missing_values_with_none(tmp_path):
    primary = _create_query_workbook(
        tmp_path,
        "primary.xlsx",
        ["Region", "Sales", "Bonus"],
        [("North", 10, 1)],
    )
    secondary = _create_query_workbook(
        tmp_path,
        "secondary.xlsx",
        ["Region", "Sales"],
        [("South", 6)],
    )

    result = union_tables_impl(
        [primary, secondary],
        sheet_name="Sheet1",
        select=["Region", "Bonus"],
        schema_mode="union",
        include_source_columns=False,
        sort_by="Region",
    )

    assert result["schema_mode"] == "union"
    assert result["rows"] == [
        ["North", 1],
        ["South", None],
    ]
    assert result["schema_summary"]["strict_compatible"] is False


def test_union_tables_rejects_source_column_dedupe_refs(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age"],
        [("Bob", 25)],
    )

    with pytest.raises(DataError, match="cannot reference multi-workbook source columns directly"):
        union_tables_impl(
            [january, february],
            sheet_name="Sheet1",
            dedupe_on=["_source_file"],
        )


def test_union_tables_tool_returns_json_envelope(tmp_path):
    january = _create_query_workbook(
        tmp_path,
        "january.xlsx",
        ["Name", "Age"],
        [("Alice", 30)],
    )
    february = _create_query_workbook(
        tmp_path,
        "february.xlsx",
        ["Name", "Age"],
        [("Bob", 25)],
    )

    payload = _load_tool_payload(
        union_tables_tool(
            [january, february],
            sheet_name="Sheet1",
            select=["Name"],
            limit=1,
        )
    )

    assert payload["operation"] == "union_tables"
    assert payload["data"]["union_row_count"] == 2
    assert payload["data"]["returned_rows"] == 1
