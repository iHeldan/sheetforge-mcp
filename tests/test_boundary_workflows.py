"""Actual MCP acceptance for typed copies and guarded transaction boundaries."""

import asyncio
import hashlib
import json
from pathlib import Path
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table

from tests.test_xlsx_input import _zip64_input


ROOT = Path(__file__).resolve().parents[1]


def run_boundary_workflow(parameters, tmp_path):
    source = tmp_path / "workflow.xlsx"
    book = Workbook()
    data = book.active
    data.title = "Data"
    data.append(["Id", "Region", "Amount"])
    for index in range(400):
        data.append([index, f"Region-{index:04d}", index + 1])
    data.add_table(Table(displayName="Sales", ref="A1:C401"))
    data["E1"], data["E2"] = "=A2", "#N/A"
    for cell in ("E1", "E2"):
        data[cell].data_type = "s"
        data[cell].quotePrefix = True
    book.create_sheet("Report")
    book.create_sheet("Hidden").sheet_state = "hidden"
    native_only = book.create_sheet("NativeOnly")
    for row in [("Id", "Value"), (1, 10), (2, 20)]:
        native_only.append(row)
    native_only.add_table(Table(displayName="Controls", ref="A1:B3"))
    book.save(source)
    book.close()

    array = tmp_path / "array.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active["A1"] = ArrayFormula(ref="A1:A5", text="=ROW(A1:A5)")
    book.active["A5"] = 5
    book.save(array)
    book.close()
    array_before = array.read_bytes()
    unsupported = tmp_path / "unsupported.xlsm"
    unsupported.write_bytes(source.read_bytes())
    macro_before = unsupported.read_bytes()
    directory = tmp_path / "directory.xlsx"
    directory.write_bytes(_zip64_input(conflict="count"))
    calls = []

    async def exercise():
        with (tmp_path / "stdio.log").open("w") as log:
            async with stdio_client(parameters, errlog=log) as (read, write):
                async with ClientSession(read, write) as client:
                    await client.initialize()
                    assert len((await client.list_tools()).tools) == 86

                    async def call(tool, path=source, *, ok=True, **args):
                        reply = await client.call_tool(tool, {"filepath": str(path), **args})
                        payload = json.loads("".join(item.text for item in reply.content if item.type == "text"))
                        assert payload["ok"] is ok, payload
                        calls.append({"tool": tool, "ok": ok})
                        return payload

                    await call("copy_range", sheet_name="Data", source_start="E1", source_end="E2", target_start="F1")
                    frozen = source.read_bytes()
                    await call("create_table", ok=False, sheet_name="Data", data_range="A1:B2", table_name="Overlap")
                    await call("create_table", ok=False, sheet_name="Data", data_range="XFE1:XFE2", table_name="Outside")
                    assert source.read_bytes() == frozen
                    await call("insert_rows", path=array, ok=False, sheet_name="Data", start_row=3)
                    await call("delete_sheet_rows", path=array, ok=False, sheet_name="Data", start_row=3, dry_run=True)
                    assert array.read_bytes() == array_before
                    refused = await call("write_data_to_excel", path=unsupported, ok=False, sheet_name="Data", data=[[2]])
                    assert refused["error"]["details"]["reason"] == "unsupported_extension"
                    assert unsupported.read_bytes() == macro_before
                    refused = await call("quick_read", path=directory, ok=False)
                    assert refused["error"]["details"]["reason"] == "inconsistent_zip64_directory"

                    alias = tmp_path / "source-alias.xlsx"
                    alias.symlink_to(source)
                    operations = [{"tool": "write_data_to_excel", "args": {"sheet_name": "Report", "start_cell": "D1", "data": [[1]]}}]
                    await call("apply_workbook_changeset", ok=False, operations=operations, snapshot_filepath=str(alias))
                    assert source.read_bytes() == frozen

                    await call("create_named_range", name="Window", range_ref="A2", sheet_name="Data")
                    frozen = source.read_bytes()
                    await call("create_named_range", name="Window", range_ref="B2", sheet_name="Data", replace=True, dry_run=True)
                    assert source.read_bytes() == frozen
                    await call("create_named_range", name="Window", range_ref="B2", sheet_name="Data", replace=True)

                    frozen = source.read_bytes()
                    repair = await call("apply_workbook_repairs", repair_types=["reveal_hidden_sheets"], dry_run=True)
                    assert repair["data"]["diff_basis"] == "dry_run_preview"
                    assert repair["data"]["audit_after_basis"] == "unchanged_source"
                    assert repair["data"]["diff"]["summary"]["sheet_property_change_count"] == 1
                    assert source.read_bytes() == frozen
                    applied = await call("apply_workbook_repairs", repair_types=["reveal_hidden_sheets"], dry_run=False)
                    assert applied["data"]["diff"] == repair["data"]["diff"]

                    matrix = [[row * 20 + column for column in range(20)] for row in range(40)]
                    operations[0]["args"]["data"] = matrix
                    frozen = source.read_bytes()
                    preview = (await call("apply_workbook_changeset", operations=operations, create_snapshot=False))["data"]
                    assert preview["ready_to_commit"] and source.read_bytes() == frozen
                    await call("apply_workbook_changeset", operations=operations, create_snapshot=False, mode="commit",
                               expected_workbook_sha256=preview["expected_workbook_sha256"], changeset_token=preview["changeset_token"])

                    recipe = {"version": 1, "source": {"table_name": "Sales"}, "group_by": ["Region"],
                              "metrics": [{"op": "sum", "field": "Amount", "as": "Revenue"}],
                              "output": {"sheet_name": "Report", "range_ref": "A1:B401"}}
                    frozen = source.read_bytes()
                    preview = (await call("run_workbook_recipe", recipe=recipe))["data"]
                    assert preview["ready_to_commit"] and preview["output_row_count"] == 400
                    assert source.read_bytes() == frozen
                    committed = (await call("run_workbook_recipe", recipe=recipe, mode="commit",
                                            expected_workbook_sha256=preview["expected_workbook_sha256"], recipe_token=preview["recipe_token"]))["data"]
                    snapshot = Path(committed["receipt"]["snapshot"]["path"])
                    assert snapshot.read_bytes() == frozen
                    assert committed["receipt"]["after_sha256"] == hashlib.sha256(source.read_bytes()).hexdigest()

    asyncio.run(asyncio.wait_for(exercise(), timeout=90))
    book = load_workbook(source)
    try:
        assert [(book["Data"][cell].value, book["Data"][cell].data_type) for cell in ("F1", "F2")] == [("=A2", "s"), ("#N/A", "s")]
        assert book["Hidden"].sheet_state == "visible"
        assert book.defined_names["Window"].attr_text == "'Data'!$B$2"
        assert book["Data"].tables["Sales"].ref == "A1:C401"
        assert book["NativeOnly"].tables["Controls"].ref == "A1:B3"
        assert list(book["NativeOnly"].values) == [("Id", "Value"), (1, 10), (2, 20)]
        assert [(book["Report"].cell(row, 1).value, book["Report"].cell(row, 2).value) for row in range(1, 402)] == [
            ("Region", "Revenue"), *[(f"Region-{index:04d}", index + 1) for index in range(400)],
        ]
        for row in range(40):
            assert [book["Report"].cell(row + 1, column + 4).value for column in range(20)] == [row * 20 + column for column in range(20)]
    finally:
        book.close()
    return calls


def test_guarded_boundaries_and_compact_transactions_over_actual_stdio(tmp_path):
    parameters = StdioServerParameters(
        command=sys.executable, args=["-m", "excel_mcp", "stdio"], cwd=ROOT,
        env={"PYTHONPATH": str(ROOT / "src")},
    )
    assert len(run_boundary_workflow(parameters, tmp_path)) == 17
