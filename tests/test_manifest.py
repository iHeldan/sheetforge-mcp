import json
import re
from pathlib import Path

import pytest

from excel_mcp.server import mcp


@pytest.mark.parametrize("name", [
    "read_excel_as_table", "quick_read", "query_table", "bulk_filter_workbooks",
    "union_tables", "cross_workbook_lookup", "aggregate_table",
    "bulk_aggregate_workbooks", "read_excel_table",
])
def test_tabular_read_schemas_advertise_only_supported_row_modes(name):
    field = mcp._tool_manager._tools[name].parameters["properties"]["row_mode"]
    assert field["type"] == "string" and field["default"] == "arrays"
    assert field["enum"] == ["arrays", "objects"]


def test_changeset_description_explains_operation_path_inheritance():
    description = mcp._tool_manager._tools["apply_workbook_changeset"].description
    assert "Operations inherit the outer filepath" in description
    assert "Omit filepath, dry_run and" in description


def test_manifest_tool_list_matches_registered_tools():
    manifest_path = Path(__file__).resolve().parents[1] / "manifest.json"
    manifest = json.loads(manifest_path.read_text())

    manifest_tools = {tool["name"] for tool in manifest["tools"]}
    registered_tools = set(mcp._tool_manager._tools.keys())

    assert manifest_tools == registered_tools


def test_tools_reference_lists_all_registered_tools():
    tools_doc_path = Path(__file__).resolve().parents[1] / "TOOLS.md"
    tools_doc = tools_doc_path.read_text()

    documented_tools = set(re.findall(r"^- `([a-zA-Z0-9_]+)\(", tools_doc, flags=re.MULTILINE))
    registered_tools = set(mcp._tool_manager._tools.keys())

    assert documented_tools == registered_tools


def test_public_docs_tool_counts_match_registered_tools():
    root = Path(__file__).resolve().parents[1]
    registered_tool_count = len(mcp._tool_manager._tools)

    readme = (root / "README.md").read_text()
    readme_match = re.search(r"currently exposes `(\d+)` MCP tools", readme)
    assert readme_match is not None
    assert int(readme_match.group(1)) == registered_tool_count

    landing_page = (root / "docs" / "index.html").read_text()
    landing_page_match = re.search(r"<strong>(\d+) MCP tools</strong>", landing_page)
    assert landing_page_match is not None
    assert int(landing_page_match.group(1)) == registered_tool_count
