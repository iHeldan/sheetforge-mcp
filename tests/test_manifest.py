import json
import re
from pathlib import Path

from excel_mcp.server import mcp


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
