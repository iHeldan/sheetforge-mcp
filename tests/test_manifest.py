import json
from pathlib import Path

from excel_mcp.server import mcp


def test_manifest_tool_list_matches_registered_tools():
    manifest_path = Path(__file__).resolve().parents[1] / "manifest.json"
    manifest = json.loads(manifest_path.read_text())

    manifest_tools = {tool["name"] for tool in manifest["tools"]}
    registered_tools = set(mcp._tool_manager._tools.keys())

    assert manifest_tools == registered_tools
