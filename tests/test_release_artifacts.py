import json
import re
from pathlib import Path


def _package_version(root: Path) -> str:
    pyproject = (root / "pyproject.toml").read_text()
    match = re.search(r'^version = "([^"]+)"$', pyproject, flags=re.MULTILINE)
    assert match is not None
    return match.group(1)


def test_release_versions_stay_in_sync():
    root = Path(__file__).resolve().parents[1]
    version = _package_version(root)

    manifest = json.loads((root / "manifest.json").read_text())
    assert manifest["version"] == version

    readme = (root / "README.md").read_text()
    assert f"Published package release: `{version}`" in readme

    landing_page = (root / "docs" / "index.html").read_text()
    assert f"Published package release: <strong>{version}</strong>" in landing_page


def test_tracked_bundle_matches_package_version():
    root = Path(__file__).resolve().parents[1]
    version = _package_version(root)

    bundles = sorted(root.glob("sheetforge-mcp-*.mcpb"))
    assert bundles == [root / f"sheetforge-mcp-{version}.mcpb"]
