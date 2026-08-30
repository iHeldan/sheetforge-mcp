import importlib.util
import io
import json
import re
import tarfile
from pathlib import Path

import pytest


def _load_release_verifier():
    script_path = Path(__file__).resolve().parents[1] / "scripts" / "verify_release_artifacts.py"
    spec = importlib.util.spec_from_file_location("sheetforge_release_verifier", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _write_tar_entry(archive: tarfile.TarFile, name: str, content: bytes = b"test") -> None:
    member = tarfile.TarInfo(name)
    member.size = len(content)
    archive.addfile(member, io.BytesIO(content))


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


def test_sdist_uses_public_allowlist_and_private_artifact_excludes():
    root = Path(__file__).resolve().parents[1]
    pyproject = (root / "pyproject.toml").read_text()
    sdist_section = pyproject.split("[tool.hatch.build.targets.sdist]", 1)[1]
    sdist_section = sdist_section.split("\n[", 1)[0]

    for required_include in ("/src", "/README.md", "/LICENSE", "/pyproject.toml"):
        assert f'"{required_include}"' in sdist_section

    for required_exclude in (
        "/CONTEXT.md",
        "/context.md",
        "/LOCAL_*.md",
        "/*.tape",
        "/*.log",
        "/.env*",
    ):
        assert f'"{required_exclude}"' in sdist_section


def test_every_distribution_build_workflow_runs_shared_artifact_verifier():
    root = Path(__file__).resolve().parents[1]
    workflows = sorted((root / ".github" / "workflows").glob("*.yml"))
    build_workflows = [
        workflow
        for workflow in workflows
        if "uv build" in workflow.read_text() or "hatch build" in workflow.read_text()
    ]

    assert build_workflows
    for workflow in build_workflows:
        assert (
            "python scripts/verify_release_artifacts.py dist" in workflow.read_text()
        ), f"{workflow.name} builds distributions without the shared artifact verifier"


@pytest.mark.parametrize(
    ("member_name", "expected_error"),
    [
        ("sheetforge_mcp-0.8.0/LOCAL_PRIVATE.md", "Forbidden file"),
        ("sheetforge_mcp-0.8.0/private-notes.txt", "Unexpected top-level file"),
        ("sheetforge_mcp-0.8.0/../secret.txt", "Unsafe archive path"),
    ],
)
def test_release_verifier_rejects_non_public_sdist_files(
    tmp_path,
    member_name,
    expected_error,
):
    verifier = _load_release_verifier()
    artifact = tmp_path / "sheetforge_mcp-0.8.0.tar.gz"
    with tarfile.open(artifact, "w:gz") as archive:
        _write_tar_entry(archive, member_name)

    with pytest.raises(AssertionError, match=expected_error):
        verifier._verify_sdist(artifact)


def test_release_verifier_rejects_sdist_symlinks(tmp_path):
    verifier = _load_release_verifier()
    artifact = tmp_path / "sheetforge_mcp-0.8.0.tar.gz"
    with tarfile.open(artifact, "w:gz") as archive:
        member = tarfile.TarInfo("sheetforge_mcp-0.8.0/README-link.md")
        member.type = tarfile.SYMTYPE
        member.linkname = "../../private/README.md"
        archive.addfile(member)

    with pytest.raises(AssertionError, match="Non-regular archive entry"):
        verifier._verify_sdist(artifact)
