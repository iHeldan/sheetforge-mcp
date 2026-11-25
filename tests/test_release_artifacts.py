import importlib.util
import io
import json
import copy
import os
import re
import shlex
import subprocess
import sys
import tarfile
import zipfile
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
    _load_release_verifier().verify_manifest_runtime(manifest, expected_version=version)

    readme = (root / "README.md").read_text()
    assert f"Published package release: `{version}`" in readme

    landing_page = (root / "docs" / "index.html").read_text()
    assert f"Published package release: <strong>{version}</strong>" in landing_page


def test_mcp_dependency_excludes_incompatible_v2_sdk():
    from packaging.requirements import Requirement

    root = Path(__file__).resolve().parents[1]
    text = (root / "pyproject.toml").read_text()
    match = re.search(r'^    "(mcp\[cli\][^"]+)"', text, re.MULTILINE)
    assert match is not None
    dependency = Requirement(match[1])
    assert "1.10.1" in dependency.specifier
    assert "1.99.0" in dependency.specifier
    assert "1.10.0" not in dependency.specifier
    assert "2.0.0" not in dependency.specifier
    assert "2.3.0" not in dependency.specifier


def test_tracked_bundle_matches_package_version():
    root = Path(__file__).resolve().parents[1]
    version = _package_version(root)

    bundles = sorted(root.glob("sheetforge-mcp-*.mcpb"))
    assert bundles == [root / f"sheetforge-mcp-{version}.mcpb"]
    verifier = _load_release_verifier()
    verifier.verify_repository_bundle(root)


def test_tracked_bundle_contains_current_public_files():
    root = Path(__file__).resolve().parents[1]
    version = _package_version(root)
    verifier = _load_release_verifier()
    with zipfile.ZipFile(root / f"sheetforge-mcp-{version}.mcpb") as bundle:
        # A published bundle is immutable while unreleased source evolves. Always
        # check its internal release metadata; require exact source parity again
        # when preparing a release with an empty Unreleased section.
        bundled_manifest = json.loads(bundle.read("manifest.json"))
        bundled_readme = bundle.read("README.md").decode()
        assert bundled_manifest["version"] == version
        assert f"Published package release: `{version}`" in bundled_readme
        match = re.search(r"currently exposes `(\d+)` MCP tools", bundled_readme)
        assert match and int(match.group(1)) == len(bundled_manifest["tools"])
        unreleased = (root / "CHANGELOG.md").read_text().split("## Unreleased\n", 1)[1].split("\n## ", 1)[0].strip()
        if unreleased:
            return
        for member in verifier.ALLOWED_MCPB_MEMBERS:
            assert bundle.read(str(member)) == (root / str(member)).read_bytes()


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


def test_mcpbignore_uses_the_same_root_file_allowlist_as_verification():
    root = Path(__file__).resolve().parents[1]
    patterns = [line for line in (root / ".mcpbignore").read_text().splitlines() if line and not line.startswith("#")]
    assert patterns[0] == "*"
    assert set(patterns[1:]) == {f"!/{member}" for member in _load_release_verifier().ALLOWED_MCPB_MEMBERS}


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
    ("ref_type", "ref", "accepted"),
    [
        ("tag", "refs/tags/v{version}", True),
        ("branch", "refs/heads/main", False),
        ("branch", "refs/heads/v{version}", False),
        ("branch", "refs/tags/v{version}", False),
        ("tag", "refs/heads/v{version}", False),
        ("tag", "refs/tags/v{version}.0", False),
        ("tag", "refs/tags/{version}", False),
        ("tag", "refs/tags/v{version}-unexpected", False),
        (None, None, False),
        ("tag", None, False),
    ],
)
def test_publish_release_ref_guard(monkeypatch, ref_type, ref, accepted):
    root = Path(__file__).resolve().parents[1]
    version = _package_version(root)
    for name, value in (("GITHUB_REF_TYPE", ref_type), ("GITHUB_REF", ref)):
        if value is None:
            monkeypatch.delenv(name, raising=False)
        else:
            monkeypatch.setenv(name, value.format(version=version))
    result = subprocess.run(
        [sys.executable, "scripts/verify_release_artifacts.py", "--check-release-ref"],
        cwd=root, capture_output=True, text=True, timeout=10,
    )
    assert result.returncode == (0 if accepted else 1), result.stderr
    if accepted:
        assert "Verified release tag" in result.stdout
    else:
        assert "requires a release tag matching pyproject.toml" in result.stderr


def test_publish_checks_release_ref_unconditionally_before_building():
    root = Path(__file__).resolve().parents[1]
    workflow = (root / ".github/workflows/publish.yml").read_text()
    guard = "      - name: Verify release tag\n        run: python scripts/verify_release_artifacts.py --check-release-ref\n\n"
    install = workflow.index("      - name: Install hatch dependencies")
    assert workflow[workflow.index("      - name: Verify release tag"):install] == guard
    assert install < workflow.index("run: hatch build") < workflow.index("python scripts/verify_release_artifacts.py dist")
    assert workflow.index("python scripts/verify_release_artifacts.py dist") < workflow.index("uses: pypa/gh-action-pypi-publish@release/v1")


def test_ci_uses_the_audited_dependency_lock():
    root = Path(__file__).resolve().parents[1]
    workflow = (root / ".github/workflows/ci.yml").read_text()
    assert "uv sync --locked --extra dev" in workflow


def test_ci_adds_one_required_windows_full_suite_without_expanding_every_python():
    import itertools
    import yaml

    root = Path(__file__).resolve().parents[1]
    workflow = yaml.safe_load((root / ".github/workflows/ci.yml").read_text())
    job = workflow["jobs"]["lint-and-test"]
    matrix = job["strategy"]["matrix"]
    assert set(matrix) == {"os", "python-version", "include"}
    combinations = list(itertools.product(matrix["os"], matrix["python-version"]))
    combinations.extend((entry["os"], entry["python-version"]) for entry in matrix["include"])
    assert len(combinations) == 4
    assert set(combinations) == {
        ("ubuntu-latest", "3.10"), ("ubuntu-latest", "3.11"),
        ("ubuntu-latest", "3.12"), ("windows-latest", "3.12"),
    }
    assert job["runs-on"] == "${{ matrix.os }}"
    assert job["name"] == "${{ matrix.name || format('Lint and test (Python {0})', matrix.python-version) }}"
    assert matrix["include"][0]["name"] == "Lint and test (Windows, Python 3.12)"
    assert job["timeout-minutes"] == 10
    assert job["strategy"]["fail-fast"] is False
    assert "if" not in job
    assert not job.get("continue-on-error", False)
    commands = {step.get("run"): step for step in job["steps"] if "run" in step}
    for command in (
        "uv sync --locked --extra dev", "uv run ruff check src tests",
        "uv run pytest -q -ra", "uv build", "python scripts/verify_release_artifacts.py dist",
    ):
        assert command in commands
        assert "if" not in commands[command]
        assert not commands[command].get("continue-on-error", False)


@pytest.mark.parametrize("version", ["current", "wrong", "missing"])
def test_ci_verifies_the_actual_matrix_python_before_testing(version):
    import yaml

    root = Path(__file__).resolve().parents[1]
    workflow = yaml.safe_load((root / ".github/workflows/ci.yml").read_text())
    job = workflow["jobs"]["lint-and-test"]
    assert job["env"]["UV_PYTHON"] == "${{ matrix.python-version }}"
    steps = job["steps"]
    matches = [step for step in steps if step.get("name") == "Verify Python runtime"]
    assert len(matches) == 1
    check = matches[0]
    assert "if" not in check
    assert "env" not in check
    assert not check.get("continue-on-error", False)
    commands = [step.get("run") for step in steps]
    assert commands.index("uv sync --locked --extra dev") < steps.index(check)
    assert steps.index(check) < commands.index("uv run pytest -q -ra")
    command = shlex.split(check["run"])
    assert command[:5] == ["uv", "run", "--locked", "python", "-c"]
    assert len(command) == 6
    env = os.environ.copy()
    env.pop("UV_PYTHON", None)
    if version != "missing":
        env["UV_PYTHON"] = f"{sys.version_info.major}.{sys.version_info.minor}" if version == "current" else "0.0"
    result = subprocess.run(
        [sys.executable, "-c", command[5]], env=env, capture_output=True, text=True,
        timeout=30,
    )
    assert (result.returncode == 0) == (version == "current")
    assert sys.version.split()[0] in result.stdout


def test_ci_tests_the_built_wheel_once_after_artifact_verification():
    import yaml

    root = Path(__file__).resolve().parents[1]
    workflow = yaml.safe_load((root / ".github/workflows/ci.yml").read_text())
    steps = workflow["jobs"]["lint-and-test"]["steps"]
    matches = [step for step in steps if step.get("run") == "python tests/installed_wheel.py dist"]
    assert len(matches) == 1
    installed = matches[0]
    assert installed["if"] == "matrix.os == 'ubuntu-latest' && matrix.python-version == '3.12'"
    assert not installed.get("continue-on-error", False)
    commands = [step.get("run") for step in steps]
    assert commands.index("uv build") < commands.index("python scripts/verify_release_artifacts.py dist")
    assert commands.index("python scripts/verify_release_artifacts.py dist") < steps.index(installed)


def test_github_actions_use_node24_compatible_setup_versions():
    root = Path(__file__).resolve().parents[1]
    expected_actions = {
        "ci.yml": ("actions/setup-python@v6", "astral-sh/setup-uv@v9.0.0"),
        "publish.yml": ("actions/setup-python@v6",),
        "release-build.yml": ("actions/setup-python@v6",),
    }

    for workflow_name, required_actions in expected_actions.items():
        content = (root / ".github" / "workflows" / workflow_name).read_text()
        for action in required_actions:
            assert action in content


@pytest.mark.parametrize(
    ("member_name", "expected_error"),
    [
        ("sheetforge_mcp-0.8.0/LOCAL_PRIVATE.md", "Forbidden file"),
        ("sheetforge_mcp-0.8.0/private-notes.txt", "Unexpected top-level file"),
        ("sheetforge_mcp-0.8.0/../secret.txt", "Unsafe archive path"),
        ("sheetforge_mcp-0.8.0/docs/.DS_Store", "Forbidden file"),
        ("sheetforge_mcp-0.8.0/src/excel_mcp/._recipes.py", "Forbidden file"),
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


def test_release_verifier_rejects_private_mcpb_entries(tmp_path):
    verifier = _load_release_verifier()
    artifact = tmp_path / "sheetforge-mcp-0.8.0.mcpb"
    with zipfile.ZipFile(artifact, "w") as bundle:
        for member in verifier.ALLOWED_MCPB_MEMBERS:
            content = b'\x89PNG\r\n' if member.name == "icon.png" else b"public"
            if member.name == "manifest.json":
                content = b'{"version":"0.8.0"}'
            bundle.writestr(str(member), content)
        bundle.writestr("LOCAL_ROADMAP.md", "private")

    with pytest.raises(AssertionError, match="Forbidden file"):
        verifier._verify_mcpb(artifact, expected_version="0.8.0")


def test_release_verifier_rejects_mcpb_manifest_version_drift(tmp_path):
    verifier = _load_release_verifier()
    artifact = tmp_path / "sheetforge-mcp-0.8.0.mcpb"
    with zipfile.ZipFile(artifact, "w") as bundle:
        for member in verifier.ALLOWED_MCPB_MEMBERS:
            content = b'\x89PNG\r\n' if member.name == "icon.png" else b"public"
            if member.name == "manifest.json":
                content = b'{"version":"0.7.0"}'
            bundle.writestr(str(member), content)

    with pytest.raises(AssertionError, match="manifest version does not match"):
        verifier._verify_mcpb(artifact, expected_version="0.8.0")


def _runtime_manifest(version="0.15.3"):
    return {
        "version": version,
        "server": {"mcp_config": {
            "command": "uvx",
            "args": ["--from", f"sheetforge-mcp=={version}", "sheetforge-mcp", "stdio"],
        }},
    }


def _write_runtime_bundle(path, manifest):
    verifier = _load_release_verifier()
    with zipfile.ZipFile(path, "w") as bundle:
        for member in verifier.ALLOWED_MCPB_MEMBERS:
            content = json.dumps(manifest).encode() if member.name == "manifest.json" else b"public"
            bundle.writestr(str(member), content)


@pytest.mark.parametrize("version", ["0.15.2", "0.15.3"])
def test_release_verifier_accepts_exact_pinned_runtime(tmp_path, version):
    manifest = _runtime_manifest(version)
    artifact = tmp_path / f"sheetforge-mcp-{version}.mcpb"
    _write_runtime_bundle(artifact, manifest)
    _load_release_verifier()._verify_mcpb(artifact, expected_version=version)


@pytest.mark.parametrize(
    "change",
    [
        {"args": ["sheetforge-mcp", "stdio"]},
        {"args": ["--from", "sheetforge-mcp==0.15.2", "sheetforge-mcp", "stdio"]},
        {"args": ["--from", "sheetforge-mcp>=0.15.3", "sheetforge-mcp", "stdio"]},
        {"args": ["sheetforge-mcp@latest", "stdio"]},
        {"args": ["--from", "sheetforge-mcp==${version}", "sheetforge-mcp", "stdio"]},
        {"command": "python"},
        {"args": "--from sheetforge-mcp==0.15.3 sheetforge-mcp stdio"},
        {"platform_overrides": {"win32": {"args": ["sheetforge-mcp", "stdio"]}}},
        {"platform_overrides": []},
        {"platform_overrides": {"darwin": None}},
    ],
)
def test_release_verifier_rejects_runtime_drift_inside_bundle(tmp_path, change):
    manifest = _runtime_manifest()
    manifest["server"]["mcp_config"].update(copy.deepcopy(change))
    artifact = tmp_path / "sheetforge-mcp-0.15.3.mcpb"
    _write_runtime_bundle(artifact, manifest)
    with pytest.raises(AssertionError, match="Manifest runtime must use"):
        _load_release_verifier()._verify_mcpb(artifact, expected_version="0.15.3")


def test_release_verifier_accepts_environment_only_platform_override():
    manifest = _runtime_manifest()
    manifest["server"]["mcp_config"]["platform_overrides"] = {"darwin": {"env": {"EXAMPLE": "value"}}}
    _load_release_verifier().verify_manifest_runtime(manifest, expected_version="0.15.3")


def test_release_verifier_checks_source_manifest_before_historical_bundle(tmp_path):
    (tmp_path / "pyproject.toml").write_text('[project]\nversion = "0.15.2"\n')
    manifest = _runtime_manifest("0.15.2")
    manifest["server"]["mcp_config"]["args"] = ["sheetforge-mcp", "stdio"]
    (tmp_path / "manifest.json").write_text(json.dumps(manifest))
    with pytest.raises(AssertionError, match="Manifest runtime must use"):
        _load_release_verifier().verify_repository_bundle(tmp_path)


def test_legacy_exception_does_not_allow_rebuilt_unpinned_0152(tmp_path):
    manifest = _runtime_manifest("0.15.2")
    manifest["server"]["mcp_config"]["args"] = ["sheetforge-mcp", "stdio"]
    artifact = tmp_path / "sheetforge-mcp-0.15.2.mcpb"
    _write_runtime_bundle(artifact, manifest)
    with pytest.raises(AssertionError, match="Manifest runtime must use"):
        _load_release_verifier()._verify_mcpb(artifact, expected_version="0.15.2")


def test_repository_verifier_refuses_to_replace_published_0152_even_with_a_pin(tmp_path):
    (tmp_path / "pyproject.toml").write_text('[project]\nversion = "0.15.2"\n')
    manifest = _runtime_manifest("0.15.2")
    (tmp_path / "manifest.json").write_text(json.dumps(manifest))
    _write_runtime_bundle(tmp_path / "sheetforge-mcp-0.15.2.mcpb", manifest)
    with pytest.raises(AssertionError, match="Published 0.15.2 MCPB must remain unchanged"):
        _load_release_verifier().verify_repository_bundle(tmp_path)


@pytest.mark.parametrize("unpinned_first", [True, False])
def test_runtime_pin_cannot_be_shadowed_by_duplicate_manifests(tmp_path, unpinned_first):
    pinned = _runtime_manifest()
    unpinned = copy.deepcopy(pinned)
    unpinned["server"]["mcp_config"]["args"] = ["sheetforge-mcp", "stdio"]
    artifact = tmp_path / "sheetforge-mcp-0.15.3.mcpb"
    _write_runtime_bundle(artifact, unpinned if unpinned_first else pinned)
    with zipfile.ZipFile(artifact, "a") as bundle, pytest.warns(UserWarning, match="Duplicate name"):
        bundle.writestr("manifest.json", json.dumps(pinned if unpinned_first else unpinned))
    with pytest.raises(AssertionError, match="Duplicate MCPB entry"):
        _load_release_verifier()._verify_mcpb(artifact, expected_version="0.15.3")


@pytest.mark.parametrize("payload", [None, [], "not a manifest"])
def test_bundle_manifest_must_be_a_json_object(tmp_path, payload):
    artifact = tmp_path / "sheetforge-mcp-0.15.3.mcpb"
    _write_runtime_bundle(artifact, payload)
    with pytest.raises(AssertionError, match="manifest version does not match"):
        _load_release_verifier()._verify_mcpb(artifact, expected_version="0.15.3")
