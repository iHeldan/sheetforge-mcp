#!/usr/bin/env python3
"""Verify that SheetForge release archives contain code and no local artifacts."""

from __future__ import annotations

import sys
import stat
import tarfile
import zipfile
import json
import os
import re
import hashlib
import subprocess
from pathlib import Path, PurePosixPath

FORBIDDEN_DIRECTORY_NAMES = {".git", ".notes", ".venv", "__pycache__"}
FORBIDDEN_FILE_NAMES = {"context.md", ".ds_store"}
FORBIDDEN_SUFFIXES = {".key", ".log", ".mcpb", ".pem", ".tape"}
ALLOWED_SDIST_TOP_LEVEL = {
    ".gitattributes",
    ".gitignore",
    "assets",
    "CHANGELOG.md",
    "docs",
    "icon.png",
    "LICENSE",
    "manifest.json",
    "PKG-INFO",
    "pyproject.toml",
    "README.md",
    "scripts",
    "src",
    "tests",
    "TOOLS.md",
}
ALLOWED_MCPB_MEMBERS = {
    PurePosixPath(".python-version"),
    PurePosixPath("CHANGELOG.md"),
    PurePosixPath("LICENSE"),
    PurePosixPath("README.md"),
    PurePosixPath("icon.png"),
    PurePosixPath("manifest.json"),
}
# Preserve only the exact already-published unpinned wrapper, not future rebuilds.
LEGACY_UNPINNED_MCPB_SHA256 = "17029a278e76c89075637b3e85971e08192845ef2b6e0ce6a97b331a3f877fb3"


def _assert_safe_relative_path(relative_path: PurePosixPath, *, artifact: Path) -> None:
    if relative_path.is_absolute() or any(part in {"", ".", ".."} for part in relative_path.parts):
        raise AssertionError(f"Unsafe archive path in {artifact.name}: {relative_path}")

    lower_parts = [part.casefold() for part in relative_path.parts]
    if any(part in FORBIDDEN_DIRECTORY_NAMES for part in lower_parts):
        raise AssertionError(f"Forbidden directory in {artifact.name}: {relative_path}")

    filename = relative_path.name.casefold()
    if (
        filename in FORBIDDEN_FILE_NAMES
        or filename.startswith("local_")
        or filename.startswith(".env")
        or filename.startswith("._")
        or ".log." in filename
        or relative_path.suffix.casefold() in FORBIDDEN_SUFFIXES
    ):
        raise AssertionError(f"Forbidden file in {artifact.name}: {relative_path}")


def _verify_wheel(wheel_path: Path) -> None:
    with zipfile.ZipFile(wheel_path) as wheel:
        members: set[PurePosixPath] = set()
        entry_point_texts: list[str] = []
        for archive_member in wheel.infolist():
            member = PurePosixPath(archive_member.filename)
            if archive_member.is_dir():
                continue
            unix_mode = archive_member.external_attr >> 16
            if stat.S_IFMT(unix_mode) == stat.S_IFLNK:
                raise AssertionError(f"Symlink in {wheel_path.name}: {member}")
            top_level = member.parts[0]
            if top_level != "excel_mcp" and not (
                top_level.startswith("sheetforge_mcp-")
                and top_level.endswith(".dist-info")
            ):
                raise AssertionError(
                    f"Unexpected top-level wheel entry in {wheel_path.name}: {member}"
                )
            if member.name == "entry_points.txt" and top_level.endswith(".dist-info"):
                entry_point_texts.append(wheel.read(archive_member).decode("utf-8"))
            members.add(member)

    for member in members:
        _assert_safe_relative_path(member, artifact=wheel_path)
    if PurePosixPath("excel_mcp/__main__.py") not in members:
        raise AssertionError(f"{wheel_path.name} is missing excel_mcp/__main__.py")
    if len(entry_point_texts) != 1 or (
        "sheetforge-mcp = excel_mcp.__main__:app" not in entry_point_texts[0]
    ):
        raise AssertionError(
            f"{wheel_path.name} is missing the sheetforge-mcp console entry point"
        )


def _verify_sdist(sdist_path: Path) -> None:
    with tarfile.open(sdist_path, mode="r:gz") as archive:
        archive_members = archive.getmembers()

    members: set[PurePosixPath] = set()
    root_name: str | None = None
    for archive_member in archive_members:
        raw_member = PurePosixPath(archive_member.name)
        if archive_member.isdir():
            continue
        if not archive_member.isfile():
            raise AssertionError(
                f"Non-regular archive entry in {sdist_path.name}: {raw_member}"
            )
        if len(raw_member.parts) < 2:
            raise AssertionError(f"Unexpected sdist path in {sdist_path.name}: {raw_member}")
        if root_name is None:
            root_name = raw_member.parts[0]
        elif raw_member.parts[0] != root_name:
            raise AssertionError(
                f"Multiple sdist roots in {sdist_path.name}: {raw_member.parts[0]}"
            )
        relative_member = PurePosixPath(*raw_member.parts[1:])
        _assert_safe_relative_path(relative_member, artifact=sdist_path)
        if relative_member.parts[0] not in ALLOWED_SDIST_TOP_LEVEL:
            raise AssertionError(
                f"Unexpected top-level file in {sdist_path.name}: {relative_member}"
            )
        members.add(relative_member)

    required = {
        PurePosixPath("LICENSE"),
        PurePosixPath("README.md"),
        PurePosixPath("pyproject.toml"),
        PurePosixPath("scripts/verify_release_artifacts.py"),
        PurePosixPath("src/excel_mcp/__main__.py"),
    }
    missing = sorted(str(path) for path in required - members)
    if missing:
        raise AssertionError(f"{sdist_path.name} is missing required files: {missing}")


def _package_version(root: Path) -> str:
    pyproject = (root / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'^version = "([^"]+)"$', pyproject, flags=re.MULTILINE)
    if match is None:
        raise AssertionError("Unable to read project version from pyproject.toml")
    return match.group(1)


def verify_release_ref(root: Path) -> None:
    expected_ref = f"refs/tags/v{_package_version(root)}"
    if (
        os.environ.get("GITHUB_REF_TYPE") != "tag"
        or os.environ.get("GITHUB_REF") != expected_ref
    ):
        raise SystemExit(
            f"PyPI publishing requires a release tag matching pyproject.toml: {expected_ref}"
        )


REQUIRED_CI_JOBS = {
    "Lint and test (Python 3.10)",
    "Lint and test (Python 3.11)",
    "Lint and test (Python 3.12)",
    "Lint and test (Windows, Python 3.12)",
}
REQUIRED_CI_STEPS = (
    "Install dependencies", "Verify Python runtime", "Run Ruff", "Run pytest",
    "Build package", "Verify release artifacts",
)


def _require_ci(condition: bool, message: str) -> None:
    if not condition:
        raise SystemExit(f"Release CI check failed: {message}; publication refused")


def _github_api(endpoint: str, *, pages: bool = False, **fields):
    command = ["gh", "api", "--hostname", "github.com", "--method", "GET", endpoint]
    if pages:
        command.extend(["--paginate", "--slurp"])
    for name, value in fields.items():
        command.extend(["-f", f"{name}={value}"])
    try:
        response = subprocess.run(command, check=True, capture_output=True, text=True, timeout=30)
        return json.loads(response.stdout)
    except (OSError, subprocess.SubprocessError, ValueError):
        raise SystemExit(
            "Release CI lookup failed; require GitHub CLI and actions: read permission; publication refused"
        ) from None


def _github_items(endpoint: str, key: str, **fields) -> list[dict]:
    pages = _github_api(endpoint, pages=True, per_page=100, **fields)
    _require_ci(isinstance(pages, list), "invalid paginated GitHub response")
    result = []
    for page in pages:
        _require_ci(isinstance(page, dict) and isinstance(page.get(key), list), "invalid GitHub page")
        _require_ci(all(isinstance(item, dict) for item in page[key]), "invalid GitHub item")
        result.extend(page[key])
    return result


def _require_ci_success(item: dict, label: str) -> None:
    _require_ci(item.get("status") == "completed" and item.get("conclusion") == "success", label)


def verify_release_ci(root: Path) -> None:
    verify_release_ref(root)
    repo, sha = os.environ.get("GITHUB_REPOSITORY", ""), os.environ.get("GITHUB_SHA", "")
    _require_ci(bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*/[A-Za-z0-9][A-Za-z0-9_.-]*", repo)), "missing repository")
    _require_ci(bool(re.fullmatch(r"[0-9a-f]{40}", sha)), "missing commit SHA")
    tag = os.environ["GITHUB_REF"].removeprefix("refs/tags/")
    base = f"repos/{repo}/actions"
    runs = _github_items(
        f"{base}/workflows/ci.yml/runs", "workflow_runs", head_sha=sha, branch=tag, event="push",
    )
    runs = [run for run in runs if run.get("head_sha") == sha and run.get("head_branch") == tag
            and run.get("event") == "push" and str(run.get("path", "")).split("@")[0] == ".github/workflows/ci.yml"
            and run.get("head_repository", {}).get("full_name", "").casefold() == repo.casefold()]
    _require_ci(bool(runs), "no push CI run for this exact tag and commit")
    _require_ci(all(type(run.get(key)) is int and run[key] > 0
                    for run in runs for key in ("id", "run_number", "run_attempt")), "invalid CI run identity")
    run = max(runs, key=lambda item: item["run_number"])
    _require_ci_success(run, "latest tag CI run is not successful; wait for CI to finish")
    jobs = _github_items(f"{base}/runs/{run['id']}/jobs", "jobs", filter="all")
    latest, seen = {}, set()
    # A failed-jobs-only rerun reuses successes from earlier attempts of this run.
    for job in jobs:
        name, attempt = job.get("name"), job.get("run_attempt")
        if name not in REQUIRED_CI_JOBS:
            continue
        _require_ci(job.get("run_id") == run["id"] and job.get("head_sha") == sha
                    and type(attempt) is int and 1 <= attempt <= run["run_attempt"], "invalid CI job identity")
        _require_ci((name, attempt) not in seen, "duplicate required CI job")
        seen.add((name, attempt))
        if name not in latest or attempt > latest[name]["run_attempt"]:
            latest[name] = job
    _require_ci(set(latest) == REQUIRED_CI_JOBS, "required Linux/Windows jobs are missing")
    for name, job in latest.items():
        _require_ci_success(job, f"required job {name!r} is not successful")
        steps = REQUIRED_CI_STEPS + (("Test installed wheel",) if name == "Lint and test (Python 3.12)" else ())
        for required in steps:
            matches = [step for step in job.get("steps", []) if step.get("name") == required]
            _require_ci(len(matches) == 1, f"required step {required!r} missing or duplicated in {name!r}")
            _require_ci_success(matches[0], f"required step {required!r} did not succeed in {name!r}")
    fresh = _github_api(f"{base}/runs/{run['id']}")
    _require_ci(isinstance(fresh, dict) and all(fresh.get(key) == run.get(key) for key in (
        "id", "run_attempt", "head_sha", "head_branch", "event", "status", "conclusion",
    )), "CI changed during verification; retry after CI completes")
    print(f"Verified release CI run {run['id']} attempt {run['run_attempt']} for {tag} at {sha}")


def verify_manifest_runtime(manifest: dict, *, expected_version: str) -> None:
    if not isinstance(manifest, dict) or manifest.get("version") != expected_version:
        raise AssertionError(f"Manifest version does not match {expected_version}")
    server = manifest.get("server")
    config = server.get("mcp_config") if isinstance(server, dict) else None
    error = f"Manifest runtime must use uvx --from sheetforge-mcp=={expected_version} sheetforge-mcp stdio"
    if not isinstance(config, dict):
        raise AssertionError(error)
    overrides = config.get("platform_overrides", {})
    if not isinstance(overrides, dict):
        raise AssertionError(error)
    expected_args = ["--from", f"sheetforge-mcp=={expected_version}", "sheetforge-mcp", "stdio"]
    for override in [{}, *overrides.values()]:
        if not isinstance(override, dict):
            raise AssertionError(error)
        effective = config | override
        if effective.get("command") != "uvx" or effective.get("args") != expected_args:
            raise AssertionError(error)


def _verify_mcpb(bundle_path: Path, *, expected_version: str) -> None:
    with zipfile.ZipFile(bundle_path) as bundle:
        members: set[PurePosixPath] = set()
        manifest_payload = None
        for archive_member in bundle.infolist():
            member = PurePosixPath(archive_member.filename)
            if archive_member.is_dir():
                continue
            unix_mode = archive_member.external_attr >> 16
            if stat.S_IFMT(unix_mode) == stat.S_IFLNK:
                raise AssertionError(f"Symlink in {bundle_path.name}: {member}")
            _assert_safe_relative_path(member, artifact=bundle_path)
            if member in members:
                raise AssertionError(f"Duplicate MCPB entry in {bundle_path.name}: {member}")
            if len(member.parts) != 1 or member not in ALLOWED_MCPB_MEMBERS:
                raise AssertionError(
                    f"Unexpected MCPB entry in {bundle_path.name}: {member}"
                )
            if member == PurePosixPath("manifest.json"):
                manifest_payload = json.loads(bundle.read(archive_member).decode("utf-8"))
            members.add(member)

    if members != ALLOWED_MCPB_MEMBERS:
        missing = sorted(str(path) for path in ALLOWED_MCPB_MEMBERS - members)
        extras = sorted(str(path) for path in members - ALLOWED_MCPB_MEMBERS)
        raise AssertionError(
            f"MCPB member mismatch in {bundle_path.name}; missing={missing}, extras={extras}"
        )
    if not isinstance(manifest_payload, dict) or manifest_payload.get("version") != expected_version:
        raise AssertionError(
            f"{bundle_path.name} manifest version does not match {expected_version}"
        )
    if (
        expected_version == "0.15.2"
        and hashlib.sha256(bundle_path.read_bytes()).hexdigest() == LEGACY_UNPINNED_MCPB_SHA256
    ):
        return
    verify_manifest_runtime(manifest_payload, expected_version=expected_version)


def verify_repository_bundle(root: Path) -> None:
    version = _package_version(root)
    verify_manifest_runtime(json.loads((root / "manifest.json").read_text()), expected_version=version)
    bundles = sorted(root.glob("sheetforge-mcp-*.mcpb"))
    expected_bundle = root / f"sheetforge-mcp-{version}.mcpb"
    if bundles != [expected_bundle]:
        raise AssertionError(
            f"Expected only {expected_bundle.name}, found {[path.name for path in bundles]}"
        )
    if (
        version == "0.15.2"
        and hashlib.sha256(expected_bundle.read_bytes()).hexdigest() != LEGACY_UNPINNED_MCPB_SHA256
    ):
        raise AssertionError("Published 0.15.2 MCPB must remain unchanged; prepare a new release version")
    _verify_mcpb(expected_bundle, expected_version=version)


def verify_dist(dist_dir: Path) -> None:
    wheels = sorted(dist_dir.glob("sheetforge_mcp-*.whl"))
    sdists = sorted(dist_dir.glob("sheetforge_mcp-*.tar.gz"))
    if len(wheels) != 1 or len(sdists) != 1:
        raise AssertionError(
            f"Expected one wheel and one sdist in {dist_dir}, found "
            f"{len(wheels)} wheel(s) and {len(sdists)} sdist(s)"
        )
    _verify_wheel(wheels[0])
    _verify_sdist(sdists[0])


if __name__ == "__main__":
    repository_root = Path(__file__).resolve().parents[1]
    if sys.argv[1:] == ["--check-release-ref"]:
        verify_release_ref(repository_root)
        print("Verified release tag matches pyproject.toml")
    elif sys.argv[1:] == ["--check-release-ci"]:
        verify_release_ci(repository_root)
    else:
        output_dir = Path(sys.argv[1] if len(sys.argv) > 1 else "dist")
        verify_dist(output_dir)
        verify_repository_bundle(repository_root)
        print(f"Verified SheetForge release artifacts in {output_dir} and repository MCPB")
