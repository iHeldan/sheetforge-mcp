#!/usr/bin/env python3
"""Verify that SheetForge release archives contain code and no local artifacts."""

from __future__ import annotations

import sys
import stat
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

FORBIDDEN_DIRECTORY_NAMES = {".git", ".notes", ".venv", "__pycache__"}
FORBIDDEN_FILE_NAMES = {"context.md"}
FORBIDDEN_SUFFIXES = {".key", ".log", ".mcpb", ".pem", ".tape"}
ALLOWED_SDIST_TOP_LEVEL = {
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
    output_dir = Path(sys.argv[1] if len(sys.argv) > 1 else "dist")
    verify_dist(output_dir)
    print(f"Verified SheetForge release artifacts in {output_dir}")
