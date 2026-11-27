"""CI entrypoint: install the built wheel, then exercise its real CLI in isolation."""

from __future__ import annotations

import argparse
import importlib.metadata
import importlib.util
import json
import os
from pathlib import Path
import shutil
import site
import subprocess
import sys
import tempfile
import venv


ROOT = Path(__file__).resolve().parents[1]


def clean_environment(root: Path, home: Path) -> dict[str, str]:
    keep = ("PATH", "SYSTEMROOT", "WINDIR", "COMSPEC", "PATHEXT", "TMP", "TEMP", "TMPDIR", "LANG", "LC_ALL")
    env = {name: os.environ[name] for name in keep if name in os.environ}
    paths = [entry for entry in env.get("PATH", "").split(os.pathsep) if entry]
    env["PATH"] = os.pathsep.join(entry for entry in paths if not Path(entry).resolve().is_relative_to(root))
    env.update(
        HOME=str(home), USERPROFILE=str(home),
        APPDATA=str(home / "AppData/Roaming"), LOCALAPPDATA=str(home / "AppData/Local"),
        PYTHONNOUSERSITE="1", PYTHONSAFEPATH="1", PYTHONDONTWRITEBYTECODE="1",
    )
    return env


def manifest_expectations() -> tuple[str, set[str]]:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    if not isinstance(manifest, dict):
        raise ValueError("Manifest must be an object declaring a version and tool names")
    version = manifest.get("version")
    tools = manifest.get("tools")
    if not isinstance(version, str) or not version or not isinstance(tools, list) or not tools:
        raise ValueError("Manifest must declare a version and tool names")
    names = [tool.get("name") if isinstance(tool, dict) else None for tool in tools]
    if any(not isinstance(name, str) or not name for name in names) or len(set(names)) != len(names):
        raise ValueError("Manifest tool names must be nonempty and unique")
    return version, set(names)


def find_wheel(dist: Path) -> Path:
    wheels = sorted(dist.glob("*.whl"))
    if len(wheels) != 1 or not wheels[0].name.startswith("sheetforge_mcp-"):
        raise ValueError("Expected exactly one built SheetForge wheel")
    return wheels[0].resolve(strict=True)


def environment_paths(prefix: Path) -> tuple[Path, Path]:
    if os.name == "nt":
        return prefix / "Scripts/python.exe", prefix / "Scripts/sheetforge-mcp.exe"
    return prefix / "bin/python", prefix / "bin/sheetforge-mcp"


def console_command(prefix: Path) -> Path:
    _, command = environment_paths(prefix)
    if not command.is_file() or not command.resolve().is_relative_to(prefix):
        raise RuntimeError("Built wheel did not install its own sheetforge-mcp CLI")
    return command


def verify_installed_runtime(prefix: Path, expected_version: str) -> None:
    if sys.prefix == sys.base_prefix or Path(sys.prefix).resolve() != prefix:
        raise RuntimeError("Wheel probe must run in its fresh virtual environment")
    if not sys.flags.isolated or sys.flags.optimize or site.ENABLE_USER_SITE:
        raise RuntimeError("Wheel probe requires isolated, non-optimized Python")
    if any(name in os.environ for name in ("PYTHONPATH", "PYTHONHOME", "VIRTUAL_ENV")):
        raise RuntimeError("Wheel probe inherited Python environment overrides")
    if any(Path(entry).resolve().is_relative_to(ROOT) for entry in sys.path):
        raise RuntimeError("Wheel probe can import from the checkout")
    spec = importlib.util.find_spec("excel_mcp")
    if spec is None or spec.origin is None or not Path(spec.origin).resolve().is_relative_to(prefix):
        raise RuntimeError("SheetForge imports do not come from the installed wheel")
    distribution = importlib.metadata.distribution("sheetforge-mcp")
    if distribution.version != expected_version:
        raise RuntimeError("Installed wheel version does not match manifest")
    if not Path(distribution.locate_file("")).resolve().is_relative_to(prefix):
        raise RuntimeError("SheetForge metadata is outside the fresh environment")
    if any(importlib.util.find_spec(name) is not None for name in ("pytest", "ruff")):
        raise RuntimeError("Wheel environment contains development tools")


def probe(prefix: Path) -> None:
    if Path.cwd().resolve() != prefix.parent or prefix.parent.is_relative_to(ROOT):
        raise RuntimeError("Wheel probe must run in its own scratch directory outside the checkout")
    version, names = manifest_expectations()
    verify_installed_runtime(prefix, version)
    command = console_command(prefix)
    work = Path.cwd() / "contract-workflow"
    work.mkdir()
    # Load only the client test driver explicitly; never add the checkout to sys.path.
    spec = importlib.util.spec_from_file_location("contract_workflow", ROOT / "tests/test_contract_workflows.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    from mcp import StdioServerParameters

    parameters = StdioServerParameters(command=str(command), args=["stdio"], cwd=work, env=dict(os.environ))
    result = module.exercise_contract_workflow(
        work, installed=True, server_parameters=parameters, expected_tools=names,
    )
    print(json.dumps({"status": "passed", "version": version, "cli": "passed",
                      "isolation": "passed", "manifest_tools": len(names), "workflow": result}))


def verify_wheel(dist: Path) -> None:
    wheel = find_wheel(dist)
    manifest_expectations()
    uv = shutil.which("uv")
    if uv is None:
        raise RuntimeError("uv is required to install the wheel's runtime dependencies")
    with tempfile.TemporaryDirectory(prefix="sheetforge-wheel-") as folder:
        scratch = Path(folder).resolve()
        if scratch.is_relative_to(ROOT):
            raise RuntimeError("Wheel acceptance scratch directory must be outside the checkout")
        home, prefix = scratch / "home", scratch / "venv"
        home.mkdir()
        env = clean_environment(ROOT, home)
        # POSIX interpreter builds may resolve libpython relative to the binary.
        # Match the venv CLI default instead of relocating that binary by copy.
        venv.EnvBuilder(system_site_packages=False, with_pip=False, symlinks=os.name != "nt").create(prefix)
        python, _ = environment_paths(prefix)
        subprocess.run(
            [uv, "--no-config", "pip", "install", "--python", str(python),
             "--index-url", "https://pypi.org/simple", str(wheel)],
            cwd=scratch, env=env, check=True, timeout=180,
        )
        subprocess.run(
            [uv, "--no-config", "pip", "check", "--python", str(python)],
            cwd=scratch, env=env, check=True, timeout=30,
        )
        subprocess.run(
            [str(python), "-I", "-B", str(Path(__file__).resolve()), "--probe", str(prefix)],
            cwd=scratch, env=env, check=True, timeout=90,
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="Directory containing exactly one built wheel")
    parser.add_argument("--probe", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.probe:
        probe(args.path.resolve())
    else:
        verify_wheel(args.path.resolve())


if __name__ == "__main__":
    main()
