"""Offline regressions for the separate installed-wheel CI acceptance driver."""

import json
import os
from types import SimpleNamespace

import pytest

from tests import installed_wheel
from tests import test_contract_workflows
from tests.test_contract_workflows import assert_tool_inventory


def test_wheel_environment_drops_source_dev_config_and_credentials(tmp_path, monkeypatch):
    root = (tmp_path / "checkout").resolve()
    outside = tmp_path / "checkout-other/bin"
    monkeypatch.setenv("PATH", os.pathsep.join([str(root / ".venv/bin"), str(outside), ""]))
    for name in ("PYTHONPATH", "PYTHONHOME", "PYTHONUSERBASE", "PYTHONOPTIMIZE",
                 "VIRTUAL_ENV", "UV_INDEX", "UV_CONSTRAINT", "PIP_EXTRA_INDEX_URL",
                 "EXCEL_FILES_PATH", "GITHUB_TOKEN"):
        monkeypatch.setenv(name, "synthetic-unwanted-setting")
    home = tmp_path / "home"
    env = installed_wheel.clean_environment(root, home)
    assert env["PATH"] == str(outside)
    assert env["HOME"] == env["USERPROFILE"] == str(home)
    assert env["PYTHONNOUSERSITE"] == "1"
    assert "synthetic-unwanted-setting" not in env.values()
    assert not any(name.startswith(("PIP_", "UV_")) for name in env)


@pytest.mark.parametrize("count", [0, 2])
def test_wheel_selection_rejects_missing_or_ambiguous_builds(tmp_path, count):
    for number in range(count):
        (tmp_path / f"sheetforge_mcp-0.1.{number}-py3-none-any.whl").touch()
    with pytest.raises(ValueError, match="exactly one"):
        installed_wheel.find_wheel(tmp_path)


def test_wheel_selection_uses_the_exact_built_file(tmp_path):
    target = tmp_path / "sheetforge_mcp-0.15.3-py3-none-any.whl"
    target.touch()
    (tmp_path / "sheetforge_mcp-0.15.3.tar.gz").touch()
    assert installed_wheel.find_wheel(tmp_path) == target.resolve()


@pytest.mark.parametrize("tools", [[], [{"name": "one"}, {"name": "one"}], [{"name": None}], ["one"]])
def test_manifest_inventory_is_nonempty_and_unambiguous(tmp_path, monkeypatch, tools):
    monkeypatch.setattr(installed_wheel, "ROOT", tmp_path)
    (tmp_path / "manifest.json").write_text(json.dumps({"version": "1.0", "tools": tools}))
    with pytest.raises(ValueError, match="tool names"):
        installed_wheel.manifest_expectations()


@pytest.mark.parametrize("actual", [["one", "wrong"], ["one", "one"], ["one"]])
def test_actual_inventory_requires_exact_names_not_only_count(actual):
    with pytest.raises(AssertionError, match="manifest|duplicate"):
        assert_tool_inventory([SimpleNamespace(name=name) for name in actual], {"one", "two"})


def test_actual_inventory_accepts_exact_names_in_any_order():
    assert_tool_inventory([SimpleNamespace(name=name) for name in ["two", "one"]], {"one", "two"})


def test_contract_workflow_uses_the_supplied_cli_parameters(tmp_path, monkeypatch):
    sentinel, received = object(), []

    def intercept(parameters, **kwargs):
        received.append(parameters)
        raise RuntimeError("intercepted CLI startup")

    monkeypatch.setattr(test_contract_workflows, "stdio_client", intercept)
    with pytest.raises(RuntimeError, match="intercepted CLI startup"):
        test_contract_workflows.exercise_contract_workflow(tmp_path, installed=True, server_parameters=sentinel)
    assert len(received) == 1 and received[0] is sentinel


@pytest.mark.parametrize("case", ["checkout", "unrelated"])
def test_probe_never_writes_into_checkout_or_an_unrelated_directory(tmp_path, monkeypatch, case):
    monkeypatch.chdir(tmp_path)
    root = tmp_path if case == "checkout" else tmp_path / "checkout"
    prefix = tmp_path / "venv" if case == "checkout" else tmp_path / "other/venv"
    monkeypatch.setattr(installed_wheel, "ROOT", root)
    with pytest.raises(RuntimeError, match="scratch directory outside"):
        installed_wheel.probe(prefix)


def test_missing_console_script_never_falls_back_to_module_or_path(tmp_path, monkeypatch):
    monkeypatch.setenv("PATH", str(tmp_path))
    with pytest.raises(RuntimeError, match="own sheetforge-mcp CLI"):
        installed_wheel.console_command(tmp_path)


def test_install_and_probe_use_fresh_environment_without_dev_lock(tmp_path, monkeypatch):
    root = tmp_path / "checkout"
    root.mkdir()
    (root / "manifest.json").write_text(json.dumps({"version": "1.0", "tools": [{"name": "one"}]}))
    dist = root / "dist"
    dist.mkdir()
    wheel = dist / "sheetforge_mcp-1.0-py3-none-any.whl"
    wheel.touch()
    monkeypatch.setattr(installed_wheel, "ROOT", root)
    monkeypatch.setattr(installed_wheel.shutil, "which", lambda _: "/synthetic/uv")
    calls, builders = [], []

    def builder(**kwargs):
        builders.append(kwargs)
        return SimpleNamespace(create=lambda prefix: prefix.mkdir())

    def run(command, **kwargs):
        calls.append((command, kwargs))

    monkeypatch.setattr(installed_wheel.venv, "EnvBuilder", builder)
    monkeypatch.setattr(installed_wheel.subprocess, "run", run)
    installed_wheel.verify_wheel(dist)
    assert builders == [{"system_site_packages": False, "with_pip": False, "symlinks": os.name != "nt"}]
    assert len(calls) == 3
    install, check, probe = [command for command, _ in calls]
    assert install[:4] == ["/synthetic/uv", "--no-config", "pip", "install"]
    assert install[-1] == str(wheel.resolve())
    assert install[install.index("--index-url") + 1] == "https://pypi.org/simple"
    assert check[:4] == ["/synthetic/uv", "--no-config", "pip", "check"]
    assert probe[1:3] == ["-I", "-B"] and probe[-2] == "--probe"
    assert "--extra" not in install and "--group" not in install and "--no-deps" not in install
    for _, options in calls:
        assert not options["cwd"].is_relative_to(root)
        assert options["check"] is True and options["timeout"] > 0
        assert "PYTHONPATH" not in options["env"] and "VIRTUAL_ENV" not in options["env"]


@pytest.mark.parametrize("failed_stage", [0, 1, 2])
def test_installer_errors_are_not_reported_as_success(tmp_path, monkeypatch, failed_stage):
    monkeypatch.setattr(installed_wheel, "manifest_expectations", lambda: ("1.0", {"one"}))
    monkeypatch.setattr(installed_wheel, "find_wheel", lambda _: tmp_path / "one.whl")
    monkeypatch.setattr(installed_wheel.shutil, "which", lambda _: "/synthetic/uv")
    monkeypatch.setattr(installed_wheel.venv, "EnvBuilder", lambda **_: SimpleNamespace(create=lambda p: p.mkdir()))

    calls = []

    def fail(*args, **kwargs):
        calls.append(args[0])
        if len(calls) == failed_stage + 1:
            raise installed_wheel.subprocess.CalledProcessError(1, args[0])

    monkeypatch.setattr(installed_wheel.subprocess, "run", fail)
    with pytest.raises(installed_wheel.subprocess.CalledProcessError):
        installed_wheel.verify_wheel(tmp_path)
    assert len(calls) == failed_stage + 1


@pytest.mark.parametrize("case", [
    "valid", "no-venv", "not-isolated", "optimized", "user-site", "overrides",
    "checkout-path", "wrong-origin", "wrong-version", "outside-metadata", "dev-tools",
])
def test_runtime_isolation_guards_fail_closed(tmp_path, monkeypatch, case):
    prefix = (tmp_path / "venv").resolve()
    root = (tmp_path / "checkout").resolve()
    flags = SimpleNamespace(isolated=1, optimize=0)
    runtime = SimpleNamespace(prefix=str(prefix), base_prefix="/base", flags=flags, path=["/stdlib"])
    user_site = SimpleNamespace(ENABLE_USER_SITE=False)
    package = SimpleNamespace(origin=str(prefix / "site-packages/excel_mcp/__init__.py"))
    distribution = SimpleNamespace(version="1.0", locate_file=lambda _: prefix / "site-packages")
    specs = {"excel_mcp": package}
    for name in ("PYTHONPATH", "PYTHONHOME", "VIRTUAL_ENV"):
        monkeypatch.delenv(name, raising=False)
    if case == "no-venv":
        runtime.base_prefix = runtime.prefix
    elif case == "not-isolated":
        flags.isolated = 0
    elif case == "optimized":
        flags.optimize = 1
    elif case == "user-site":
        user_site.ENABLE_USER_SITE = True
    elif case == "overrides":
        monkeypatch.setenv("PYTHONPATH", str(root / "src"))
    elif case == "checkout-path":
        runtime.path.append(str(root / "src"))
    elif case == "wrong-origin":
        package.origin = str(root / "src/excel_mcp/__init__.py")
    elif case == "wrong-version":
        distribution.version = "2.0"
    elif case == "outside-metadata":
        distribution.locate_file = lambda _: root / ".venv/site-packages"
    elif case == "dev-tools":
        specs["pytest"] = SimpleNamespace(origin="/dev/pytest.py")
    monkeypatch.setattr(installed_wheel, "ROOT", root)
    monkeypatch.setattr(installed_wheel, "sys", runtime)
    monkeypatch.setattr(installed_wheel, "site", user_site)
    monkeypatch.setattr(installed_wheel, "importlib", SimpleNamespace(
        util=SimpleNamespace(find_spec=lambda name: specs.get(name)),
        metadata=SimpleNamespace(distribution=lambda _: distribution),
    ))
    if case == "valid":
        installed_wheel.verify_installed_runtime(prefix, "1.0")
    else:
        with pytest.raises(RuntimeError):
            installed_wheel.verify_installed_runtime(prefix, "1.0")
