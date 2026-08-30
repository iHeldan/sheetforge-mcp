import pytest

import excel_mcp.server as server_module


def test_stdio_paths_must_be_absolute(monkeypatch):
    monkeypatch.setattr(server_module, "EXCEL_FILES_PATH", None)

    with pytest.raises(ValueError, match="must be absolute"):
        server_module.get_excel_path("report.xlsx")


def test_file_transport_resolves_paths_inside_base(tmp_path, monkeypatch):
    base_path = tmp_path / "excel-files"
    base_path.mkdir()
    monkeypatch.setattr(server_module, "EXCEL_FILES_PATH", str(base_path))

    expected = base_path / "nested" / "report.xlsx"
    assert server_module.get_excel_path("nested/report.xlsx") == str(expected)
    assert server_module.get_excel_path(str(expected)) == str(expected)


@pytest.mark.parametrize("filepath", ["../outside.xlsx", "nested/../../outside.xlsx"])
def test_file_transport_rejects_parent_traversal(tmp_path, monkeypatch, filepath):
    base_path = tmp_path / "excel-files"
    base_path.mkdir()
    monkeypatch.setattr(server_module, "EXCEL_FILES_PATH", str(base_path))

    with pytest.raises(ValueError, match="must stay within"):
        server_module.get_excel_path(filepath)


def test_file_transport_rejects_absolute_path_outside_base(tmp_path, monkeypatch):
    base_path = tmp_path / "excel-files"
    base_path.mkdir()
    monkeypatch.setattr(server_module, "EXCEL_FILES_PATH", str(base_path))

    with pytest.raises(ValueError, match="must stay within"):
        server_module.get_excel_path(str(tmp_path / "outside.xlsx"))


def test_file_transport_rejects_symlink_escape(tmp_path, monkeypatch):
    base_path = tmp_path / "excel-files"
    outside_path = tmp_path / "outside"
    base_path.mkdir()
    outside_path.mkdir()
    symlink_path = base_path / "linked"
    try:
        symlink_path.symlink_to(outside_path, target_is_directory=True)
    except OSError:
        pytest.skip("Symlinks are not available on this platform")
    monkeypatch.setattr(server_module, "EXCEL_FILES_PATH", str(base_path))

    with pytest.raises(ValueError, match="must stay within"):
        server_module.get_excel_path("linked/report.xlsx")


@pytest.mark.parametrize("host", ["0.0.0.0", "::", "192.168.1.10", "sheetforge.local"])
def test_remote_binding_requires_explicit_opt_in(monkeypatch, host):
    monkeypatch.setenv("FASTMCP_HOST", host)
    monkeypatch.delenv("SHEETFORGE_ALLOW_REMOTE", raising=False)

    with pytest.raises(RuntimeError, match="SHEETFORGE_ALLOW_REMOTE=true"):
        server_module._validate_network_binding()


@pytest.mark.parametrize("host", ["127.0.0.1", "127.0.0.2", "::1", "localhost"])
def test_loopback_binding_is_allowed(monkeypatch, host):
    monkeypatch.setenv("FASTMCP_HOST", host)
    monkeypatch.delenv("SHEETFORGE_ALLOW_REMOTE", raising=False)

    server_module._validate_network_binding()


def test_remote_binding_can_be_enabled_explicitly(monkeypatch):
    monkeypatch.setenv("FASTMCP_HOST", "0.0.0.0")
    monkeypatch.setenv("SHEETFORGE_ALLOW_REMOTE", "true")

    server_module._validate_network_binding()
