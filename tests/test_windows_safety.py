"""Separate portable backend-contract checks from actual Windows acceptance."""

import errno
import os
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

import excel_mcp.workbook as workbook_module
from excel_mcp.exceptions import WorkbookError
from excel_mcp.workbook import safe_workbook


@pytest.mark.parametrize("contended", [False, True])
def test_msvcrt_nonblocking_lock_uses_one_byte_at_offset_zero(tmp_path, monkeypatch, contended):
    calls = []

    def locking(fd, mode, count):
        calls.append((mode, count, os.lseek(fd, 0, os.SEEK_CUR)))
        if contended:
            raise OSError(errno.EACCES, "simulated lock contention")

    monkeypatch.setattr(workbook_module, "fcntl", None)
    monkeypatch.setattr(workbook_module, "msvcrt", SimpleNamespace(LK_NBLCK=1, locking=locking))
    with (tmp_path / "backend.lock").open("w+b") as handle:
        handle.write(b"\0" * 8)
        handle.flush()
        os.lseek(handle.fileno(), 5, os.SEEK_SET)
        assert workbook_module._try_lock_file(handle.fileno()) is not contended
    assert calls == [(1, 1, 0)]


def test_msvcrt_unlock_uses_one_byte_at_offset_zero(tmp_path, monkeypatch):
    calls = []

    def locking(fd, mode, count):
        calls.append((mode, count, os.lseek(fd, 0, os.SEEK_CUR)))

    monkeypatch.setattr(workbook_module, "fcntl", None)
    monkeypatch.setattr(workbook_module, "msvcrt", SimpleNamespace(LK_UNLCK=2, locking=locking))
    with (tmp_path / "backend.lock").open("w+b") as handle:
        handle.write(b"\0" * 8)
        handle.flush()
        os.lseek(handle.fileno(), 5, os.SEEK_SET)
        workbook_module._unlock_file(handle.fileno())
    assert calls == [(2, 1, 0)]


def test_msvcrt_context_initializes_once_and_releases_after_error(tmp_path, monkeypatch):
    calls = []
    path = tmp_path / "backend.lock"

    def locking(fd, mode, count):
        assert os.fstat(fd).st_size == 1
        assert os.lseek(fd, 0, os.SEEK_CUR) == 0
        calls.append((fd, mode, count))

    monkeypatch.setattr(workbook_module, "fcntl", None)
    monkeypatch.setattr(
        workbook_module, "msvcrt", SimpleNamespace(LK_NBLCK=1, LK_UNLCK=2, locking=locking),
    )
    monkeypatch.setattr(workbook_module, "_workbook_lock_path", lambda _: path)
    source = tmp_path / "source.xlsx"
    with pytest.raises(RuntimeError, match="inside lock"):
        with workbook_module._exclusive_workbook_lock(source):
            assert path.read_bytes() == b"\0"
            raise RuntimeError("inside lock")
    with pytest.raises(OSError) as caught:
        os.fstat(calls[0][0])
    assert caught.value.errno == errno.EBADF
    with workbook_module._exclusive_workbook_lock(source):
        assert path.read_bytes() == b"\0"
    assert [mode for _, mode, _ in calls] == [1, 2, 1, 2]
    assert all(count == 1 for _, _, count in calls)


@pytest.mark.skipif(sys.platform != "win32", reason="requires native msvcrt byte-range locks")
def test_windows_uses_real_msvcrt_and_blocks_a_second_handle(tmp_path):
    import msvcrt

    assert workbook_module.fcntl is None
    assert workbook_module.msvcrt is msvcrt
    source = tmp_path / "source.xlsx"
    path = workbook_module._workbook_lock_path(source)
    with workbook_module._exclusive_workbook_lock(source):
        fd = os.open(path, os.O_RDWR)
        try:
            assert not workbook_module._try_lock_file(fd)
        finally:
            os.close(fd)
    fd = os.open(path, os.O_RDWR)
    try:
        assert workbook_module._try_lock_file(fd)
        workbook_module._unlock_file(fd)
    finally:
        os.close(fd)


@pytest.mark.skipif(sys.platform != "win32", reason="requires Windows open-file replacement semantics")
def test_windows_open_destination_refuses_replace_without_losing_data(tmp_workbook):
    path = Path(tmp_workbook)
    before = path.read_bytes()
    with path.open("rb") as reader:
        with pytest.raises(WorkbookError, match="Failed to save workbook atomically"):
            with safe_workbook(str(path), save=True) as wb:
                wb["Sheet1"]["D1"] = "must not persist"
        assert reader.read() == path.read_bytes() == before
    assert list(path.parent.glob(f".{path.stem}*.sheetforge-*")) == []
    with safe_workbook(str(path), save=True) as wb:
        wb["Sheet1"]["D1"] = "after close"
    with safe_workbook(str(path)) as wb:
        assert wb["Sheet1"]["D1"].value == "after close"
    assert list(path.parent.glob(f".{path.stem}*.sheetforge-*")) == []


@pytest.mark.skipif(sys.platform != "win32", reason="requires native Windows symlink privileges")
def test_windows_file_and_directory_symlinks_are_not_silently_skipped_in_ci(tmp_path):
    target = tmp_path / "source.xlsx"
    target.write_bytes(b"synthetic capability probe")
    directory = tmp_path / "directory"
    directory.mkdir()
    alias, directory_alias = tmp_path / "alias.xlsx", tmp_path / "directory-alias"
    try:
        alias.symlink_to(target)
        directory_alias.symlink_to(directory, target_is_directory=True)
    except OSError as exc:
        if getattr(exc, "winerror", None) == 1314 and os.environ.get("GITHUB_ACTIONS") != "true":
            pytest.skip("Windows symlinks require Developer Mode or symlink privilege")
        raise
    assert alias.is_symlink() and alias.samefile(target)
    assert directory_alias.is_symlink() and directory_alias.samefile(directory)
