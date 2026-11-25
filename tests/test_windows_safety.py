"""Separate portable backend-contract checks from actual Windows acceptance."""

import errno
import hashlib
import os
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

import excel_mcp.workbook as workbook_module
from excel_mcp.exceptions import PreconditionFailedError, WorkbookError
from excel_mcp.workbook import safe_workbook


def _fake_file_state_os(monkeypatch, platform, path, *, field=None, timing=None):
    stamp = path.lstat()
    calls = []
    original = os.fstat

    def fstat(fd):
        info = original(fd)
        fields = {name: getattr(info, name) for name in (
            "st_mode", "st_dev", "st_ino", "st_size", "st_mtime_ns", "st_ctime_ns",
        )}
        fields["st_ctime_ns"] = stamp.st_ctime_ns + (10 if platform == "nt" else 0)
        if field is not None and len(calls) == timing:
            fields[field] += 1
        calls.append(fd)
        return SimpleNamespace(**fields)

    monkeypatch.setattr(workbook_module, "os", SimpleNamespace(**{
        **vars(os), "name": platform, "fstat": fstat,
    }))
    return calls


def test_file_state_accepts_native_stable_regular_file(tmp_path):
    path = tmp_path / "state.bin"
    body = b"stable native file"
    path.write_bytes(body)
    state = workbook_module._workbook_file_state(path)
    assert state == (workbook_module._guarded_source_stamp(path), hashlib.sha256(body).hexdigest())
    assert workbook_module._workbook_file_state(path) == state


def test_file_state_separates_windows_path_and_handle_ctimes(tmp_path, monkeypatch):
    path = tmp_path / "state.bin"
    body = b"stable but different ctime semantics"
    path.write_bytes(body)
    calls = _fake_file_state_os(monkeypatch, "nt", path)
    assert workbook_module._workbook_file_state(path) == (
        workbook_module._guarded_source_stamp(path), hashlib.sha256(body).hexdigest(),
    )
    assert len(calls) == 2
    with pytest.raises(OSError) as caught:
        os.fstat(calls[0])
    assert caught.value.errno == errno.EBADF


@pytest.mark.parametrize("platform", ["nt", "posix"])
@pytest.mark.parametrize("timing", [0, 1], ids=["opening", "reading"])
@pytest.mark.parametrize("field", ["st_dev", "st_ino", "st_size", "st_mtime_ns"])
def test_file_state_rejects_handle_identity_drift(tmp_path, monkeypatch, platform, timing, field):
    path = tmp_path / "state.bin"
    path.write_bytes(b"must detect a changed file")
    calls = _fake_file_state_os(monkeypatch, platform, path, field=field, timing=timing)
    with pytest.raises(PreconditionFailedError, match="changed while") as caught:
        workbook_module._workbook_file_state(path)
    assert caught.value.code == "stale_workbook"
    with pytest.raises(OSError):
        os.fstat(calls[0])


@pytest.mark.parametrize("platform", ["nt", "posix"])
def test_file_state_rejects_handle_ctime_changes_during_read(tmp_path, monkeypatch, platform):
    path = tmp_path / "state.bin"
    path.write_bytes(b"must detect metadata drift")
    _fake_file_state_os(monkeypatch, platform, path, field="st_ctime_ns", timing=1)
    with pytest.raises(PreconditionFailedError, match="changed while reading"):
        workbook_module._workbook_file_state(path)


@pytest.mark.parametrize("platform", ["nt", "posix"])
@pytest.mark.parametrize("field_index", range(5), ids=["device", "inode", "size", "mtime", "ctime"])
def test_file_state_rejects_path_stamp_drift(tmp_path, monkeypatch, platform, field_index):
    path = tmp_path / "state.bin"
    path.write_bytes(b"must detect path substitution")
    _fake_file_state_os(monkeypatch, platform, path)
    original = workbook_module._guarded_source_stamp
    calls = []

    def changed(target):
        stamp = list(original(target))
        if calls:
            stamp[field_index] += 1
        calls.append(target)
        return tuple(stamp)

    monkeypatch.setattr(workbook_module, "_guarded_source_stamp", changed)
    with pytest.raises(PreconditionFailedError, match="changed while reading"):
        workbook_module._workbook_file_state(path)


def test_file_state_retains_posix_cross_api_ctime_check(tmp_path, monkeypatch):
    path = tmp_path / "state.bin"
    path.write_bytes(b"must detect pre-open metadata drift on POSIX")
    _fake_file_state_os(monkeypatch, "posix", path, field="st_ctime_ns", timing=0)
    with pytest.raises(PreconditionFailedError, match="changed while opening"):
        workbook_module._workbook_file_state(path)


@pytest.mark.parametrize(("platform", "mode"), [("nt", "r+b"), ("posix", "rb")])
def test_file_fsync_uses_native_handle_mode_without_truncating(tmp_path, monkeypatch, platform, mode):
    path = tmp_path / "flush.bin"
    original = b"must remain unchanged"
    path.write_bytes(original)
    modes, descriptors = [], []

    class RecordedPath:
        def open(self, value):
            modes.append(value)
            return path.open(value)

    def fsync(fd):
        assert os.fstat(fd).st_size == len(original)
        descriptors.append(fd)

    monkeypatch.setattr(workbook_module, "os", SimpleNamespace(name=platform, fsync=fsync))
    workbook_module._fsync_file(RecordedPath())
    assert modes == [mode] and len(descriptors) == 1
    assert path.read_bytes() == original
    with pytest.raises(OSError) as caught:
        os.fstat(descriptors[0])
    assert caught.value.errno == errno.EBADF


def test_native_file_fsync_preserves_existing_bytes(tmp_path):
    path = tmp_path / "flush.bin"
    original = b"native durable flush"
    path.write_bytes(original)
    workbook_module._fsync_file(path)
    assert path.read_bytes() == original


@pytest.mark.parametrize("error_code", [errno.EIO, errno.EACCES])
def test_file_fsync_failures_are_not_silenced(tmp_path, monkeypatch, error_code):
    path = tmp_path / "flush.bin"
    path.write_bytes(b"original")
    expected = OSError(error_code, "injected flush failure")

    def fail(fd):
        raise expected

    monkeypatch.setattr(workbook_module, "os", SimpleNamespace(name=os.name, fsync=fail))
    with pytest.raises(OSError) as caught:
        workbook_module._fsync_file(path)
    assert caught.value is expected and path.read_bytes() == b"original"


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
