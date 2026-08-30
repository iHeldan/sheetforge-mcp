from pathlib import Path
from threading import Event, Thread

import pytest

import excel_mcp.workbook as workbook_module
from excel_mcp.exceptions import WorkbookError
from excel_mcp.workbook import safe_workbook


def _assert_no_sheetforge_artifacts(workbook_path: Path) -> None:
    temp_leftovers = list(
        workbook_path.parent.glob(
            f".{workbook_path.stem}.sheetforge-*{workbook_path.suffix or '.xlsx'}"
        )
    )
    backup_leftovers = list(
        workbook_path.parent.glob(f".{workbook_path.name}.sheetforge-backup-*.bak")
    )
    assert temp_leftovers == []
    assert backup_leftovers == []


def test_safe_workbook_closes_on_success(tmp_workbook):
    """Workbook should be closed after successful context manager exit."""
    with safe_workbook(tmp_workbook) as wb:
        assert "Sheet1" in wb.sheetnames


def test_safe_workbook_closes_on_error(tmp_workbook):
    """Workbook should be closed even when an exception occurs."""
    try:
        with safe_workbook(tmp_workbook) as wb:
            raise ValueError("simulated error")
    except ValueError:
        pass
    with safe_workbook(tmp_workbook) as wb:
        assert "Sheet1" in wb.sheetnames


def test_safe_workbook_saves_when_requested(tmp_workbook):
    """Workbook should save changes when save=True."""
    with safe_workbook(tmp_workbook, save=True) as wb:
        ws = wb["Sheet1"]
        ws["D1"] = "NewColumn"

    with safe_workbook(tmp_workbook) as wb:
        assert wb["Sheet1"]["D1"].value == "NewColumn"


def test_safe_workbook_does_not_save_on_error(tmp_workbook):
    """save=True should only persist changes when the block exits successfully."""
    try:
        with safe_workbook(tmp_workbook, save=True) as wb:
            wb["Sheet1"]["D1"] = "UnsavedColumn"
            raise ValueError("simulated error")
    except ValueError:
        pass

    with safe_workbook(tmp_workbook) as wb:
        assert wb["Sheet1"]["D1"].value is None


def test_safe_workbook_atomic_save_leaves_no_temp_artifacts(tmp_workbook):
    workbook_path = Path(tmp_workbook)

    with safe_workbook(tmp_workbook, save=True) as wb:
        wb["Sheet1"]["D1"] = "NewColumn"

    _assert_no_sheetforge_artifacts(workbook_path)


def test_safe_workbook_raises_workbook_error_on_post_save_verify_failure(tmp_workbook, monkeypatch):
    def _boom(filepath: str) -> None:
        raise OSError("verification failed")

    monkeypatch.setattr(workbook_module, "_verify_saved_workbook", _boom)

    with pytest.raises(WorkbookError, match="verification failed"):
        with safe_workbook(tmp_workbook, save=True) as wb:
            wb["Sheet1"]["D1"] = "NewColumn"

    workbook_path = Path(tmp_workbook)
    _assert_no_sheetforge_artifacts(workbook_path)

    with safe_workbook(tmp_workbook) as wb:
        assert wb["Sheet1"]["D1"].value is None


def test_safe_workbook_serializes_concurrent_mutations(tmp_workbook):
    first_inside = Event()
    release_first = Event()
    second_inside = Event()
    errors: list[BaseException] = []

    def first_writer() -> None:
        try:
            with safe_workbook(tmp_workbook, save=True) as wb:
                wb["Sheet1"]["B1"] = "left"
                first_inside.set()
                assert release_first.wait(timeout=5)
        except BaseException as exc:  # pragma: no cover - surfaced below
            errors.append(exc)

    def second_writer() -> None:
        try:
            assert first_inside.wait(timeout=5)
            with safe_workbook(tmp_workbook, save=True) as wb:
                second_inside.set()
                wb["Sheet1"]["C1"] = "right"
        except BaseException as exc:  # pragma: no cover - surfaced below
            errors.append(exc)

    first_thread = Thread(target=first_writer)
    second_thread = Thread(target=second_writer)
    first_thread.start()
    second_thread.start()

    assert first_inside.wait(timeout=5)
    assert not second_inside.wait(timeout=0.2)
    release_first.set()
    first_thread.join(timeout=5)
    second_thread.join(timeout=5)

    assert not first_thread.is_alive()
    assert not second_thread.is_alive()
    assert errors == []
    assert second_inside.is_set()

    with safe_workbook(tmp_workbook) as wb:
        assert wb["Sheet1"]["B1"].value == "left"
        assert wb["Sheet1"]["C1"].value == "right"


def test_safe_workbook_preserves_symlink_when_saving(tmp_workbook, tmp_path):
    workbook_path = Path(tmp_workbook)
    alias_path = tmp_path / "alias.xlsx"
    try:
        alias_path.symlink_to(workbook_path)
    except OSError:
        pytest.skip("Symlinks are not available on this platform")

    with safe_workbook(str(alias_path), save=True) as wb:
        wb["Sheet1"]["D1"] = "via alias"

    assert alias_path.is_symlink()
    with safe_workbook(tmp_workbook) as wb:
        assert wb["Sheet1"]["D1"].value == "via alias"
