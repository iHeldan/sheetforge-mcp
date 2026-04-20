from pathlib import Path

import pytest

import excel_mcp.workbook as workbook_module
from excel_mcp.exceptions import WorkbookError
from excel_mcp.workbook import safe_workbook


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

    leftovers = list(workbook_path.parent.glob(f".{workbook_path.name}.sheetforge-*.tmp"))

    assert leftovers == []


def test_safe_workbook_raises_workbook_error_on_post_save_verify_failure(tmp_workbook, monkeypatch):
    def _boom(filepath: str) -> None:
        raise OSError("verification failed")

    monkeypatch.setattr(workbook_module, "_verify_saved_workbook", _boom)

    with pytest.raises(WorkbookError, match="verification failed"):
        with safe_workbook(tmp_workbook, save=True) as wb:
            wb["Sheet1"]["D1"] = "NewColumn"

    workbook_path = Path(tmp_workbook)
    leftovers = list(workbook_path.parent.glob(f".{workbook_path.name}.sheetforge-*.tmp"))
    assert leftovers == []
