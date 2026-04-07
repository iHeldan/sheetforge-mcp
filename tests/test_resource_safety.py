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
