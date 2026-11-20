"""Shared scalar identity and lossless standard-writer value policy."""

from datetime import date, datetime, time
from decimal import Decimal
import json
import math
from typing import Any

from openpyxl.cell.rich_text import CellRichText
from openpyxl.compat.strings import safe_string

from .exceptions import DataError


def value_key(value: Any, *, case_sensitive: bool = True) -> tuple:
    if isinstance(value, bool):
        return ("boolean", value)
    if isinstance(value, (int, float, Decimal)):
        return ("number", value)
    if isinstance(value, CellRichText):
        value = str(value)
    if isinstance(value, str):
        return ("string", value if case_sensitive else value.casefold())
    if isinstance(value, (date, datetime, time)):
        return (type(value).__name__, value.isoformat())
    if isinstance(value, (dict, list, tuple, set)):
        return (type(value).__name__, json.dumps(value, sort_keys=True, default=str))
    return (type(value).__name__, value)


def cell_values_equal(left: Any, right: Any) -> bool:
    return value_key(left) == value_key(right)


def validate_cell_value(value: Any, *, data_type: str | None = None) -> None:
    """Reject values the standard XLSX writer cannot faithfully round-trip."""
    if value is None or isinstance(value, bool):
        return
    if isinstance(value, (int, float, Decimal)):
        try:
            encoded = safe_string(value)
            decoded = Decimal(encoded)
            if not decoded.is_finite() or decoded != Decimal(str(value)):
                raise ValueError
            if not math.isfinite(float(decoded)):
                raise ValueError
        except (ArithmeticError, TypeError, ValueError, OverflowError):
            raise DataError(
                "Numeric value cannot round-trip exactly through the standard XLSX writer; "
                "use a finite representable number or explicitly supply text for an identifier"
            ) from None
        return
    if isinstance(value, (str, CellRichText)):
        text = str(value)
        if not text:
            raise DataError("Empty strings do not round-trip as text; use null to clear a cell")
        try:
            units = len(text.encode("utf-16-le")) // 2
        except UnicodeEncodeError:
            raise DataError("Text contains an unsupported Unicode scalar") from None
        if units > 32767:
            raise DataError("Cell text is limited to 32767 UTF-16 code units; no truncation was applied")
        if data_type != "s" and isinstance(value, str) and value.startswith("=") and units > 8192:
            raise DataError("Excel formulas are limited to 8192 UTF-16 code units; no truncation was applied")
        if "\r" in text:
            raise DataError("Standard writes cannot preserve carriage returns; use LF text or preserving mode")
        if any((ord(ch) < 32 and ch not in "\t\n") or ord(ch) in {0xFFFE, 0xFFFF} for ch in text):
            raise DataError("Cell text contains an unsupported XML character")
        return
    if isinstance(value, (date, datetime, time)):
        if getattr(value, "tzinfo", None) is not None or getattr(value, "microsecond", 0) % 1000:
            raise DataError("Date/time writes require naive values with at most millisecond precision")
        return
    raise DataError("Unsupported cell value type; use text, finite numbers, booleans, null or naive dates/times")


def verify_cell_writes(
    book: Any,
    writes: dict[tuple[str, int, int], Any],
    *,
    data_types: dict[tuple[str, int, int], str] | None = None,
) -> None:
    for (sheet, row, column), expected in writes.items():
        if isinstance(expected, date) and not isinstance(expected, datetime):
            expected = datetime.combine(expected, time())
        cell = book[sheet]._cells.get((row, column))
        actual = cell.value if cell is not None else None
        expected_type = (data_types or {}).get((sheet, row, column))
        if expected_type is not None and getattr(cell, "data_type", None) != expected_type:
            raise DataError(
                f"Serialized cell type at {sheet}!{row},{column} differs from the source; no changes were saved"
            )
        numeric = (int, float, Decimal)
        equal = cell_values_equal(expected, actual)
        if isinstance(expected, numeric) and isinstance(actual, numeric) and not (
            isinstance(expected, bool) or isinstance(actual, bool)
        ):
            equal = Decimal(str(expected)) == Decimal(str(actual))
        if not equal:
            raise DataError(
                f"Serialized value at {sheet}!{row},{column} differs from the request; no changes were saved"
            )
