class ExcelMCPError(Exception):
    """Base exception for Excel MCP errors."""
    pass

class WorkbookError(ExcelMCPError):
    """Raised when workbook operations fail."""
    pass

class SheetError(ExcelMCPError):
    """Raised when sheet operations fail."""
    pass

class DataError(ExcelMCPError):
    """Raised when data operations fail."""
    pass

class ValidationError(ExcelMCPError):
    """Raised when validation fails."""
    pass

class FormattingError(ExcelMCPError):
    """Raised when formatting operations fail."""
    pass

class CalculationError(ExcelMCPError):
    """Raised when formula calculations fail."""
    pass

class PivotError(ExcelMCPError):
    """Raised when pivot table operations fail."""
    pass

class ChartError(ExcelMCPError):
    """Raised when chart operations fail."""
    pass


class ResponseTooLargeError(ExcelMCPError):
    """Raised when a serialized MCP response would exceed the practical payload limit."""

    def __init__(
        self,
        message: str,
        *,
        estimated_size: int,
        limit: int,
        hints: list[str] | None = None,
    ) -> None:
        super().__init__(message)
        self.estimated_size = estimated_size
        self.limit = limit
        self.hints = hints or []


class PreconditionFailedError(ExcelMCPError):
    """Raised when a caller-provided workbook or dataset precondition no longer holds."""

    def __init__(
        self,
        message: str,
        *,
        code: str = "precondition_failed",
        details: dict | None = None,
        suggested_next_tool: str | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.details = details or {}
        self.suggested_next_tool = suggested_next_tool
