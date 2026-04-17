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
