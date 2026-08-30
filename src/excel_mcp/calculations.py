from collections import Counter
from typing import Any
import logging

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import range_boundaries
from .workbook import require_worksheet, safe_workbook
from .cell_utils import validate_cell_reference
from .exceptions import ValidationError, CalculationError
from .validation import UNSAFE_FORMULA_FUNCTIONS, validate_formula

logger = logging.getLogger(__name__)
VOLATILE_FORMULA_FUNCTIONS = {
    "CELL",
    "INFO",
    "INDIRECT",
    "NOW",
    "OFFSET",
    "RAND",
    "RANDBETWEEN",
    "TODAY",
}


def _normalize_formula(formula: str) -> str:
    normalized = str(formula or "").strip()
    if not normalized:
        raise CalculationError("formula is required")
    if not normalized.startswith("="):
        normalized = f"={normalized}"
    return normalized


def _classify_reference_token(token_value: str) -> str:
    local_reference = token_value.rsplit("!", 1)[-1].replace("$", "")
    if "[" in local_reference:
        return "structured_reference"

    try:
        range_boundaries(local_reference)
    except ValueError:
        return "named_or_identifier"
    return "worksheet_range"


def inspect_formula(formula: str) -> dict[str, Any]:
    """Inspect a formula string without requiring workbook context."""
    try:
        normalized_formula = _normalize_formula(formula)

        tokenization_error = None
        token_values: list[dict[str, Any]] = []
        reference_tokens: list[str] = []
        function_names: list[str] = []

        try:
            tokenizer = Tokenizer(normalized_formula)
        except Exception as exc:
            tokenization_error = str(exc)
        else:
            for token in tokenizer.items:
                token_record = {
                    "type": token.type,
                    "subtype": token.subtype,
                    "value": str(token.value),
                }
                token_values.append(token_record)

                if token.type == "FUNC" and token.subtype == "OPEN":
                    function_name = str(token.value).rstrip("(").upper()
                    if function_name:
                        function_names.append(function_name)
                elif token.type == "OPERAND" and token.subtype == "RANGE":
                    token_value = str(token.value).strip()
                    if token_value:
                        reference_tokens.append(token_value)

        function_counter = Counter(function_names)
        references = [
            {
                "token": token_value,
                "reference_type": _classify_reference_token(token_value),
                "sheet_qualified": "!" in token_value,
            }
            for token_value in reference_tokens
        ]
        volatile_functions = sorted(
            function_name
            for function_name in function_counter
            if function_name in VOLATILE_FORMULA_FUNCTIONS
        )
        unsafe_functions = sorted(
            function_name
            for function_name in function_counter
            if function_name in UNSAFE_FORMULA_FUNCTIONS
        )
        literal_token_count = sum(
            1
            for token in token_values
            if token["type"] == "OPERAND" and token["subtype"] != "RANGE"
        )
        syntax_valid = tokenization_error is None

        result = {
            "formula": normalized_formula,
            "syntax_valid": syntax_valid,
            "summary": {
                "token_count": len(token_values),
                "function_count": sum(function_counter.values()),
                "unique_function_count": len(function_counter),
                "reference_count": len(references),
                "literal_token_count": literal_token_count,
                "uses_volatile_functions": bool(volatile_functions),
                "uses_unsafe_functions": bool(unsafe_functions),
            },
            "functions": [
                {
                    "name": function_name,
                    "count": function_counter[function_name],
                    "volatile": function_name in VOLATILE_FORMULA_FUNCTIONS,
                    "unsafe": function_name in UNSAFE_FORMULA_FUNCTIONS,
                }
                for function_name in sorted(function_counter)
            ],
            "references": references,
        }
        if tokenization_error is not None:
            result["tokenization_error"] = tokenization_error
            result["warnings"] = [
                "Formula tokenization failed; function and reference analysis may be incomplete."
            ]
        return result

    except CalculationError:
        raise
    except Exception as e:
        logger.error(f"Failed to inspect formula: {e}")
        raise CalculationError(str(e))

def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> dict[str, Any]:
    """Apply any Excel formula to a cell."""
    try:
        if not validate_cell_reference(cell):
            raise ValidationError(f"Invalid cell reference: {cell}")

        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = f'={formula}'

        # Validate formula syntax
        is_valid, message = validate_formula(formula)
        if not is_valid:
            raise CalculationError(f"Invalid formula syntax: {message}")

        with safe_workbook(filepath, save=True) as wb:
            sheet = require_worksheet(
                wb,
                sheet_name,
                error_cls=ValidationError,
                operation="applying formulas",
            )

            try:
                # Apply formula to the cell
                cell_obj = sheet[cell]
                cell_obj.value = formula
            except Exception as e:
                raise CalculationError(f"Failed to apply formula to cell: {str(e)}")

        return {
            "message": f"Applied formula '{formula}' to cell {cell}",
            "cell": cell,
            "formula": formula
        }

    except (ValidationError, CalculationError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formula: {e}")
        raise CalculationError(str(e))
