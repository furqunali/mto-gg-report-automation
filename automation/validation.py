"""Pure validation helpers for the monthly report pipeline."""

from math import isfinite


def require_nonempty(value: str, field_name: str) -> str:
    """Return a trimmed string or raise a useful configuration error."""
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} must be a non-empty string")
    return value.strip()


def require_finite_number(value: float, field_name: str) -> float:
    """Validate numeric pipeline inputs and reject NaN/infinity."""
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be numeric") from exc
    if not isfinite(number):
        raise ValueError(f"{field_name} must be finite")
    return number


def validate_month_period(year: int, month: int) -> tuple[int, int]:
    """Validate a calendar period used for report naming and trend updates."""
    if not isinstance(year, int) or not 2000 <= year <= 2100:
        raise ValueError("year must be between 2000 and 2100")
    if not isinstance(month, int) or not 1 <= month <= 12:
        raise ValueError("month must be between 1 and 12")
    return year, month
