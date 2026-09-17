import pytest

from automation.validation import require_finite_number, require_nonempty, validate_month_period


def test_require_nonempty_trims_input():
    assert require_nonempty("  July 2026  ", "period") == "July 2026"


def test_require_nonempty_rejects_blank():
    with pytest.raises(ValueError, match="non-empty"):
        require_nonempty("   ", "period")


def test_require_finite_number_rejects_non_finite_values():
    with pytest.raises(ValueError, match="finite"):
        require_finite_number(float("nan"), "sales")
    with pytest.raises(ValueError, match="finite"):
        require_finite_number(float("inf"), "sales")


def test_validate_month_period_accepts_valid_period():
    assert validate_month_period(2026, 9) == (2026, 9)


@pytest.mark.parametrize("year, month", [(1999, 1), (2101, 1), (2026, 0), (2026, 13)])
def test_validate_month_period_rejects_invalid_period(year, month):
    with pytest.raises(ValueError):
        validate_month_period(year, month)
