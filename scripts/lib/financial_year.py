from __future__ import annotations

from datetime import date


def get_financial_year(value: date, start_month: int = 4) -> str:
    if not 1 <= start_month <= 12:
        raise ValueError("start_month must be between 1 and 12")

    if value.month >= start_month:
        start_year = value.year
    else:
        start_year = value.year - 1

    return f"FY{start_year}-{start_year + 1}"


def get_run_folder(value: date) -> str:
    return value.isoformat()
