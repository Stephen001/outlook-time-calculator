from __future__ import annotations

from datetime import date

import pytest

from outlook_time_calculator.models import CategoryRow, Meeting


def meeting(
    categories: list[str],
    duration_minutes: int,
    response_status: str = "Accepted",
    subject: str = "Test Meeting",
) -> Meeting:
    return Meeting(
        subject=subject,
        start=date(2025, 1, 6),
        end=date(2025, 1, 6),
        duration_minutes=duration_minutes,
        categories=categories,
        response_status=response_status,
    )


def row(
    category: str,
    accepted: float = 0.0,
    tentative: float = 0.0,
    declined: float = 0.0,
) -> CategoryRow:
    return CategoryRow(
        category=category,
        accepted_hours=accepted,
        tentative_hours=tentative,
        declined_hours=declined,
        total_hours=round(accepted + tentative + declined, 2),
    )
