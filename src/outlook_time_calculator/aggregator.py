from __future__ import annotations

from collections import defaultdict

from outlook_time_calculator.models import CategoryRow, Meeting


def aggregate(meetings: list[Meeting]) -> list[CategoryRow]:
    """Aggregate meeting durations into hours per category and response status."""
    # Accumulate minutes keyed by (category, status)
    buckets: dict[str, dict[str, float]] = defaultdict(
        lambda: {"Accepted": 0.0, "Tentative": 0.0, "Declined": 0.0}
    )

    for meeting in meetings:
        for category in meeting.categories:
            buckets[category][meeting.response_status] += meeting.duration_minutes

    rows: list[CategoryRow] = []
    for category, status_minutes in sorted(buckets.items()):
        accepted_hours = round(status_minutes["Accepted"] / 60, 2)
        tentative_hours = round(status_minutes["Tentative"] / 60, 2)
        declined_hours = round(status_minutes["Declined"] / 60, 2)
        total_hours = round(accepted_hours + tentative_hours + declined_hours, 2)
        rows.append(
            CategoryRow(
                category=category,
                accepted_hours=accepted_hours,
                tentative_hours=tentative_hours,
                declined_hours=declined_hours,
                total_hours=total_hours,
            )
        )

    return rows
