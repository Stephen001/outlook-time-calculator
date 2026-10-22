from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

from outlook_time_calculator.models import CategoryRow

_COLUMNS = [
    "Category",
    "Accepted Hours",
    "Tentative Hours",
    "Declined Hours",
    "Total Hours",
]


def write_csv(
    rows: list[CategoryRow],
    output_dir: Path,
    start: date,
    end: date,
) -> Path:
    """Write aggregated category rows to a CSV file and return the output path."""
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = f"meeting_report_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.csv"
    output_path = output_dir / filename

    with output_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "Category": row.category,
                    "Accepted Hours": row.accepted_hours,
                    "Tentative Hours": row.tentative_hours,
                    "Declined Hours": row.declined_hours,
                    "Total Hours": row.total_hours,
                }
            )

    return output_path
