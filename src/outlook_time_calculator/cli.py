from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Invalid date '{value}'. Expected format: YYYY-MM-DD"
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="otc",
        description=(
            "Calculate Outlook meeting hours per category for a given date range."
        ),
    )
    parser.add_argument(
        "--start",
        required=True,
        type=_parse_date,
        metavar="YYYY-MM-DD",
        help="Start date (inclusive)",
    )
    parser.add_argument(
        "--end",
        required=True,
        type=_parse_date,
        metavar="YYYY-MM-DD",
        help="End date (inclusive)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path.cwd(),
        metavar="DIR",
        help="Directory to write output files (default: current working directory)",
    )
    return parser.parse_args()
