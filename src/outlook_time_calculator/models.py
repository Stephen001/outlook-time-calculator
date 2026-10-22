from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass
class Meeting:
    subject: str
    start: date
    end: date
    duration_minutes: int
    categories: list[str]
    response_status: str  # "Accepted" | "Tentative" | "Declined"


@dataclass
class CategoryRow:
    category: str
    accepted_hours: float
    tentative_hours: float
    declined_hours: float
    total_hours: float
