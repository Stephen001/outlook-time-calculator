from __future__ import annotations
from datetime import date
from outlook_time_calculator.models import Meeting
import win32com.client 

# COM ResponseStatus integer → normalised string
# 0 = olResponseNone      → Accepted (treat as accepted/no response needed)
# 1 = olResponseOrganized → Accepted (the user organised it)
# 2 = olResponseTentative → Tentative
# 3 = olResponseAccepted  → Accepted
# 4 = olResponseDeclined  → Declined
# 5 = olResponseNotResponded → Accepted (treat as accepted until declined)
_STATUS_MAP: dict[int, str] = {
    0: "Accepted",
    1: "Accepted",
    2: "Tentative",
    3: "Accepted",
    4: "Declined",
    5: "Accepted",
}

_OL_FOLDER_CALENDAR = 9


def fetch_meetings(start: date, end: date) -> list[Meeting]:
    """Return categorised calendar items in [start, end] via Outlook COM."""
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(_OL_FOLDER_CALENDAR)

    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Outlook Restrict uses MM/DD/YYYY HH:MM format
    start_str = start.strftime("%m/%d/%Y 00:00")
    end_str = end.strftime("%m/%d/%Y 23:59")
    restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"
    restricted = items.Restrict(restriction)

    meetings: list[Meeting] = []
    for item in restricted:
        try:
            raw_categories: str = item.Categories or ""
        except AttributeError:
            continue

        categories = [c.strip() for c in raw_categories.split(",") if c.strip()]
        if not categories:
            continue  # exclude uncategorised meetings

        try:
            response_status = _STATUS_MAP.get(item.ResponseStatus, "Accepted")
            duration_minutes: int = item.Duration
            subject: str = item.Subject or ""
            item_start: date = item.Start.date() if hasattr(item.Start, "date") else item.Start
            item_end: date = item.End.date() if hasattr(item.End, "date") else item.End
        except AttributeError:
            continue

        meetings.append(
            Meeting(
                subject=subject,
                start=item_start,
                end=item_end,
                duration_minutes=duration_minutes,
                categories=categories,
                response_status=response_status,
            )
        )

    return meetings
