from __future__ import annotations

from datetime import date
from outlook_time_calculator.models import Meeting
import ctypes
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

# LOCALE_SSHORTDATE — Windows locale constant for the short date format string
_LOCALE_SSHORTDATE = 0x1F


def _outlook_date_format() -> str:
    """Return a strftime format string matching the Windows short date locale.

    Outlook's Restrict filter requires dates in the system's short date format
    (e.g. '%d/%m/%Y' on en-GB, '%m/%d/%Y' on en-US). ISO 8601 is not accepted.
    """
    lcid = ctypes.windll.kernel32.GetUserDefaultLCID()
    buf = ctypes.create_unicode_buffer(80)
    ctypes.windll.kernel32.GetLocaleInfoW(lcid, _LOCALE_SSHORTDATE, buf, 80)
    windows_fmt = buf.value  # e.g. "dd/MM/yyyy"
    # Convert Windows picture format to strftime directives.
    # Replacements use a temporary sentinel to prevent double-substitution
    # (e.g. "dd" → "%d" must not then have its "d" replaced again).
    result = windows_fmt
    result = result.replace("yyyy", "\x00Y")
    result = result.replace("yy",   "\x00y")
    result = result.replace("MMMM", "\x00B")
    result = result.replace("MMM",  "\x00b")
    result = result.replace("MM",   "\x00m")
    result = result.replace("dd",   "\x00d")
    return result.replace("\x00", "%")


def fetch_meetings(start: date, end: date) -> list[Meeting]:
    """Return categorised calendar items in [start, end] via Outlook COM."""
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(_OL_FOLDER_CALENDAR)

    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    fmt = _outlook_date_format()
    start_str = f"{start.strftime(fmt)} 00:00"
    end_str = f"{end.strftime(fmt)} 23:59"
    restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"
    restricted = items.Restrict(restriction)

    meetings: list[Meeting] = []
    for item in restricted:
        try:
            raw_categories: str = item.Categories or ""
        except AttributeError:
            continue

        # For recurring meetings, Outlook may store the category only on the
        # series instead of individual occurrence. Fall back to the
        # series item's categories when the occurrence itself has none.
        if not raw_categories.strip():
            try:
                parent = item.Parent
                raw_categories = getattr(parent, "Categories", "") or ""
            except AttributeError:
                pass

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
