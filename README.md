# Outlook Time Calculator

Calculates how many hours per week are spent in meetings, broken down by Outlook calendar category and response status (Accepted / Tentative / Declined). Produces a CSV and a self-contained HTML report.

## Requirements

| Requirement | Detail |
|---|---|
| **Operating system** | Windows only |
| **Outlook version** | **Outlook Classic (Win32)** — the desktop application installed via Microsoft 365 or Office LTSC (`OUTLOOK.EXE`). The newer **New Outlook** (`olk.exe`, the Windows Store app) does not expose a COM automation interface and is not supported. |
| **Outlook state** | Outlook Classic must be running and signed in before the tool is executed. |
| **Python** | 3.14 or later |
| **Package manager** | [Poetry](https://python-poetry.org/) |

### Checking which Outlook you have

Open Task Manager while Outlook is running and look at the process name in the **Details** tab:

- `OUTLOOK.EXE` → Outlook Classic ✅ supported
- `olk.exe` → New Outlook ❌ not supported

Alternatively, check whether the following path exists:

```
C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
```

If it does not exist, install the Win32 Office installation. Classic Outlook and New Outlook can coexist on the same machine.

## Installation

```bash
git clone <repo-url>
cd outlook-time-calculator
poetry install
```

## Usage

```
otc --start YYYY-MM-DD --end YYYY-MM-DD [--output-dir DIR]
```

### Arguments

| Argument | Required | Description |
|---|---|---|
| `--start YYYY-MM-DD` | Yes | Start of the date range (inclusive). |
| `--end YYYY-MM-DD` | Yes | End of the date range (inclusive). |
| `--output-dir DIR` | No | Directory to write output files into. Defaults to the current working directory. Created automatically if it does not exist. |

### Examples

Report for a single week, written to the current directory:

```bash
otc --start 2025-06-02 --end 2025-06-06
```

Report for a full month, written to a `reports/` subdirectory:

```bash
otc --start 2025-06-01 --end 2025-06-30 --output-dir ./reports
```

### Output

Two files are written to the output directory, both using the same auto-generated name based on the date range:

| File | Description |
|---|---|
| `meeting_report_YYYYMMDD_YYYYMMDD.csv` | Comma-separated values, one row per category. |
| `meeting_report_YYYYMMDD_YYYYMMDD.html` | Self-contained styled HTML table, suitable for sharing in a browser. |

Both files contain the following columns:

| Column | Description |
|---|---|
| Category | The Outlook calendar category name. |
| Accepted Hours | Total hours of meetings with an Accepted response. |
| Tentative Hours | Total hours of meetings with a Tentative response. |
| Declined Hours | Total hours of meetings with a Declined response. |
| Total Hours | Sum of the three status columns. |

The HTML report also includes a **Grand Total** row at the bottom showing the sum of all category hours.

## Behaviour notes

- **Uncategorised meetings are excluded.** Only meetings that have at least one Outlook category assigned to them appear in the report.
- **Meetings you organised yourself** count as Accepted hours.
- **Meetings with multiple categories** contribute their full duration to each category row independently — hours are not split.
- **Recurring meetings** are expanded; each individual occurrence is counted separately within the date range.
- Hours are **rounded to 2 decimal places**.

## Limitations

- **Windows only.** The tool uses the Windows COM automation interface (`pywin32`) to communicate with the Outlook Desktop process. It cannot run on macOS or Linux.
- **Outlook Classic (Win32) only.** New Outlook (the Windows Store / Progressive Web App version, `olk.exe`) does not register a COM server and cannot be automated by this tool.
- **No Microsoft Graph API.** The tool deliberately avoids cloud authentication. It talks exclusively to the locally running Outlook process.
- **Default calendar only.** The tool reads from the default calendar folder of the signed-in account. Shared calendars, delegate mailboxes, and secondary accounts are not read.
- **System locale affects date filtering.** The Outlook COM `Restrict` filter uses `MM/DD/YYYY` date formatting internally; non-US locale machines may behave differently if the system locale does not match.

## Development

Run the test suite:

```bash
poetry run pytest
```

Tests cover CLI argument parsing, the aggregation logic, the CSV reporter, and the HTML reporter. The Outlook COM layer (`outlook.py`) is not covered by automated tests as it requires a live Outlook Desktop instance.
