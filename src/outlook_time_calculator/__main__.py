from __future__ import annotations

from outlook_time_calculator.cli import parse_args


def main() -> None:
    args = parse_args()

    try:
        from outlook_time_calculator.outlook import fetch_meetings
    except Exception as exc:
        print(f"Error: could not connect to Outlook. Is Outlook Desktop running?\n  {exc}")
        raise SystemExit(1)

    try:
        meetings = fetch_meetings(args.start, args.end)
    except Exception as exc:
        print(f"Error: failed to retrieve calendar items from Outlook.\n  {exc}")
        raise SystemExit(1)

    from outlook_time_calculator.aggregator import aggregate
    from outlook_time_calculator.reporters.csv_reporter import write_csv
    from outlook_time_calculator.reporters.html_reporter import write_html

    rows = aggregate(meetings)

    if not rows:
        print("No categorised meetings found in the specified date range.")
        raise SystemExit(0)

    csv_path = write_csv(rows, args.output_dir, args.start, args.end)
    html_path = write_html(rows, args.output_dir, args.start, args.end)

    print(f"CSV:  {csv_path}")
    print(f"HTML: {html_path}")


if __name__ == "__main__":
    main()
