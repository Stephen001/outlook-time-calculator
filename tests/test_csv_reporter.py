from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

from outlook_time_calculator.reporters.csv_reporter import write_csv
from conftest import row


class TestCSVReporter:
    def test_filename_uses_date_range(self, tmp_path: Path):
        path = write_csv([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        assert path.name == "meeting_report_20250101_20250131.csv"

    def test_returns_path_inside_output_dir(self, tmp_path: Path):
        path = write_csv([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        assert path.parent == tmp_path

    def test_creates_output_dir_if_missing(self, tmp_path: Path):
        new_dir = tmp_path / "subdir" / "nested"
        write_csv([], new_dir, date(2025, 1, 1), date(2025, 1, 31))
        assert new_dir.is_dir()

    def test_header_row_is_written(self, tmp_path: Path):
        path = write_csv([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        reader = csv.DictReader(path.read_text(encoding="utf-8").splitlines())
        assert reader.fieldnames == [
            "Category",
            "Accepted Hours",
            "Tentative Hours",
            "Declined Hours",
            "Total Hours",
        ]

    def test_data_rows_are_written_correctly(self, tmp_path: Path):
        rows = [row("Engineering", accepted=1.5, tentative=0.5)]
        path = write_csv(rows, tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        reader = list(csv.DictReader(path.read_text(encoding="utf-8").splitlines()))
        assert len(reader) == 1
        assert reader[0]["Category"] == "Engineering"
        assert reader[0]["Accepted Hours"] == "1.5"
        assert reader[0]["Tentative Hours"] == "0.5"
        assert reader[0]["Declined Hours"] == "0.0"
        assert reader[0]["Total Hours"] == "2.0"

    def test_empty_rows_produces_header_only(self, tmp_path: Path):
        path = write_csv([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        lines = [line for line in path.read_text(encoding="utf-8").splitlines() if line]
        assert len(lines) == 1  # header only
