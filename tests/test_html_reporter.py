from __future__ import annotations

from datetime import date
from pathlib import Path

from outlook_time_calculator.reporters.html_reporter import write_html
from conftest import row


class TestHTMLReporter:
    def test_filename_uses_date_range(self, tmp_path: Path):
        path = write_html([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        assert path.name == "meeting_report_20250101_20250131.html"

    def test_returns_path_inside_output_dir(self, tmp_path: Path):
        path = write_html([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        assert path.parent == tmp_path

    def test_creates_output_dir_if_missing(self, tmp_path: Path):
        new_dir = tmp_path / "sub" / "nested"
        write_html([], new_dir, date(2025, 1, 1), date(2025, 1, 31))
        assert new_dir.is_dir()

    def test_output_is_valid_html_document(self, tmp_path: Path):
        path = write_html([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        assert content.startswith("<!DOCTYPE html>")
        assert "</html>" in content

    def test_date_range_appears_in_title_and_subtitle(self, tmp_path: Path):
        path = write_html([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        assert "01 Jan 2025" in content
        assert "31 Jan 2025" in content

    def test_category_data_appears_in_table(self, tmp_path: Path):
        rows = [row("Engineering", accepted=2.0, declined=0.5)]
        path = write_html(rows, tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        assert "Engineering" in content
        assert "2.0" in content
        assert "0.5" in content

    def test_grand_total_is_correct(self, tmp_path: Path):
        rows = [
            row("A", accepted=1.0),
            row("B", accepted=2.5),
        ]
        path = write_html(rows, tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        assert "3.5" in content

    def test_html_escaping_in_category_name(self, tmp_path: Path):
        rows = [row("<Script>", accepted=1.0)]
        path = write_html(rows, tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        assert "<Script>" not in content
        assert "&lt;Script&gt;" in content

    def test_column_headers_present(self, tmp_path: Path):
        path = write_html([], tmp_path, date(2025, 1, 1), date(2025, 1, 31))
        content = path.read_text(encoding="utf-8")
        for header in ("Category", "Accepted Hours", "Tentative Hours", "Declined Hours", "Total Hours"):
            assert header in content
