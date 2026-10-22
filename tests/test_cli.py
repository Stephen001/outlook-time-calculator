from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

from outlook_time_calculator.cli import parse_args


def _parse(args: list[str]):
    sys.argv = ["otc", *args]
    return parse_args()


class TestCLI:
    def test_valid_dates_are_parsed(self):
        ns = _parse(["--start", "2025-01-01", "--end", "2025-01-31"])
        assert ns.start == date(2025, 1, 1)
        assert ns.end == date(2025, 1, 31)

    def test_output_dir_defaults_to_cwd(self):
        ns = _parse(["--start", "2025-01-01", "--end", "2025-01-31"])
        assert ns.output_dir == Path.cwd()

    def test_output_dir_is_set_when_provided(self, tmp_path: Path):
        ns = _parse(
            ["--start", "2025-01-01", "--end", "2025-01-31", "--output-dir", str(tmp_path)]
        )
        assert ns.output_dir == tmp_path

    def test_invalid_start_date_exits(self):
        with pytest.raises(SystemExit):
            _parse(["--start", "not-a-date", "--end", "2025-01-31"])

    def test_invalid_end_date_exits(self):
        with pytest.raises(SystemExit):
            _parse(["--start", "2025-01-01", "--end", "31-01-2025"])

    def test_missing_start_exits(self):
        with pytest.raises(SystemExit):
            _parse(["--end", "2025-01-31"])

    def test_missing_end_exits(self):
        with pytest.raises(SystemExit):
            _parse(["--start", "2025-01-01"])
