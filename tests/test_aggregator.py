from __future__ import annotations

from outlook_time_calculator.aggregator import aggregate
from conftest import meeting


class TestAggregate:
    def test_empty_input_returns_empty_list(self):
        assert aggregate([]) == []

    def test_single_meeting_single_category(self):
        rows = aggregate([meeting(["Engineering"], 60, "Accepted")])
        assert len(rows) == 1
        assert rows[0].category == "Engineering"
        assert rows[0].accepted_hours == 1.0
        assert rows[0].tentative_hours == 0.0
        assert rows[0].declined_hours == 0.0
        assert rows[0].total_hours == 1.0

    def test_meeting_with_multiple_categories_fans_out(self):
        rows = aggregate([meeting(["A", "B"], 30, "Accepted")])
        categories = {r.category for r in rows}
        assert categories == {"A", "B"}
        for r in rows:
            assert r.accepted_hours == 0.5

    def test_status_buckets_are_separate(self):
        meetings = [
            meeting(["Cat"], 60, "Accepted"),
            meeting(["Cat"], 30, "Tentative"),
            meeting(["Cat"], 15, "Declined"),
        ]
        rows = aggregate(meetings)
        assert len(rows) == 1
        r = rows[0]
        assert r.accepted_hours == 1.0
        assert r.tentative_hours == 0.5
        assert r.declined_hours == 0.25
        assert r.total_hours == 1.75

    def test_rows_are_sorted_alphabetically(self):
        meetings = [
            meeting(["Zebra"], 10),
            meeting(["Alpha"], 10),
            meeting(["Mango"], 10),
        ]
        rows = aggregate(meetings)
        assert [r.category for r in rows] == ["Alpha", "Mango", "Zebra"]

    def test_hours_rounded_to_two_decimal_places(self):
        # 100 minutes / 60 = 1.6666... → 1.67
        rows = aggregate([meeting(["Cat"], 100, "Accepted")])
        assert rows[0].accepted_hours == 1.67

    def test_multiple_meetings_accumulate(self):
        meetings = [
            meeting(["Cat"], 30, "Accepted"),
            meeting(["Cat"], 30, "Accepted"),
        ]
        rows = aggregate(meetings)
        assert rows[0].accepted_hours == 1.0
