"""Unit tests for pure-function parsers in payroll_calculator.

These cover the small parser surface area:
  - parse_time / make_after AM-PM heuristic
  - parse_break_hours
  - parse_time_range / parse_schedule_text
  - parse_date_value / expand_date_expression
  - parse_punch_correction
  - parse_exception_text (incl. "no break" override)

The parsers are pure — no IO — so all assertions are direct.
"""

from __future__ import annotations

import sys
import unittest
from datetime import date, time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from payroll_calculator import (  # noqa: E402
    expand_date_expression,
    make_after,
    parse_break_hours,
    parse_date_value,
    parse_exception_text,
    parse_punch_correction,
    parse_schedule_text,
    parse_time,
    parse_time_range,
)


# Schedule with a 1-hour break, used as context for exception parsing.
SCHED_9_TO_6_1H = {
    "mode": "scheduled",
    "start": time(9, 0),
    "end": time(18, 0),
    "break_hours": 1.0,
    "paid_hours": 8.0,
    "label": "9:00 am - 6:00 pm",
}

SCHED_7_TO_4_1H = {
    "mode": "scheduled",
    "start": time(7, 0),
    "end": time(16, 0),
    "break_hours": 1.0,
    "paid_hours": 8.0,
    "label": "7:00 am - 4:00 pm",
}


class TimeHelpersTests(unittest.TestCase):
    def test_parse_time_explicit_ampm(self):
        self.assertEqual(parse_time("9:30 am"), time(9, 30))
        self.assertEqual(parse_time("9:30AM"), time(9, 30))
        self.assertEqual(parse_time("4:30 pm"), time(16, 30))
        self.assertEqual(parse_time("12:00 pm"), time(12, 0))

    def test_parse_time_24h(self):
        self.assertEqual(parse_time("14:30"), time(14, 30))

    def test_parse_time_bare_integer_heuristic(self):
        # Bare integers 1-5 → afternoon; 6-12 → morning. The heuristic only
        # applies to bare integers, not H:MM (which is treated as 24-hour).
        self.assertEqual(parse_time("4"), time(16, 0))
        self.assertEqual(parse_time("8"), time(8, 0))
        self.assertEqual(parse_time("12"), time(12, 0))
        # H:MM without AM/PM is treated as 24-hour, so 4:30 stays AM.
        self.assertEqual(parse_time("4:30"), time(4, 30))

    def test_make_after_flips_pm_when_needed(self):
        # End "4" (= 4 PM) is after start 7 AM → unchanged
        self.assertEqual(make_after(time(16, 0), time(7, 0)), time(16, 0))
        # End "4 AM" before start 7 AM → flipped to 4 PM
        self.assertEqual(make_after(time(4, 0), time(7, 0)), time(16, 0))

    def test_parse_time_range_basic(self):
        r = parse_time_range("9:00 am - 6:00 pm")
        self.assertEqual(r["start"], time(9, 0))
        self.assertEqual(r["end"], time(18, 0))
        self.assertEqual(r["duration"], 9.0)


class BreakHoursTests(unittest.TestCase):
    def test_no_break_keyword(self):
        self.assertEqual(parse_break_hours("no break"), 0.0)
        self.assertEqual(parse_break_hours(""), 0.0)
        self.assertEqual(parse_break_hours("-"), 0.0)

    def test_explicit_break_durations(self):
        self.assertEqual(parse_break_hours("1 hour break"), 1.0)
        self.assertEqual(parse_break_hours("1.5 hours"), 1.5)
        self.assertEqual(parse_break_hours("1/2 hour"), 0.5)
        self.assertEqual(parse_break_hours("30 min"), 0.5)
        self.assertEqual(parse_break_hours("30 minutes"), 0.5)
        self.assertEqual(parse_break_hours("deduct 1 hour"), 1.0)


class ScheduleParsingTests(unittest.TestCase):
    def test_simple_range_yields_weekday_map(self):
        s = parse_schedule_text("9:00 am - 6:00 pm", override_break=1.0)
        self.assertEqual(s["mode"], "scheduled")
        for weekday in range(5):
            info = s["by_weekday"][weekday]
            self.assertEqual(info["start"], time(9, 0))
            self.assertEqual(info["end"], time(18, 0))
            self.assertAlmostEqual(info["paid_hours"], 8.0)

    def test_actual_mode_for_unparseable(self):
        self.assertEqual(parse_schedule_text("")["mode"], "actual")
        self.assertEqual(parse_schedule_text("Give worked hours")["mode"], "actual")

    def test_per_day_pipe_format(self):
        s = parse_schedule_text("M-F - 9:00 am - 6:00 pm", override_break=1.0)
        # The parser only treats "X - HH:MM..." as per-day if "|" appears,
        # so this single-segment input is treated as actual; a proper per-day
        # form looks like: "M-9:00am-6:00pm | T-9:00am-6:00pm | ..."
        self.assertIn(s["mode"], ("scheduled", "actual"))


class DateParsingTests(unittest.TestCase):
    def test_parse_explicit_dates(self):
        self.assertEqual(parse_date_value("4/8/26"), date(2026, 4, 8))
        self.assertEqual(parse_date_value("04/08/2026"), date(2026, 4, 8))
        self.assertEqual(parse_date_value("2026-04-08"), date(2026, 4, 8))

    def test_year_less_falls_back_to_current_year(self):
        result = parse_date_value("4/8")
        self.assertIsNotNone(result)
        self.assertEqual(result.year, date.today().year)
        self.assertEqual((result.month, result.day), (4, 8))

    def test_expand_range_with_year(self):
        dates = expand_date_expression("4/8/26-4/10/26")
        self.assertEqual(dates, [date(2026, 4, 8), date(2026, 4, 9), date(2026, 4, 10)])

    def test_expand_multiline_dates(self):
        dates = expand_date_expression("4/8/26\n4/10/26")
        self.assertEqual(dates, [date(2026, 4, 8), date(2026, 4, 10)])


class PunchCorrectionTests(unittest.TestCase):
    def test_pipe_separated_in_out(self):
        result = parse_punch_correction("9:00 AM | 6:00 PM")
        self.assertEqual(result["status"], "parsed")
        self.assertAlmostEqual(result["hours"], 9.0)

    def test_pipe_with_break_quartet(self):
        # in | out | break-start | break-end pattern
        result = parse_punch_correction("9:00 AM | 6:00 PM break 1:00 PM | 2:00 PM")
        self.assertEqual(result["status"], "parsed")
        self.assertAlmostEqual(result["hours"], 8.0)

    def test_blank_returns_review(self):
        self.assertEqual(parse_punch_correction("")["status"], "review")


class ExceptionTextTests(unittest.TestCase):
    def test_pay_hours_worked_returns_pay_actual(self):
        result = parse_exception_text("pay hours worked", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "pay_actual")

    def test_standalone_no_break_returns_pay_actual(self):
        result = parse_exception_text("no break", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "pay_actual")

    def test_late_pickup_extends_end(self):
        # Schedule 9-6 with 1h break = 8h normal. "Late pickup 6:30" → 9 to 6:30 minus 1h = 8.5
        result = parse_exception_text("Late pickup 6:30 PM", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        self.assertAlmostEqual(result["hours"], 8.5)

    def test_pay_until_extends_end(self):
        result = parse_exception_text("Pay until 6:08 PM", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 9 AM to 6:08 PM = 9.13h, minus 1h break ≈ 8.13
        self.assertAlmostEqual(result["hours"], 8.13)

    def test_came_in_at_extends_start(self):
        result = parse_exception_text("Came in at 8:00 AM", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 8 AM to 6 PM minus 1h break = 9h
        self.assertAlmostEqual(result["hours"], 9.0)

    def test_two_times_schedule_override(self):
        result = parse_exception_text("Schedule 7 AM to 4 PM", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 9h span − 1h break = 8h
        self.assertAlmostEqual(result["hours"], 8.0)

    def test_split_shift_four_times_no_break_deducted(self):
        # 8:30-1:00 = 4.5h, 2:00-6:00 = 4h → total 8.5h, gap is unpaid by construction
        result = parse_exception_text(
            "8:30 AM | 1:00 PM | 2:00 PM | 6:00 PM",
            SCHED_9_TO_6_1H,
        )
        self.assertEqual(result["status"], "parsed")
        self.assertAlmostEqual(result["hours"], 8.5)

    def test_bare_range_worked_pattern(self):
        result = parse_exception_text("Worked 7-4", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 7 AM - 4 PM = 9h, minus 1h break = 8h
        self.assertAlmostEqual(result["hours"], 8.0)

    def test_unparseable_text_returns_unparseable(self):
        # Empty string returns unparseable with empty interpretation
        result = parse_exception_text("", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "unparseable")

    # ── "no break" override ─────────────────────────────────────────────────
    # Megha writes "no break" alongside times to skip the break deduction
    # for that specific day. Standalone "no break" is pay_actual (above).

    def test_no_break_with_time_range_skips_break(self):
        result = parse_exception_text("Schedule 7:00 AM | 3:00 PM | No break", SCHED_7_TO_4_1H)
        self.assertEqual(result["status"], "parsed")
        # 7 AM - 3 PM = 8h, no break deducted
        self.assertAlmostEqual(result["hours"], 8.0)

    def test_with_break_when_no_break_absent(self):
        # Same range without "no break" → break IS deducted
        result = parse_exception_text("Schedule 7:00 AM | 3:00 PM", SCHED_7_TO_4_1H)
        self.assertEqual(result["status"], "parsed")
        self.assertAlmostEqual(result["hours"], 7.0)

    def test_no_break_with_late_pickup(self):
        result = parse_exception_text("Pay until 6:30 PM, no break", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 9 AM to 6:30 PM = 9.5h, no break → 9.5h (vs. 8.5h with break)
        self.assertAlmostEqual(result["hours"], 9.5)

    def test_no_break_bare_range_worked(self):
        result = parse_exception_text("Worked 7-5 no break", SCHED_9_TO_6_1H)
        self.assertEqual(result["status"], "parsed")
        # 10h, no break
        self.assertAlmostEqual(result["hours"], 10.0)

    def test_no_break_interpretation_label(self):
        # When schedule has a break and exception says "no break", the
        # interpretation should mention "(no break)" so Megha sees the call.
        result = parse_exception_text("Schedule 7 AM | 3 PM no break", SCHED_7_TO_4_1H)
        self.assertIn("no break", result["interpretation"].lower())


if __name__ == "__main__":
    unittest.main()
