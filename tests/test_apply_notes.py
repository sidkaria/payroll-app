"""Integration tests for the full pipeline: parse_adp + parse_notes + apply_notes.

Each test builds a synthetic ADP CSV + notes XLSX in a temp dir using
fixtures.write_adp_csv / fixtures.write_notes_xlsx (no PII), runs the
pipeline, and asserts on day_records / leave_totals / logs.

These tests guard the user-visible behaviors that have flip-flopped at
least once during development:
  - Hours capped to schedule unless an approved exception exists
  - Approved-numeric exceptions pay exactly that number
  - Free-text exceptions ("Late pickup", "Pay until", split-shift) parse
  - "no break" alongside times skips the break deduction
  - Time off and holiday flow into leave_totals
  - Schedule anomalies are flagged but DO NOT change paid hours
"""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from payroll_calculator import (  # noqa: E402
    apply_notes,
    parse_adp,
    parse_notes,
    parse_workbook_rules,
)

from tests.fixtures import (  # noqa: E402
    AdpEntry,
    NoteRow,
    NotesWorkbook,
    temp_path,
    write_adp_csv,
    write_notes_xlsx,
)


def run_pipeline(entries: list[AdpEntry], notes_rows: list[NoteRow], prompts=None):
    """Build synthetic files, run the pipeline, return the apply_notes result."""
    csv_path = temp_path(".csv")
    xlsx_path = temp_path(".xlsx")
    try:
        write_adp_csv(
            csv_path,
            pay_start="4/6/2026",
            pay_end="4/17/2026",
            entries=entries,
        )
        write_notes_xlsx(
            xlsx_path,
            NotesWorkbook(rows=notes_rows, prompts=prompts or []),
        )
        adp = parse_adp(str(csv_path))
        notes = parse_notes(str(xlsx_path))
        rules = parse_workbook_rules(notes)
        day_records, logs, leave_totals = apply_notes(adp["entries"], notes, rules)
        return adp, notes, day_records, logs, leave_totals
    finally:
        for p in (csv_path, xlsx_path):
            try:
                os.unlink(p)
            except FileNotFoundError:
                pass


def find_record(day_records, employee, date_str):
    for rec in day_records.values():
        if rec["employee"] == employee and rec["date"].strftime("%-m/%-d/%Y") == date_str:
            return rec
    return None


class CapToScheduleTests(unittest.TestCase):
    def test_overage_is_capped_to_schedule(self):
        # Schedule 9-6 with 1h break = 8h. Worked 9-7 = 10h actual.
        # Expected paid = 8h (cap), with anomaly flagged.
        entries = [
            AdpEntry("Doe, Jane", "4/7/2026", "9:00 AM", "7:00 PM", 10.0),
        ]
        rows = [NoteRow(name="Doe, Jane", schedule="9:00 am - 6:00 pm", break_time="1 hour break")]
        _, _, day_records, logs, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/7/2026")
        self.assertIsNotNone(rec)
        self.assertAlmostEqual(rec["final_work_hours"], 8.0)
        # Anomaly should be present but not change the paid amount
        anomaly_emps = {a["employee"] for a in logs.get("anomalies", [])}
        self.assertIn("Doe, Jane", anomaly_emps)

    def test_under_schedule_pays_actual(self):
        # Worked 7h on an 8h schedule → pay 7h (no upward fill from cap)
        entries = [
            AdpEntry("Doe, Jane", "4/7/2026", "9:00 AM", "5:00 PM", 7.0),
        ]
        rows = [NoteRow(name="Doe, Jane", schedule="9:00 am - 6:00 pm", break_time="1 hour break")]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/7/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.0)


class CapToScheduledWindowTests(unittest.TestCase):
    """Megha's request (2026-06-15): a late arrival made up by staying late or
    taking a shorter break must NOT yield the full scheduled hours. Pay is
    clamped to the overlap of actual punches with the scheduled window, minus
    the FULL scheduled break, with a 5-minute grace on each edge."""

    SCHED = "8:00 am - 5:00 pm"   # 9h span − 1h break = 8h paid

    def _row(self):
        return NoteRow(name="Doe, Jane", schedule=self.SCHED, break_time="1 hour break")

    def test_late_arrival_not_recovered_by_staying_late(self):
        # Megha's canonical case. In 8:15, out 5:15, single punch = 9h actual.
        # Window 8:15→5:00 = 8.75h, − 1h break = 7.75h (NOT 8.0).
        entries = [AdpEntry("Doe, Jane", "4/7/2026", "8:15 AM", "5:15 PM", 9.0)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/7/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.75)

    def test_short_break_with_late_start_is_capped(self):
        # In 8:30 (30 min late), 30-min clocked break, out 5:00 → 8.0h actual.
        # Window 8:30→5:00 = 8.5h − 1h FULL break = 7.5h (short break ignored).
        entries = [
            AdpEntry("Doe, Jane", "4/8/2026", "8:30 AM", "12:00 PM", 3.5),
            AdpEntry("Doe, Jane", "4/8/2026", "12:30 PM", "5:00 PM", 4.5),
        ]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.5)

    def test_early_departure_is_docked(self):
        # In 8:00, out 4:30 (30 min early), single punch = 8.5h actual.
        # Window 8:00→4:30 = 8.5h − 1h break = 7.5h.
        entries = [AdpEntry("Doe, Jane", "4/9/2026", "8:00 AM", "4:30 PM", 8.5)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/9/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.5)

    def test_lateness_is_docked_to_the_minute(self):
        # In 8:06 (6 min late), out 5:00, single punch = 8.9h.
        # Window 8:06→5:00 = 8.9h − 1h break = 7.9h.
        entries = [AdpEntry("Doe, Jane", "4/10/2026", "8:06 AM", "5:00 PM", 8.9)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/10/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.9)

    def test_small_lateness_is_docked_no_grace(self):
        # No grace (Megha, 2026-06-15): even 4 min late is docked to the minute.
        # In 8:04, out 5:04 (stay-late clamped to 5:00). Window 8:04→5:00 =
        # 8.933h − 1h break = 7.93h.
        entries = [AdpEntry("Doe, Jane", "4/13/2026", "8:04 AM", "5:04 PM", 9.0)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/13/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 7.93)

    def test_partial_day_keeps_actual_no_phantom_break(self):
        # Worked an afternoon only: in 2:00 PM, out 5:00 PM = 3h actual, which
        # is well under the 8h schedule. Pay actual 3h — must NOT deduct a full
        # break the employee never took (would underpay to 2h).
        entries = [AdpEntry("Doe, Jane", "4/14/2026", "2:00 PM", "5:00 PM", 3.0)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/14/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 3.0)

    def test_perfect_attendance_still_pays_full_schedule(self):
        # On time, leaves on time, no clocked break (single punch) = 9h actual.
        # Window 8:00→5:00 = 9h − 1h break = 8.0h. No reduction for a clean day.
        entries = [AdpEntry("Doe, Jane", "4/15/2026", "8:00 AM", "5:00 PM", 9.0)]
        _, _, day_records, _, _ = run_pipeline(entries, [self._row()])
        rec = find_record(day_records, "Doe, Jane", "4/15/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 8.0)


class OverrideBeatsWindowCapTests(unittest.TestCase):
    """Megha's follow-up (2026-06-15): the scheduled-window cap must still be
    overridable. When she records an approved exception (pay-actual, late
    pickup, schedule change, split shift, or an explicit number), it is honored
    in FULL — even above the cap — exactly as before. The window cap only
    applies to days with NO recorded exception."""

    SCHED = "8:00 am - 5:00 pm"   # 9h span − 1h break = 8h paid

    def _row(self, **kw):
        return NoteRow(name="Doe, Jane", schedule=self.SCHED,
                       break_time="1 hour break", **kw)

    def test_pay_hours_worked_override_beats_window_cap(self):
        # Late arrival 8:30 + stayed to 6:00 = 9.5h raw. The cap alone would dock
        # this to 7.5h, but an approved "pay hours worked" pays the full 9.5h.
        entries = [AdpEntry("Doe, Jane", "4/8/2026", "8:30 AM", "6:00 PM", 9.5)]
        rows = [self._row(outside_date="4/8/2026", outside_desc="Pay hours worked")]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 9.5)

    def test_late_pickup_override_beats_window_cap(self):
        # On time, stayed to 6:30 for an approved late pickup. The cap would clamp
        # the late stay back to 5:00 (→8.0); the override extends end to 6:30.
        entries = [AdpEntry("Doe, Jane", "4/8/2026", "8:00 AM", "6:30 PM", 10.5)]
        rows = [self._row(outside_date="4/8/2026", outside_desc="Late pickup 6:30 PM")]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 9.5)  # 8:00–6:30 − 1h break

    def test_numeric_approved_hours_override_beats_window_cap(self):
        # Late arrival 8:30 would be docked to 7.5 by the cap, but Megha wrote an
        # explicit approved 9.0 in the Hours cell → pay exactly 9.0.
        entries = [AdpEntry("Doe, Jane", "4/8/2026", "8:30 AM", "6:00 PM", 9.5)]
        rows = [self._row(outside_date="4/8/2026", outside_desc="9")]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 9.0)

    def test_no_break_override_skips_break_over_cap(self):
        # "Worked without a break" (director approved): pay the full window with
        # NO break deduction, even though the standard rule would deduct 1h.
        entries = [AdpEntry("Doe, Jane", "4/8/2026", "8:00 AM", "5:00 PM", 9.0)]
        rows = [self._row(outside_date="4/8/2026",
                          outside_desc="8:00 AM - 5:00 PM no break")]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 9.0)


class ExceptionPassThroughTests(unittest.TestCase):
    def test_late_pickup_extends_end_time(self):
        entries = [
            AdpEntry("Doe, Jane", "4/8/2026", "9:00 AM", "6:30 PM", 9.5),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                outside_date="4/8/2026",
                outside_desc="Late pickup 6:30 PM",
            )
        ]
        _, _, day_records, logs, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        # 9 to 6:30 minus 1h break = 8.5h
        self.assertAlmostEqual(rec["final_work_hours"], 8.5)
        self.assertTrue(any(x["employee"] == "Doe, Jane" for x in logs.get("exceptions", [])))

    def test_no_break_with_time_range_skips_break(self):
        # The user-requested feature: "no break" alongside a time range
        # should skip break deduction.
        entries = [
            AdpEntry("Doe, Jane", "4/8/2026", "7:00 AM", "3:00 PM", 8.0),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="7:00 am - 4:00 pm",
                break_time="1 hour break",
                outside_date="4/8/2026",
                outside_desc="Schedule change 7:00 AM | 3:00 PM No break",
            )
        ]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        # 7-3 = 8h, no break → 8h (vs. 7h with the standard break deduction)
        self.assertAlmostEqual(rec["final_work_hours"], 8.0)

    def test_no_break_with_late_pickup(self):
        entries = [
            AdpEntry("Doe, Jane", "4/8/2026", "9:00 AM", "6:30 PM", 9.5),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                outside_date="4/8/2026",
                outside_desc="Pay until 6:30 PM no break",
            )
        ]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        # 9 AM to 6:30 PM = 9.5h, no break deducted
        self.assertAlmostEqual(rec["final_work_hours"], 9.5)

    def test_pay_hours_worked_uses_actual(self):
        # "Pay hours worked" should pay raw ADP hours, no cap, no break adjustment
        entries = [
            AdpEntry("Doe, Jane", "4/8/2026", "9:00 AM", "8:00 PM", 11.0),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                outside_date="4/8/2026",
                outside_desc="Pay hours worked",
            )
        ]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 11.0)

    def test_split_shift_four_times(self):
        # Worked 8:30-1:00 (4.5h) + 2:00-6:00 (4h) = 8.5h, gap unpaid
        entries = [
            # ADP rolls all punches into a single line in our synthetic; the
            # exception note is the source of truth for split shifts.
            AdpEntry("Doe, Jane", "4/8/2026", "8:30 AM", "6:00 PM", 9.5),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                outside_date="4/8/2026",
                outside_desc="8:30 AM | 1:00 PM | 2:00 PM | 6:00 PM",
            )
        ]
        _, _, day_records, _, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Doe, Jane", "4/8/2026")
        self.assertAlmostEqual(rec["final_work_hours"], 8.5)


class LeaveAllocationTests(unittest.TestCase):
    def test_lump_time_off_added_to_total(self):
        entries = [
            AdpEntry("Doe, Jane", "4/7/2026", "9:00 AM", "6:00 PM", 8.0),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                time_off_hours="8",
            )
        ]
        _, _, _, _, leave_totals = run_pipeline(entries, rows)
        self.assertAlmostEqual(leave_totals["Doe, Jane"]["Time Off"], 8.0)

    def test_holiday_with_date_added(self):
        entries = [
            AdpEntry("Doe, Jane", "4/7/2026", "9:00 AM", "6:00 PM", 8.0),
        ]
        rows = [
            NoteRow(
                name="Doe, Jane",
                schedule="9:00 am - 6:00 pm",
                break_time="1 hour break",
                holiday_date="4/6/2026",
                holiday_hours="8",
            )
        ]
        _, _, _, _, leave_totals = run_pipeline(entries, rows)
        self.assertAlmostEqual(leave_totals["Doe, Jane"]["Holiday"], 8.0)


class NoScheduleEmployeeTests(unittest.TestCase):
    def test_employee_with_punches_but_no_schedule_is_flagged(self):
        # Employee in ADP but absent from notes → shouldn't be capped, and
        # should appear as a review item so Megha can add them to notes.
        entries = [
            AdpEntry("Stranger, Sam", "4/7/2026", "9:00 AM", "6:00 PM", 9.0),
        ]
        rows = [NoteRow(name="Doe, Jane", schedule="9:00 am - 6:00 pm", break_time="1 hour break")]
        _, _, day_records, logs, _ = run_pipeline(entries, rows)
        rec = find_record(day_records, "Stranger, Sam", "4/7/2026")
        self.assertIsNotNone(rec)
        self.assertAlmostEqual(rec["final_work_hours"], 9.0)  # not capped
        review_emps = {item.get("employee") for item in logs.get("review", [])}
        self.assertIn("Stranger, Sam", review_emps)


if __name__ == "__main__":
    unittest.main()
