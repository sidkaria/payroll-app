"""Integration smoke test over REAL payroll files in fixtures_real/.

These files contain employee PII and are git-ignored, so this test self-skips
when the directory is empty/absent (CI, fresh clones). When the files ARE
present (local dev) it runs the full pipeline on every site and asserts
PII-free structural invariants — no real names, schedules, or hours appear in
this file.

Convention: each site is a pair `<site>_adp.csv` + `<site>_notes.xlsx` in
fixtures_real/. See fixtures_real/README.md.
"""

from __future__ import annotations

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

FIXTURES = REPO_ROOT / "fixtures_real"


def discover_pairs():
    if not FIXTURES.is_dir():
        return []
    pairs = []
    for adp in sorted(FIXTURES.glob("*_adp.csv")):
        notes = FIXTURES / f"{adp.name[:-len('_adp.csv')]}_notes.xlsx"
        if notes.exists():
            pairs.append((adp, notes))
    return pairs


class RealFixtureInvariantTests(unittest.TestCase):
    def test_pipeline_invariants_on_real_files(self):
        pairs = discover_pairs()
        if not pairs:
            self.skipTest("fixtures_real/ has no <site>_adp.csv/<site>_notes.xlsx pairs")

        for adp_path, notes_path in pairs:
            with self.subTest(site=adp_path.name):
                adp = parse_adp(str(adp_path))
                # Mirror the app: feed the ADP roster so the right notes tab is
                # chosen in multi-tab workbooks.
                notes = parse_notes(
                    str(notes_path),
                    adp_employees={e["employee"] for e in adp["entries"]},
                )
                rules = parse_workbook_rules(notes)
                day_records, logs, leave_totals = apply_notes(adp["entries"], notes, rules)

                self.assertGreater(len(day_records), 0, "expected at least one day record")

                for rec in day_records.values():
                    final = rec["final_work_hours"]
                    # No day is ever negative.
                    self.assertGreaterEqual(
                        final, 0.0, f"negative hours for {rec['employee']} {rec['date']}"
                    )

                    # The cap invariant this fix guards: on a plain capped day
                    # (has a schedule, no approved exception, no missed-punch
                    # correction, no special-day override), paid hours never
                    # exceed min(actual, scheduled). Staying late / short breaks
                    # can't push pay above the scheduled cap.
                    is_plain_cap = (
                        rec["schedule_hours"] is not None
                        and not rec["outside_approved"]
                        and rec["corrected_hours"] is None
                        and rec["special_adjustment"] is None
                    )
                    if is_plain_cap:
                        ceiling = min(rec["actual_hours"], rec["schedule_hours"])
                        self.assertLessEqual(
                            final, ceiling + 1e-6,
                            f"{rec['employee']} {rec['date']}: paid {final} exceeds "
                            f"cap {ceiling}",
                        )


if __name__ == "__main__":
    unittest.main()
