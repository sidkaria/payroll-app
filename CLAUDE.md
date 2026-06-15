# Payroll Calculator — Repo Guide

A Streamlit + Python tool that processes ADP time-clock exports against a
school administrator's ("Megha") payroll-notes workbook and produces a
corrected hours report. Deployed on Streamlit Community Cloud; the user
uploads two files in the browser and downloads the result.

## File layout

```
payroll_app/
├── app.py                  # Streamlit UI: uploads, summary metrics, download
├── payroll_calculator.py   # All parsing + calculation logic (single module)
├── requirements.txt        # streamlit, openpyxl, python-docx
├── tests/                  # unittest suite — see "Tests" below
└── .gitignore              # blocks *.csv/*.xlsx so real payroll never leaks
```

Real ADP exports and notes workbooks contain employee PII. **Never commit
them.** Tests use synthetic fixtures (`tests/fixtures.py`), which build
ADP CSVs and notes XLSX files in temp dirs with placeholder names like
"Doe, Jane".

## Inputs

1. **ADP export** — CSV or XLSX. Row 1 has `Date range, M/D/YY, M/D/YY`.
   Row 2 onward is one punch per row: `Employee Name | Pay # | Date |
   In time | Out time | Hours | …`.
2. **Notes workbook** — XLSX (or legacy DOCX). Single sheet (the parser
   reads `sheetnames[0]`). Row 4 is the header row, row 5 has sub-headers,
   and employee rows start at row 6:

   | Col | Header                          | Notes                                     |
   |-----|---------------------------------|-------------------------------------------|
   | 1   | Employee                        | "Last, First"                             |
   | 2   | Status                          | "Full Time" / "Part Time" / "Contractor"  |
   | 3   | Scheduled hours                 | e.g. `9:00 am - 6:00 pm`                  |
   | 4   | Break Time                      | e.g. `1 hour break`, `1/2 hour`, `30 min` |
   | 5–6 | Missed punches (Date, In/Out)   | one date or many; pipe-separated punches  |
   | 7–8 | Worked Outside Schedule (Date, Description) | free-text — see "Exceptions" |
   | 9–10| Time Off (Date, Hours)          | hours treated as a lump sum               |
   | 11–12| Holiday (Date, Hours)          | date required; hours paid as-is           |
   | 13  | Notes                           | informational                             |

   **Multi-sheet caveat:** the parser reads only the first sheet. If
   Megha keeps an empty `Sample Template` tab in front of the real data
   tab, she must remove the template (or move it after the data tab) so
   the data sheet is the first one.

## Calculation behavior

The behaviors below have flip-flopped at least once during development —
each is now pinned by an integration test in `tests/test_apply_notes.py`.

### Cap to scheduled window (default)

Paid work hours are capped to the **scheduled window**, not just the daily
total. On a day the employee worked at least a full day's hours
(`actual >= scheduled`), pay is the overlap of the actual punches with the
scheduled time range, minus the **full** scheduled break:

```
paid = (min(last_out, sched_end) − max(first_in, sched_start)) − full_break
final = min(actual, paid)          # never above the plain cap
```

A **5-minute grace** is applied on each edge (`LATE_GRACE`), so trivial clock
drift isn't docked, but anything beyond is. The key consequence (Megha,
2026-06-15): a **late arrival or early departure can't be recovered** by
staying late, coming in early, or taking a shorter break than scheduled — and
the full break is always deducted even if a shorter one was taken. The logic
lives in `paid_within_schedule()`.

This only fires on full-day records. On a **genuine short day**
(`actual < scheduled`) we keep paying actual (`min(actual, scheduled)`), so we
never deduct a break that wasn't taken or dock honest partial hours.
**Missed-punch–corrected days** keep the plain `min(actual, scheduled)` cap,
since their raw punches are incomplete by definition. Overage above schedule
is still **not** paid unless an exception is recorded — Megha's earlier
preference (2026-04-21): "Keep the cap. I will fix the approved time manually."

When actual > scheduled + 15 min, an anomaly is logged in
`logs["anomalies"]` and surfaced under "Schedule Anomalies" in the UI,
but it does not change paid hours. Same for clock-in-early and
clock-out-late anomalies (15-min tolerance). These anomaly thresholds are
independent of the 5-min pay grace above.

### Approved exceptions (override the cap)

Three ways to authorize hours over the cap, evaluated in this order:

1. **Numeric `Hours` cell** (column 10 next to outside-schedule date and
   description) — pays exactly that number.

2. **Free-text description** in the description cell — auto-parsed by
   `parse_exception_text` against the schedule for that day:

   | Phrase example                          | Behavior                                       |
   |-----------------------------------------|------------------------------------------------|
   | `pay hours worked` / `pay actual`       | Pay raw ADP hours, no cap, no break deduction  |
   | `no break` (alone)                      | Same as above                                  |
   | `Late pickup 6:30 PM` / `Pay until 6:30`| Extend schedule end to that time, deduct break |
   | `Came in at 7 AM` / `Started 7`         | Extend schedule start to that time, ded. break |
   | `Schedule 7 AM - 4 PM` / `Worked 7-4`   | Override the day's schedule with that range    |
   | `8:30 AM \| 1:00 PM \| 2:00 PM \| 6:00 PM`| Split shift — sum each segment, no break ded. |
   | `... no break` alongside any of above   | Skip the break deduction for this day          |

   "no break" is detected as a substring; standalone `no break` falls
   into the pay-actual path; combined with times it skips the break in
   patterns 2/3/4.

3. **Unparseable description** — paid at schedule (cap applies) and
   added to "Needs Review" so Megha can clean it up.

### Time off and holiday

- **Time Off** is treated as a lump sum (column 10 hours total). Dates
  in column 9 are not required and not used for date-by-date allocation.
  The total accumulates in `leave_totals[employee]["Time Off"]`.
- **Holiday** requires a date in column 11 and hours in column 12. Hours
  are paid as-is — never capped to the schedule.

### Employees in ADP but not in notes

If an employee has punches but isn't in the notes sheet, their hours are
**not** capped (no schedule to cap against), and a "No schedule found"
review item is added so Megha can adjust manually or add them to notes.

## Tests

```bash
cd payroll_app
python3 -m unittest discover tests -v
```

Tests use only stdlib `unittest` plus existing deps (no pytest). Two
files:

- `tests/test_parsers.py` — unit tests for the pure parsers
  (`parse_time`, `parse_break_hours`, `parse_exception_text`,
  `parse_punch_correction`, date parsing, schedule parsing).
- `tests/test_apply_notes.py` — integration tests that build synthetic
  ADP CSVs + notes XLSX files via `tests/fixtures.py` and assert on the
  output of the full pipeline.

When changing calculation logic, run the suite first to confirm green,
then add a new test for the behavior you're introducing or changing.
The fixture builder accepts plain dataclasses (`AdpEntry`, `NoteRow`),
so adding a regression scenario is normally ~10 lines.

**Do not add real names, real schedules from production files, or real
hours to fixtures.** Use generic placeholders ("Doe, Jane", "Stranger,
Sam") and round numbers.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

Streamlit Community Cloud, deployed from the `main` branch of the
GitHub repo `sidkaria/payroll-app`. Pushes to `main` redeploy
automatically; Megha gets the new behavior on her next page load.


## Work tracking

Work for this repo is tracked in `.backlog/` — one markdown file per task
(`status`, `priority`, `depends_on`, `needs_keys`). `ROADMAP.md` is a generated
view (`node tools/task render`); never hand-edit it.

- User shares an idea → invoke the `capture-idea` skill (files it in `.backlog/`).
- User says "implement X" / "what's next" → invoke `focus-task`
  (`node tools/task next` → build → `./check.sh` → commit → push to main).
- `./check.sh` must exit 0 before any push; never edit it to pass.

`.backlog/` + `tools/task` + `check.sh` are the contract any external automation
consumes — the repo doesn't need to know what's driving it.
