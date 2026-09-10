# Changelog

All notable changes to the IT OPS Sprint Metrics & Review Automation.
Format loosely follows [Keep a Changelog](https://keepachangelog.com/). Dates are when the
work was built during development.

## [0.1.0] — 2026-09-09 — MVP

First stable end-to-end pipeline: **Jira → Excel → PowerPoint**, one command
(`jira_sprint_sync.py`) with four modes (`--dry-run`, `--mode excel|ppt|all`).

### Added — Jira → Excel (`--mode excel`)
- Pull per team from Jira Cloud, scrum and Kanban kept mutually exclusive (like the manual
  "flip Board Type" process):
  - **Scrum** committed / completed velocity / carried-over from the Sprint Report.
  - **Kanban** SPs + count via `"Board Type[Radio Buttons]" = Kanban` + resolutiondate window.
  - **Scrum cross-check** via `Board Type = Scrum` + window, compared to the report (prints OK / MISMATCH).
  - **5 story counts** (committed/added/completed/not-done/removed) from the Sprint Report,
    written to new columns at each tab's last-used + 5 (NET/STN/JST `AN–AR`, WIR `AO–AS`),
    with headers auto-created (idempotent).
  - **NET spikes** → `labels = "Network_Spike"` (any board type) → count (H) + SP sum (I).
  - **STN external impact** → `component in (…6 delay values…)` (both board types),
    distinct-by-key → count (F).
- **Chart-safe writer** (`excel_writer.py`): XML surgery that fills the in-progress sprint row
  and preserves the workbook's 16 hand-formatted charts byte-for-byte (openpyxl's save would
  otherwise gut them). Shifts relative formulas, keeps absolutes (`$S$5`).
- **People/capacity fill policy**: carry Number-of-People & Days-Available; zero Days-Worked;
  blank People-Change & Unplanned-OOO; **Capacity as a live formula** (`=DaysWorked/DaysAvail`,
  shows 0 until Days-Worked entered).
- **Rolling chart windows (A2-deterministic)**: after each fill, repoint that team's chart
  series to the last 6 rows; idempotent, clamped to first data row, other teams untouched.
- **Sparkline removal** helper (charts preserved).
- Per-team config extracted to `team_columns.py`.

### Added — Excel → PowerPoint (`--mode ppt`)
- Regenerate the Sprint Review deck's data slides straight from the workbook, to a **new**
  `..._generated.pptx` (template never overwritten):
  - Title slide: team/sprint label + start/end dates.
  - Metrics slide: identity + counts + health tables, Sprint Metrics chart image.
  - Kanban slide: Kanban table, Kanban chart image.
  - Goals slide: Goals Completed chart image.
- **Chart export**: Excel via AppleScript on macOS (`chart_export.py`); LibreOffice UNO fallback
  on Linux (`lo_chart_export.py`). Charts mapped to slides by **title**, not position
  (export order is metrics/goals/kanban).
- **Fixes the "Stations –" title drift** on every non-NET block automatically.
- Narrative slides (PI Commitment, Risks, Demo, Forecast, Feedback) and manual health rows
  (Impediments?, Goals Removed?, Backlog Readiness) left untouched.
- `workbook_reader.py` reads each team's reported row (+ previous row for "Previous Sprint
  Moving Avg"); `ppt_mode.py` orchestrates reader + export + deck writer.

### Fixed (during development)
- **Timezone skew**: Jira returns UTC but JQL literals are account-local — convert via
  `/myself` timeZone (fixed a 7-hour window shift).
- **Closed-sprint end time**: use `completeDate` (actual close), not scheduled `endDate`, so
  work resolved after close isn't pulled in.
- **Retired Jira endpoint**: `/rest/api/3/search` (410 Gone) → `/rest/api/3/search/jql` with
  `nextPageToken` pagination.
- **Wrong Story-Points field**: auto-detect each board's estimation field instead of a hardcoded
  custom field id (fixed 0-SP Kanban totals and an inflated committed number).
- **Duplicate/renamed sprints**: `--name-contains` lists all matches and picks the most recently
  ended; `--sprint-id` forces a specific one.
- **Header-row detection** after adding inline-string headers (idempotency).
- Deck title bug: no longer overwrites the "Sprint Review" heading; preserves the dates line.

### Known limits / deferred
- **WTE** deferred (Kanban splits into Unplanned/Triage/QRF; different layout).
- **TEL** wired in config; verify its tab columns before first commit.
- macOS: run under Homebrew Python or `uv` (system LibreSSL Python fails TLS to Atlassian).
- Live Jira call + chart export run locally (token authenticates as you; no elevated access).

---

## Unreleased — planned
- WTE multi-bucket Kanban (Excel fill + deck slides).
- TEL activation once its tab has data.
- Optional: fold `ppt_mode` helpers fully into a single package; add automated tests.
