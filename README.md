# IT OPS Sprint Metrics & Review Automation

Automates two things that used to be manual, error‑prone, and slow:

1. **Jira → Excel** — pulls each team's sprint numbers from Jira Cloud and fills the
   *IT OPS Teams Sprint Metrics* workbook (velocity, Kanban, spikes, external‑impact, story
   counts), **without ever damaging the workbook's 16 hand‑formatted charts**.
2. **Excel → PowerPoint** — regenerates the data slides of the *Sprint Review* deck straight
   from the workbook (tables, chart images, titles), and **fixes the "Stations –" title drift**
   automatically.

One command drives everything: **`jira_sprint_sync.py`**. The other files are libraries it imports.

---

## Contents

| File | Run it? | What it is |
|---|---|---|
| **`jira_sprint_sync.py`** | ✅ **Yes** — this is your command | Orchestrator: flags, Jira pulls, dispatch to Excel/PPT |
| `excel_writer.py` | ❌ | Chart‑safe XML‑surgery writer (fill row, roll chart windows, add headers, strip sparklines) |
| `team_columns.py` | ❌ (edit when a tab changes) | Per‑team column map + fill policy |
| `workbook_reader.py` | ❌ | Reads a team's reported row for the deck |
| `deck_writer.py` | ❌ | Writes titles/tables/charts into each team's 9‑slide block |
| `chart_export.py` | ❌ | Exports Excel charts → PNG on **macOS** (Excel + AppleScript) |
| `lo_chart_export.py` | ❌ | Exports Excel charts → PNG on **Linux** (LibreOffice, fallback) |
| `ppt_mode.py` | ❌ | `--mode ppt` orchestrator (reader + export + deck writer) |

Keep all files **in the same folder** as the workbook and deck.

---

## One‑time setup

```bash
pip3 install requests openpyxl python-pptx
```

Jira Cloud API token (self‑service at id.atlassian.com → Security → API tokens). Set env vars —
never hardcode:

```bash
export JIRA_SITE="https://allegiantair.atlassian.net"
export JIRA_EMAIL="you@allegiantair.com"
export JIRA_TOKEN="paste-your-token"
```

> `export` lasts only for the current terminal session. Re‑paste in a new terminal, or add the
> three lines to `~/.zshrc` (trade‑off: token then sits in a plaintext file).

**macOS note:** run under a modern Python (Homebrew or `uv`), not the system LibreSSL Python, or
TLS to Atlassian fails. Chart export uses Excel via AppleScript — no extra install.

---

## Daily / weekly workflow

### 1. Preview (always dry‑run first — writes nothing)

```bash
python3 jira_sprint_sync.py --team NET --dry-run
```

Prints the pulled numbers, the exact JQL, the sprint window, and the **scrum cross‑check**
(`OK` or `** MISMATCH`). Default targets the most recent sprint; to validate a specific one:

```bash
python3 jira_sprint_sync.py --team NET --name-contains "Sprint 5" --dry-run
python3 jira_sprint_sync.py --team NET --sprint-id 13493 --dry-run
```

### 2. Fill the workbook (chart‑safe)

```bash
python3 jira_sprint_sync.py --team NET --mode excel      # one team
python3 jira_sprint_sync.py --mode excel                 # all teams
```

Writes a timestamped backup first, fills the in‑progress sprint row, and rolls that team's chart
windows to the last 6 sprints. Then **open the workbook in Excel** (it recalculates on open) and
fill the manual fields (goals, days worked). Add `--no-roll-charts` to skip the chart roll.

### 3. Regenerate the deck

```bash
python3 jira_sprint_sync.py --mode ppt                   # all teams, from the workbook
python3 jira_sprint_sync.py --team NET --mode ppt        # one team
```

Outputs a **new** `..._generated.pptx` (never overwrites your template). Narrative slides
(PI Commitment, Risks, Demo, Forecast, Feedback) are left untouched.

### 4. Everything at once

```bash
python3 jira_sprint_sync.py --team NET --mode all        # Jira → Excel → deck
```

---

## What gets pulled (per team)

| Bucket | Source | Lands in |
|---|---|---|
| **Scrum** — committed, completed velocity, carried‑over | Jira **Sprint Report** (authoritative board report) | C / D / E |
| **Story counts** — committed/added/completed/not‑done/removed | Sprint Report | AN–AR (WIR AO–AS) |
| **Kanban** — SPs + count | `Board Type = Kanban` + resolutiondate window | V/W (NET) · T/U (others) |
| **Scrum cross‑check** | `Board Type = Scrum` + window, compared to report | printed OK/MISMATCH |
| **NET spikes** — count + SP | `labels = "Network_Spike"`, any board type, in window | H / I |
| **STN external impact** — distinct count | `component in (…6 delays…)`, both board types, in window | F |

**Scrum vs. Kanban stay mutually exclusive** by the `Board Type[Radio Buttons]` field — exactly
like flipping the radio button in the manual pull. Dates come from the sprint window (UTC →
account‑local); a **closed** sprint uses `completeDate` (actual close), an **active** sprint uses
the scheduled end. The Story‑Points field is auto‑detected per board.

### Fill policy for the people/capacity block

- **Carried forward** from the last row: Number of People, Days Available
- **Zeroed** (until you edit): Days Worked
- **Manual/blank**: People Change, Unplanned OOO
- **Live formula**: Capacity = `Days Worked / Days Available` (shows 0 until Days Worked is set)

---

## Why the chart‑safe writer exists

openpyxl's normal save **destroys** embedded Excel charts (strips series, data references, and all
color/style parts — leaving empty shells). This workbook has 16 formatted charts. So the writer
never uses openpyxl to save: it edits only the target sheet's XML inside the `.xlsx` zip and copies
every other part **byte‑for‑byte**. Validated: after a fill, only the intended cells change and all
non‑target chart parts are identical; a normal Excel open recalculates formulas with **zero new
errors**.

---

## Current scope & limits

- **Teams automated:** NET, STN, WIR, JST.
- **TEL:** wired in config; verify its tab's columns before the first commit.
- **WTE:** intentionally **deferred** — its Kanban splits into Unplanned/Triage/QRF and its layout
  differs; not yet handled by either mode.
- **Manual deck fields** left untouched: Impediments?, Goals Removed?, Backlog Readiness, and all
  narrative slides.
- The live Jira call must run from your machine/network (the token authenticates **as you**; it
  grants no extra access). Chart export runs locally too.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `SSLError` / `LibreSSL` on macOS | System Python too old — run under Homebrew Python or `uv`. |
| `410 Gone` on search | Old Jira endpoint — this build already uses `/rest/api/3/search/jql`. |
| Kanban count found but 0 SP | Wrong Story‑Points field — auto‑detected per board; override `FALLBACK_SP_FIELD` if needed. |
| Cross‑check `MISMATCH` | Usually a ticket tagged with the wrong Board Type; check the printed JQL in Jira. |
| Wrong sprint pulled | Multiple same‑named sprints — the tool lists matches; use `--sprint-id` to force one. |
| Titles still say "Stations" | Re‑run `--mode ppt`; the drift fix writes each block's correct team name. |

---

## Quick reference

```bash
# preview
python3 jira_sprint_sync.py --team NET --dry-run
# fill excel (all teams) + roll charts
python3 jira_sprint_sync.py --mode excel
# regenerate deck
python3 jira_sprint_sync.py --mode ppt
# end-to-end for one team
python3 jira_sprint_sync.py --team NET --mode all
```
