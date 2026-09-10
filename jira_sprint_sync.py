#!/usr/bin/env python3
"""
jira_sprint_sync.py — Jira Cloud -> IT OPS Sprint Metrics workbook + Sprint Review deck (chart-safe).

This is the ONLY script you run. It imports the helper modules:
  excel_writer.py, team_columns.py, workbook_reader.py, deck_writer.py,
  chart_export.py / lo_chart_export.py, ppt_mode.py

MODES:
  --dry-run        pull + print everything, write nothing
  --mode excel     fill the in-progress sprint row (chart-safe) + roll chart windows to last 6
  --mode ppt       regenerate the Sprint Review deck's data slides (new *_generated.pptx)
  --mode all       excel then ppt

DATA PULLED per team (scrum vs kanban kept mutually exclusive, like the manual process):
  SCRUM  -> Jira Sprint Report: committed / completed velocity / carried-over + 5 story counts
  KANBAN -> Board Type = Kanban + resolutiondate window: SPs + story count
  X-CHK  -> Board Type = Scrum + resolutiondate window, compared to the report (prints OK/MISMATCH)
  NET spikes (H/I)         -> labels = "Network_Spike", any board type, in window: count + SP sum
  STN external impact (F)  -> component in (6 delay values), both board types, in window: distinct count

Estimation field auto-detected per board. Timezone: API UTC -> account-local for JQL literals.
Sprint end: closed -> completeDate (actual close); active -> scheduled endDate. WTE deferred.

Auth via env: JIRA_SITE, JIRA_EMAIL, JIRA_TOKEN.

Examples:
  python jira_sprint_sync.py --team NET --dry-run
  python jira_sprint_sync.py --team NET --name-contains "Sprint 5" --dry-run
  python jira_sprint_sync.py --team NET --mode excel
  python jira_sprint_sync.py --mode ppt          # regenerate deck from the workbook
  python jira_sprint_sync.py --team NET --mode all
"""

import os, sys, argparse, shutil, datetime as dt
import requests
from requests.auth import HTTPBasicAuth

import excel_writer
from team_columns import TEAM_COLUMNS

BOARD_TYPE_FIELD = '"Board Type[Radio Buttons]"'
STD_FILTER = 'issuetype NOT IN (Initiative, Epic, Sub-task) AND resolution != Cancelled'
FALLBACK_SP_FIELD = "customfield_10037"

CONFIG = {
    "NET": {"board_id": 299,  "project": 'project = "Network Engineering Team"',
            "use_board_type": True, "spike_label": "Network_Spike"},
    "STN": {"board_id": 2774, "project": 'project = STN', "use_board_type": True,
            "delay_components": ["Delay - Facility/Airport", "Delay -  Internal/IT Other",
                                 "Delay -  Internal/IT Stations", "Delay - Internal/Non-IT Related",
                                 "Delay - Other/Misc.", "Delay - Vendor"]},
    "WIR": {"board_id": 897,  "project": 'project = "Windows, Infrastructure Reliability"',
            "use_board_type": True},
    "JST": {"board_id": 6446, "project": 'project = "Jira Support Team"', "use_board_type": True},
    "TEL": {"board_id": 7711, "project": 'project = "Telecommunications Team"', "use_board_type": True},
    # WTE deferred.
}
RUN_SCRUM_CROSSCHECK = True
CROSSCHECK_TOLERANCE = 0.5
DECK_TEMPLATE_DEFAULT = "Sprint Review - IT OPS - 20260902.pptx"

STORY_HEADERS = {
    "stories_committed": "Stories Committed at Start",
    "stories_added":     "Stories Added",
    "stories_completed": "Stories Completed",
    "stories_notdone":   "Stories Not Completed",
    "stories_removed":   "Stories Removed",
}


def auth():
    site = os.environ.get("JIRA_SITE"); email = os.environ.get("JIRA_EMAIL"); tok = os.environ.get("JIRA_TOKEN")
    if not (site and email and tok):
        sys.exit("ERROR: set JIRA_SITE, JIRA_EMAIL, JIRA_TOKEN environment variables first.")
    return site.rstrip("/"), HTTPBasicAuth(email, tok)


def jget(site, a, path, **params):
    r = requests.get(f"{site}{path}", auth=a, params=params, headers={"Accept": "application/json"}, timeout=60)
    r.raise_for_status()
    return r.json()


def account_timezone(site, a):
    try:
        return jget(site, a, "/rest/api/3/myself").get("timeZone")
    except Exception:
        return None


def board_estimation_field(site, a, board_id):
    try:
        cfg = jget(site, a, f"/rest/agile/1.0/board/{board_id}/configuration")
        return ((cfg.get("estimation") or {}).get("field") or {}).get("fieldId")
    except Exception:
        return None


def jira_dt_local(iso, tzname=None):
    if not iso:
        return None
    s = iso.replace("Z", "+0000"); d = None
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"):
        try:
            d = dt.datetime.strptime(s, fmt); break
        except ValueError:
            continue
    if d is None:
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
            try:
                return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M")
            except ValueError:
                continue
        return None
    try:
        from zoneinfo import ZoneInfo
        d = d.astimezone(ZoneInfo(tzname)) if tzname else d.astimezone()
    except Exception:
        d = d.astimezone()
    return d.strftime("%Y-%m-%d %H:%M")


def list_sprints(site, a, board_id):
    out, start = [], 0
    while True:
        j = jget(site, a, f"/rest/agile/1.0/board/{board_id}/sprint",
                 state="active,closed", startAt=start, maxResults=50)
        out += j.get("values", [])
        if j.get("isLast", True):
            break
        start += 50
    return out


def _end_key(s):
    return s.get("completeDate") or s.get("endDate") or ""


def pick_sprint(sprints, sprint_id=None, name_contains=None):
    if sprint_id:
        return next((s for s in sprints if str(s["id"]) == str(sprint_id)), None), []
    if name_contains:
        m = sorted([s for s in sprints if name_contains.lower() in s.get("name", "").lower()],
                   key=_end_key, reverse=True)
        return (m[0] if m else None), m
    dated = sorted([s for s in sprints if _end_key(s)], key=_end_key, reverse=True)
    return (dated[0] if dated else None), dated[:3]


def sprint_window(s, tzname):
    start = jira_dt_local(s.get("startDate"), tzname)
    if s.get("completeDate"):
        return start, jira_dt_local(s["completeDate"], tzname), "completeDate (actual close)"
    return start, jira_dt_local(s.get("endDate"), tzname), "endDate (scheduled — active sprint)"


def _issue_est(iss):
    est = iss.get("estimateStatistic") or {}
    v = (est.get("statFieldValue") or {}).get("value")
    if v is None:
        v = (iss.get("currentEstimateStatFieldValue") or {}).get("value")
    try:
        return float(v) if v is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def sprint_report(site, a, board_id, sprint_id):
    j = jget(site, a, "/rest/greenhopper/1.0/rapid/charts/sprintreport",
             rapidViewId=board_id, sprintId=sprint_id)
    c = j["contents"]
    def val(d): return float(d.get("value") or 0) if isinstance(d, dict) else 0.0
    completed = c.get("completedIssues", [])
    not_done = c.get("issuesNotCompletedInCurrentSprint", [])
    punted = c.get("puntedIssues", [])
    added = set((c.get("issueKeysAddedDuringSprint") or {}).keys())
    completed_sp = val(c.get("completedIssuesEstimateSum", {}))
    notdone_sp = val(c.get("issuesNotCompletedEstimateSum", {}))
    all_sp = val(c.get("allIssuesEstimateSum", {}))
    added_sp = sum(_issue_est(i) for i in (completed + not_done) if i.get("key") in added)
    return {
        "committed_sp": round(all_sp - added_sp, 2),
        "completed_sp": round(completed_sp, 2),
        "carried_over": round(notdone_sp, 2),
        "stories_committed": len(completed) + len(not_done) - len(added),
        "stories_added": len(added),
        "stories_completed": len(completed),
        "stories_notdone": len(not_done),
        "stories_removed": len(punted),
    }


def jql_totals(site, a, jql, sp_field, distinct=False):
    total, keys, count, token = 0.0, set(), 0, None
    while True:
        params = {"jql": jql, "maxResults": 100, "fields": sp_field}
        if token:
            params["nextPageToken"] = token
        j = jget(site, a, "/rest/api/3/search/jql", **params)
        for iss in j.get("issues", []):
            k = iss.get("key")
            if distinct and k in keys:
                continue
            keys.add(k); count += 1
            sp = iss.get("fields", {}).get(sp_field)
            if sp is not None:
                try:
                    total += float(sp)
                except (TypeError, ValueError):
                    pass
        token = j.get("nextPageToken")
        if j.get("isLast") or not token:
            break
    return round(total, 2), (len(keys) if distinct else count)


def build_jql(project, board_type, start_j, end_j, extra=None):
    parts = [project, f'resolutiondate >= "{start_j}"', f'resolutiondate <= "{end_j}"', STD_FILTER]
    if extra:
        parts.append(extra)
    if board_type:
        parts.append(f'{BOARD_TYPE_FIELD} = {board_type}')
    return " AND ".join(parts)


def gather_team(site, a, team, s, tzname, sp_field):
    cfg = CONFIG[team]
    start_j, end_j, end_src = sprint_window(s, tzname)
    out = {"_window": (start_j, end_j), "_end_src": end_src}
    out.update(sprint_report(site, a, cfg["board_id"], s["id"]))
    if start_j and end_j:
        proj = cfg["project"]
        if cfg.get("use_board_type"):
            k_jql = build_jql(proj, "Kanban", start_j, end_j)
            sp, cnt = jql_totals(site, a, k_jql, sp_field)
            out.update({"kanban_sp": sp, "kanban_count": cnt, "_kanban_jql": k_jql})
            if RUN_SCRUM_CROSSCHECK:
                s_jql = build_jql(proj, "Scrum", start_j, end_j)
                xsp, xcnt = jql_totals(site, a, s_jql, sp_field)
                out.update({"_scrum_check_sp": xsp, "_scrum_check_count": xcnt, "_scrum_jql": s_jql})
        if cfg.get("spike_label"):
            sp_jql = build_jql(proj, None, start_j, end_j, extra=f'labels = "{cfg["spike_label"]}"')
            ssp, scnt = jql_totals(site, a, sp_jql, sp_field)
            out.update({"spike_sp": ssp, "spike_count": scnt, "_spike_jql": sp_jql})
        if cfg.get("delay_components"):
            comp = " , ".join(f'"{c}"' for c in cfg["delay_components"])
            d_jql = build_jql(proj, None, start_j, end_j, extra=f'component in ({comp})')
            _dsp, dcnt = jql_totals(site, a, d_jql, sp_field, distinct=True)
            out.update({"external_impact": dcnt, "_delay_jql": d_jql})
    else:
        out["_warn"] = "Missing sprint start/end; skipped date-bounded queries."
    return out


def build_values(team, data):
    tc = TEAM_COLUMNS[team]
    values = {}
    for field, col in tc["jira"].items():
        if field == "sprint":
            if data.get("_sprint_name"):
                values[col] = data["_sprint_name"]
        elif data.get(field) is not None:
            values[col] = data[field]
    for field, col in tc.get("story_counts", {}).items():
        if data.get(field) is not None:
            values[col] = data[field]
    return values


def print_team(team, s, data):
    w0, w1 = data.get("_window", (None, None))
    print(f"\n[{team}] '{s['name']}' (id {s['id']}, state={s.get('state')})")
    print(f"      SP field = {data.get('_sp_field')}")
    print(f"      window  = {w0} -> {w1}  [end from {data.get('_end_src')}]")
    for k in ("committed_sp", "completed_sp", "carried_over",
              "stories_committed", "stories_added", "stories_completed",
              "stories_notdone", "stories_removed",
              "kanban_sp", "kanban_count", "spike_count", "spike_sp", "external_impact"):
        if k in data:
            print(f"      {k:18} = {data[k]}")
    if "_scrum_check_sp" in data:
        gap = round(abs(data["_scrum_check_sp"] - data["completed_sp"]), 2)
        flag = "OK" if gap <= CROSSCHECK_TOLERANCE else f"** MISMATCH (Δ {gap})"
        print(f"      scrum x-check = {data['_scrum_check_sp']} SP / {data.get('_scrum_check_count')} vs "
              f"report {data['completed_sp']} SP / {data['stories_completed']}  [{flag}]")


def ensure_story_headers(workbook, team):
    sc = TEAM_COLUMNS[team].get("story_counts")
    if not sc:
        return []
    header_map = {sc[f]: STORY_HEADERS[f] for f in sc if f in STORY_HEADERS}
    return excel_writer.add_header_cells(workbook, team, header_map)


def do_excel(workbook, team, data, roll_charts=True):
    tc = TEAM_COLUMNS[team]
    ensure_story_headers(workbook, team)
    values = build_values(team, data)
    rn = excel_writer.fill_sprint_row(
        workbook, team, values,
        complete_col=tc["jira"]["completed_sp"],
        carry_cols=tc["carry"], zero_cols=tc["zero"], blank_cols=tc["blank"],
        formula_cols=tc.get("formula"))
    rolled = []
    if roll_charts:
        rolled = excel_writer.update_chart_windows(
            workbook, team, rn, height=6, complete_col=tc["jira"]["completed_sp"])
    return rn, values, rolled


def main():
    ap = argparse.ArgumentParser(description="Sync Jira sprint metrics into the IT OPS workbook + deck.")
    ap.add_argument("--workbook", default="IT OPS Teams Sprint Metrics (Non-AI).xlsx")
    ap.add_argument("--deck-template", default=DECK_TEMPLATE_DEFAULT)
    ap.add_argument("--team")
    ap.add_argument("--sprint-id")
    ap.add_argument("--name-contains")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--mode", choices=["excel", "ppt", "all"], default="excel")
    ap.add_argument("--no-roll-charts", action="store_true", help="Skip rolling chart windows.")
    ap.add_argument("--show-jql", action="store_true")
    args = ap.parse_args()

    # --- PPT-only mode: no Jira calls needed; regenerate deck straight from the workbook ---
    if args.mode == "ppt":
        import ppt_mode
        teams = [args.team] if args.team else None
        ppt_mode.run_ppt_mode(args.workbook, args.deck_template, teams=teams)
        return

    site, a = auth()
    tzname = account_timezone(site, a)
    print(f"Account timezone: {tzname or '(unknown)'}")
    teams = [args.team] if args.team else list(CONFIG.keys())
    gathered = {}

    for t in teams:
        cfg = CONFIG.get(t)
        if not cfg:
            sys.exit(f"Unknown team {t}")
        sp_field = board_estimation_field(site, a, cfg["board_id"]) or FALLBACK_SP_FIELD
        sprints = list_sprints(site, a, cfg["board_id"])
        s, matches = pick_sprint(sprints, args.sprint_id, args.name_contains)
        if not s:
            print(f"[{t}] SKIP — no matching sprint."); continue
        if args.name_contains and len(matches) > 1:
            print(f"[{t}] NOTE: {len(matches)} match '{args.name_contains}'; chose most recent (id {s['id']}). "
                  f"Use --sprint-id to force.")
        data = gather_team(site, a, t, s, tzname, sp_field)
        data["_sp_field"] = sp_field
        data["_sprint_name"] = s["name"]
        gathered[t] = (s, data)
        print_team(t, s, data)
        if args.show_jql or args.dry_run:
            for key in ("_kanban_jql", "_scrum_jql", "_spike_jql", "_delay_jql"):
                if data.get(key):
                    print(f"      {key[1:]}: {data[key]}")
        if data.get("_warn"):
            print(f"      WARNING: {data['_warn']}")

    if not gathered:
        print("\nNothing to do."); return
    if args.dry_run:
        print("\nDRY RUN — nothing written."); return

    if args.mode in ("excel", "all"):
        if not os.path.exists(args.workbook):
            sys.exit(f"Workbook not found: {args.workbook}")
        backup = f"{args.workbook.rsplit('.',1)[0]}_backup_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(args.workbook, backup)
        print(f"\nBackup: {backup}")
        for t, (s, data) in gathered.items():
            rn, vals, rolled = do_excel(args.workbook, t, data, roll_charts=not args.no_roll_charts)
            print(f"[{t}] filled row {rn}: {vals}")
            if rolled:
                print(f"[{t}] rolled {len(rolled)} chart range(s) to last 6 (ending row {rn}).")
        print(f"Saved {args.workbook}. Open in Excel (recalcs on open); fill goals/days.")

    if args.mode == "all":
        import ppt_mode
        ppt_mode.run_ppt_mode(args.workbook, args.deck_template,
                              teams=list(gathered.keys()), gathered=gathered)


if __name__ == "__main__":
    main()
