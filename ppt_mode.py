#!/usr/bin/env python3
"""
ppt_mode.py — the `--mode ppt` entry point for the Sprint Review deck generator.

Ties together:
  workbook_reader.read_team    -> per-team sprint values
  chart_export / lo_chart_export -> per-team chart PNGs (Excel on macOS, LibreOffice elsewhere)
  deck_writer.generate_team    -> writes title/tables/charts into each team block

generate_deck(xlsx, template_pptx, out_pptx, teams, dates) -> summary
Outputs a NEW deck (never overwrites the template). Narrative slides untouched; the "Stations –"
title drift is corrected for every team. dates: {team: (title_start, title_end, ident_start, ident_end)}.
If omitted for a team, the existing slide dates are preserved (only the label/team name is fixed).
"""

import os, platform, tempfile, datetime as dt
from pptx import Presentation

import workbook_reader as wr
import deck_writer as dw

DEFAULT_TEAMS = ["NET", "STN", "WIR", "JST"]   # WTE deferred; TEL once its tab has data


def _export_charts(xlsx, team, out_dir):
    if platform.system() == "Darwin":
        import chart_export
        return chart_export.export_team_charts(xlsx, team, out_dir)
    import lo_chart_export
    return lo_chart_export.export_sheet_charts(xlsx, team, out_dir)


def _export_all_charts(xlsx, teams, out_root):
    """Efficient path: on Linux use one LibreOffice session for all teams."""
    if platform.system() == "Darwin":
        return {t: _export_charts(xlsx, t, os.path.join(out_root, t)) for t in teams}
    import lo_chart_export
    return lo_chart_export.export_many(xlsx, teams, out_root)


def fmt_dates(start_local, end_local):
    """('2026-08-19 13:00','2026-09-01 11:47') -> (MM/DD/YYYY, MM/DD/YYYY, 'DD Month YY', ...)."""
    def parse(s):
        return dt.datetime.strptime(str(s).split(" ")[0], "%Y-%m-%d")
    ds, de = parse(start_local), parse(end_local)
    return (ds.strftime("%m/%d/%Y"), de.strftime("%m/%d/%Y"),
            f"{ds.day} {ds.strftime('%B')} {ds.strftime('%y')}",
            f"{de.day} {de.strftime('%B')} {de.strftime('%y')}")


def generate_deck(xlsx, template_pptx, out_pptx, teams=None, dates=None, chart_dir=None):
    teams = teams or DEFAULT_TEAMS
    dates = dates or {}
    chart_dir = chart_dir or tempfile.mkdtemp(prefix="deck_charts_")

    all_charts = _export_all_charts(xlsx, teams, chart_dir)
    prs = Presentation(template_pptx)
    summary = {}
    for team in teams:
        data = wr.read_team(xlsx, team)
        pngs = all_charts.get(team, {})
        team_dates = dates.get(team) or ("", "", "", "")
        res = dw.generate_team(prs, xlsx, team, data, pngs, team_dates)
        summary[team] = {"sprint": data["sprint"], **res}
    prs.save(out_pptx)
    return summary


def run_ppt_mode(xlsx, template_pptx, out_pptx=None, teams=None, gathered=None):
    """Called from jira_sprint_sync.py for --mode ppt/all.
    `gathered` optionally maps team -> (sprint, data) to auto-fill dates from the Jira window."""
    out_pptx = out_pptx or (os.path.splitext(os.path.basename(template_pptx))[0] + "_generated.pptx")
    dates = {}
    if gathered:
        for t, (s, data) in gathered.items():
            w = data.get("_window")
            if w and w[0] and w[1]:
                dates[t] = fmt_dates(w[0], w[1])
    summary = generate_deck(xlsx, template_pptx, out_pptx, teams=teams, dates=dates)
    print(f"[ppt] wrote {out_pptx}")
    for t, info in summary.items():
        print(f"[ppt] {t}: '{info['sprint_label']}' ({info['charts']} charts)")
    return out_pptx
