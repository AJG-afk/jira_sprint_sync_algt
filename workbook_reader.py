#!/usr/bin/env python3
"""
workbook_reader.py — read a team's reported sprint row for the deck's data slides.

Target row = most recent row with a numeric Velocity (col D). Previous row = the row above
(for "Previous Sprint Moving Avg"). Read columns differ per team (moving-avg columns esp.).
Manual-only deck fields (Impediments?, Goals Removed?, Backlog Readiness, dates) are NOT sourced.
"""

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as _ci

READ_COLS = {
    "NET": {"sprint": "B", "committed": "C", "completed": "D", "carried": "E",
            "spr_vel_avg": "U", "kanban_pts": "V", "kanban_cnt": "W", "kanban_avg": "X",
            "capacity": "AD",
            "s_committed": "AN", "s_added": "AO", "s_completed": "AP",
            "s_notdone": "AQ", "s_removed": "AR"},
    "STN": {"sprint": "B", "committed": "C", "completed": "D", "carried": "E",
            "spr_vel_avg": "S", "kanban_pts": "T", "kanban_cnt": "U", "kanban_avg": "V",
            "capacity": "AB",
            "s_committed": "AN", "s_added": "AO", "s_completed": "AP",
            "s_notdone": "AQ", "s_removed": "AR"},
    "WIR": {"sprint": "B", "committed": "C", "completed": "D", "carried": "E",
            "spr_vel_avg": "S", "kanban_pts": "T", "kanban_cnt": "U", "kanban_avg": "W",
            "capacity": "AC",
            "s_committed": "AO", "s_added": "AP", "s_completed": "AQ",
            "s_notdone": "AR", "s_removed": "AS"},
    "JST": {"sprint": "B", "committed": "C", "completed": "D", "carried": "E",
            "spr_vel_avg": "S", "kanban_pts": "T", "kanban_cnt": "U", "kanban_avg": "V",
            "capacity": "AB",
            "s_committed": "AN", "s_added": "AO", "s_completed": "AP",
            "s_notdone": "AQ", "s_removed": "AR"},
    "TEL": {"sprint": "B", "committed": "C", "completed": "D", "carried": "E",
            "spr_vel_avg": "S", "kanban_pts": "T", "kanban_cnt": "U", "kanban_avg": "V",
            "capacity": "AB",
            "s_committed": "AN", "s_added": "AO", "s_completed": "AP",
            "s_notdone": "AQ", "s_removed": "AR"},
}


def _num(v, pct=False, ndigits=1):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if pct:
            return f"{round(v * 100)}%"
        return round(v, ndigits) if isinstance(v, float) else v
    return v


def _find_target_rows(ws, header_row):
    vel_rows = [r for r in range(header_row + 1, header_row + 120)
                if isinstance(ws.cell(row=r, column=4).value, (int, float))]
    if not vel_rows:
        return None, None
    target = vel_rows[-1]
    prev = vel_rows[-2] if len(vel_rows) >= 2 else None
    return target, prev


def read_team(xlsx_path, team):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[team]
    cols = READ_COLS[team]
    header = next((r for r in range(1, 60) if ws.cell(row=r, column=2).value == "Sprint"), None)
    if header is None:
        raise RuntimeError(f"{team}: header row not found")
    target, prev = _find_target_rows(ws, header)
    if target is None:
        raise RuntimeError(f"{team}: no completed sprint row found")

    def cell(row, key):
        return ws.cell(row=row, column=_ci(cols[key])).value

    return {
        "_target_row": target, "_prev_row": prev,
        "sprint": cell(target, "sprint"),
        "committed": _num(cell(target, "committed")),
        "completed": _num(cell(target, "completed")),
        "s_committed": _num(cell(target, "s_committed")),
        "s_added": _num(cell(target, "s_added")),
        "s_completed": _num(cell(target, "s_completed")),
        "s_notdone": _num(cell(target, "s_notdone")),
        "s_removed": _num(cell(target, "s_removed")),
        "spr_vel_avg": _num(cell(target, "spr_vel_avg")),
        "capacity": _num(cell(target, "capacity"), pct=True),
        "kanban_pts": _num(cell(target, "kanban_pts")),
        "kanban_cnt": _num(cell(target, "kanban_cnt")),
        "kanban_avg": _num(cell(target, "kanban_avg")),
        "kanban_avg_prev": _num(cell(prev, "kanban_avg")) if prev else None,
        "spr_vel_avg_prev": _num(cell(prev, "spr_vel_avg")) if prev else None,
    }
