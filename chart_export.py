#!/usr/bin/env python3
"""
chart_export.py — export each Excel chart to a standalone PNG.

Backends:
  * macOS (production): drives Excel via AppleScript (osascript) — highest fidelity, no install.
  * Linux/sandbox: delegates to lo_chart_export (LibreOffice UNO).

Public API:
  export_team_charts(xlsx, sheet_name, out_dir) -> {chart_index: png_path}
  chart_index is 0-based export order (metrics, goals, kanban by title — map via deck_writer).
"""

import os, subprocess, platform, zipfile, re

MACOS_EXPORT_SCRIPT = r'''
on run argv
    set wbPath to POSIX file (item 1 of argv) as text
    set outDir to item 2 of argv
    set sheetName to item 3 of argv
    tell application "Microsoft Excel"
        set wb to open workbook workbook file name wbPath
        set ws to worksheet sheetName of wb
        set n to count of chart objects of ws
        repeat with i from 1 to n
            set co to chart object i of ws
            set outFile to outDir & "/" & sheetName & "_chart" & (i - 1) & ".png"
            save co picture in outFile as PNG file format
        end repeat
        close wb saving no
    end tell
end run
'''


def _charts_on_sheet(xlsx, sheet_name):
    with zipfile.ZipFile(xlsx) as z:
        wb = z.read('xl/workbook.xml').decode()
        names = re.findall(r'<sheet name="([^"]*)"', wb)
        if sheet_name not in names:
            return 0
        sheet_no = names.index(sheet_name) + 1
        rels_name = f'xl/worksheets/_rels/sheet{sheet_no}.xml.rels'
        if rels_name not in z.namelist():
            return 0
        rels = z.read(rels_name).decode()
        m = re.search(r'drawing(\d+)\.xml', rels)
        if not m:
            return 0
        draw = z.read(f'xl/drawings/drawing{m.group(1)}.xml').decode()
        return len(re.findall(r'r:id="rId', draw))


def export_macos(xlsx, sheet_name, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    script_path = os.path.join(out_dir, "_export.applescript")
    with open(script_path, "w") as f:
        f.write(MACOS_EXPORT_SCRIPT)
    subprocess.run(["osascript", script_path, os.path.abspath(xlsx),
                    os.path.abspath(out_dir), sheet_name], check=True)
    out = {}
    for i in range(_charts_on_sheet(xlsx, sheet_name)):
        p = os.path.join(out_dir, f"{sheet_name}_chart{i}.png")
        if os.path.exists(p):
            out[i] = p
    return out


def export_team_charts(xlsx, sheet_name, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    if platform.system() == "Darwin":
        return export_macos(xlsx, sheet_name, out_dir)
    import lo_chart_export
    return lo_chart_export.export_sheet_charts(xlsx, sheet_name, out_dir)
