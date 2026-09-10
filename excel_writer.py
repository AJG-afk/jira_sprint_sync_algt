#!/usr/bin/env python3
"""
excel_writer.py — chart-safe row fill + chart-window roll for the IT OPS Sprint Metrics workbook.

WHY: openpyxl's load->save GUTS embedded Excel charts (strips series, data refs, color/style parts).
This workbook has 16 hand-formatted charts that must survive. So we do targeted XML surgery: edit
ONLY the parts we must inside the .xlsx zip and copy every other part byte-for-byte.

Public API:
  fill_sprint_row(xlsx, sheet, values, complete_col='D',
                  carry_cols=(), zero_cols=(), blank_cols=(), formula_cols=None) -> row number
  add_header_cells(xlsx, sheet, header_map, style_from_col='B') -> [cols added]
  update_chart_windows(xlsx, sheet, last_row, height=6) -> [(chartfile, old, new), ...]
  strip_sparklines(xlsx) -> count removed
"""

import re, zipfile, os

CELLREF = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def shift_formula(formula, delta):
    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        newrow = row if row_abs == '$' else str(int(row) + delta)
        return f'{col_abs}{col}{row_abs}{newrow}'
    return CELLREF.sub(repl, formula)


def _read_sheet_map(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    name_rid = re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb)
    rid_tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels))
    return {name: 'xl/' + rid_tgt[rid] for name, rid in name_rid if rid in rid_tgt}


def _cells_of_row(row_xml):
    for c in re.findall(r'<c [^>]*?/>|<c [^>]*?>.*?</c>', row_xml):
        ref = re.search(r'r="([A-Z]+)\d+"', c)
        if ref:
            yield ref.group(1), c


def _style_of(cell_xml):
    m = re.search(r's="(\d+)"', cell_xml)
    return m.group(1) if m else None


def _find_rows(sheet_xml):
    return [(int(m.group(1)), m.group(0))
            for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', sheet_xml)]


def _header_row(rows):
    best_rn, best_score = None, -1
    for rn, rx in rows:
        cells = list(_cells_of_row(rx))
        if len(cells) < 10:
            continue
        text_cells = sum(1 for _, cx in cells if 't="s"' in cx or 't="inlineStr"' in cx)
        if text_cells >= len(cells) - 1 and text_cells > best_score:
            best_rn, best_score = rn, text_cells
    return best_rn


def _row_has_numeric(row_xml, want_col):
    for col, cx in _cells_of_row(row_xml):
        if col == want_col and '<v>' in cx and 't="s"' not in cx:
            return True
    return False


def _last_complete_row(rows, header_row, complete_col='D'):
    best = header_row
    for rn, rx in rows:
        if rn > header_row and _row_has_numeric(rx, complete_col):
            best = max(best, rn)
    return best


def _first_data_row(rows, header_row, complete_col='D'):
    cand = [rn for rn, rx in rows if rn > header_row and _row_has_numeric(rx, complete_col)]
    return min(cand) if cand else header_row + 1


def _col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _to_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def _emit_cell(col, rn, style, kind, payload):
    s_attr = f' s="{style}"' if style else ''
    ref = f'{col}{rn}'
    if kind == 'formula':
        return f'<c r="{ref}"{s_attr}><f>{payload}</f></c>'
    if kind == 'str':
        esc = str(payload).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{esc}</t></is></c>'
    if kind == 'num':
        return f'<c r="{ref}"{s_attr}><v>{payload}</v></c>'
    if kind == 'raw':
        return payload
    return f'<c r="{ref}"{s_attr}/>'


def _cached_value_cell(template_cell_xml):
    v = re.search(r'<v>(.*?)</v>', template_cell_xml)
    if not v:
        return None
    t = re.search(r't="(\w+)"', template_cell_xml)
    return (t.group(1) if t else None), v.group(1)


def _merge_row(template_xml, template_rn, target_xml, target_rn, values,
               carry_cols=(), zero_cols=(), blank_cols=(), formula_cols=None):
    delta = target_rn - template_rn
    tmpl = {col: cx for col, cx in _cells_of_row(template_xml)}
    tgt = {col: cx for col, cx in _cells_of_row(target_xml)} if target_xml else {}
    tmpl_formulas = {col: re.search(r'<f>(.*?)</f>', cx).group(1)
                     for col, cx in tmpl.items() if '<f>' in cx}
    tmpl_style = {col: _style_of(cx) for col, cx in tmpl.items()}
    carry_cols, zero_cols, blank_cols = set(carry_cols), set(zero_cols), set(blank_cols)
    formula_cols = dict(formula_cols or {})

    ordered = sorted(set(tmpl) | set(tgt) | set(values) | carry_cols | zero_cols
                     | blank_cols | set(formula_cols), key=_col_to_idx)
    parts = []
    for col in ordered:
        style = _style_of(tgt.get(col, '')) or tmpl_style.get(col)
        if col in values:
            v = values[col]
            parts.append(_emit_cell(col, target_rn, style, 'str' if isinstance(v, str) else 'num', v))
        elif col in formula_cols:
            f = formula_cols[col].replace('{r}', str(target_rn))
            parts.append(_emit_cell(col, target_rn, style, 'formula', f))
        elif col in zero_cols:
            parts.append(_emit_cell(col, target_rn, style, 'num', 0))
        elif col in blank_cols:
            parts.append(_emit_cell(col, target_rn, style, 'empty', None))
        elif col in carry_cols:
            cv = _cached_value_cell(tmpl.get(col, ''))
            if cv is None:
                parts.append(_emit_cell(col, target_rn, style, 'empty', None))
            else:
                t_attr, val = cv
                t = f' t="{t_attr}"' if t_attr else ''
                s = f' s="{style}"' if style else ''
                parts.append(f'<c r="{col}{target_rn}"{s}{t}><v>{val}</v></c>')
        elif col in tmpl_formulas:
            parts.append(_emit_cell(col, target_rn, style, 'formula',
                                    shift_formula(tmpl_formulas[col], delta)))
        elif col in tgt:
            parts.append(_emit_cell(col, target_rn, style, 'raw', tgt[col]))
        else:
            parts.append(_emit_cell(col, target_rn, style, 'empty', None))
    return f'<row r="{target_rn}">' + ''.join(parts) + '</row>'


def _bump_dimension(sheet_xml, new_rn):
    m = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>', sheet_xml)
    if not m:
        return sheet_xml
    c1, r1, c2, r2 = m.groups()
    if int(r2) < new_rn:
        return sheet_xml.replace(m.group(0), f'<dimension ref="{c1}{r1}:{c2}{new_rn}"/>')
    return sheet_xml


def _rewrite_parts(xlsx_path, changed):
    """changed = {archive_name: new_bytes_or_str}. Copies all else byte-for-byte."""
    tmp = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in changed:
                nd = changed[item.filename]
                data = nd.encode('utf-8') if isinstance(nd, str) else nd
            zout.writestr(item, data)
    os.replace(tmp, xlsx_path)


def fill_sprint_row(xlsx_path, sheet_name, values, complete_col='D',
                    carry_cols=(), zero_cols=(), blank_cols=(), formula_cols=None):
    with zipfile.ZipFile(xlsx_path) as z:
        smap = _read_sheet_map(z)
        if sheet_name not in smap:
            raise KeyError(f"Sheet {sheet_name} not found. Have: {list(smap)}")
        sheet_path = smap[sheet_name]
        sheet_xml = z.read(sheet_path).decode('utf-8')

    rows = _find_rows(sheet_xml)
    hdr = _header_row(rows)
    if hdr is None:
        raise RuntimeError("Could not locate header row.")
    tmpl_rn = _last_complete_row(rows, hdr, complete_col)
    if tmpl_rn == hdr:
        raise RuntimeError("No complete data row found to use as template.")
    tmpl_xml = next(rx for rn, rx in rows if rn == tmpl_rn)
    target_rn = tmpl_rn + 1
    target_xml = next((rx for rn, rx in rows if rn == target_rn), None)

    new_row_xml = _merge_row(tmpl_xml, tmpl_rn, target_xml, target_rn, values,
                             carry_cols=carry_cols, zero_cols=zero_cols,
                             blank_cols=blank_cols, formula_cols=formula_cols)
    if target_xml is not None:
        sheet_xml = sheet_xml.replace(target_xml, new_row_xml, 1)
    else:
        sheet_xml = sheet_xml.replace('</sheetData>', new_row_xml + '</sheetData>', 1)
    sheet_xml = _bump_dimension(sheet_xml, target_rn)
    _rewrite_parts(xlsx_path, {sheet_path: sheet_xml})
    return target_rn


def add_header_cells(xlsx_path, sheet_name, header_map, style_from_col='B'):
    with zipfile.ZipFile(xlsx_path) as z:
        sheet_path = _read_sheet_map(z)[sheet_name]
        sheet_xml = z.read(sheet_path).decode('utf-8')
    rows = _find_rows(sheet_xml)
    hdr = _header_row(rows)
    hdr_xml = next(rx for rn, rx in rows if rn == hdr)
    existing = {col: cx for col, cx in _cells_of_row(hdr_xml)}
    borrow_style = _style_of(existing.get(style_from_col, '')) or None
    added, new_cells = [], []
    for col, text in header_map.items():
        if col in existing and ('<v>' in existing[col] or '<t>' in existing[col]):
            continue
        esc = str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        s = f' s="{borrow_style}"' if borrow_style else ''
        new_cells.append((_col_to_idx(col),
                          f'<c r="{col}{hdr}"{s} t="inlineStr"><is><t>{esc}</t></is></c>'))
        added.append(col)
    if not new_cells:
        return []
    all_cells = [(_col_to_idx(c), cx) for c, cx in _cells_of_row(hdr_xml)] + new_cells
    all_cells.sort(key=lambda t: t[0])
    row_open = re.match(r'(<row r="\d+"[^>]*>)', hdr_xml).group(1)
    rebuilt = row_open + ''.join(cx for _, cx in all_cells) + '</row>'
    sheet_xml = sheet_xml.replace(hdr_xml, rebuilt, 1)
    maxcol = max(_col_to_idx(c) for c in added)
    m = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>', sheet_xml)
    if m and _col_to_idx(m.group(3)) < maxcol:
        sheet_xml = sheet_xml.replace(
            m.group(0), f'<dimension ref="{m.group(1)}{m.group(2)}:{_to_letter(maxcol)}{m.group(4)}"/>')
    _rewrite_parts(xlsx_path, {sheet_path: sheet_xml})
    return added


def update_chart_windows(xlsx_path, sheet_name, last_row, height=6, complete_col='D'):
    with zipfile.ZipFile(xlsx_path) as z:
        smap = _read_sheet_map(z)
        sheet_path = smap[sheet_name]
        sheet_xml = z.read(sheet_path).decode('utf-8')
        chart_files = [n for n in z.namelist() if re.match(r'xl/charts/chart\d+\.xml$', n)]
        chart_src = {n: z.read(n).decode('utf-8') for n in chart_files}
    rows = _find_rows(sheet_xml)
    hdr = _header_row(rows)
    first = _first_data_row(rows, hdr, complete_col)
    new_start = max(last_row - height + 1, first)
    new_end = last_row
    rng = re.compile(rf'({re.escape(sheet_name)}!\$[A-Z]+\$)(\d+)(:\$[A-Z]+\$)(\d+)')
    changes = {}
    for fname, txt in chart_src.items():
        if f'{sheet_name}!' not in txt:
            continue

        def repl(m):
            old = m.group(0)
            new = f'{m.group(1)}{new_start}{m.group(3)}{new_end}'
            if old != new:
                changes.setdefault(fname, []).append((old, new))
            return new
        chart_src[fname] = rng.sub(repl, txt)
    if not changes:
        return []
    _rewrite_parts(xlsx_path, {f: chart_src[f] for f in changes})
    return [(f, o, n) for f, lst in changes.items() for o, n in lst]


def strip_sparklines(xlsx_path):
    removed = 0
    with zipfile.ZipFile(xlsx_path) as z:
        contents = {n: z.read(n) for n in z.namelist()}
    changed = {}
    for n in list(contents):
        if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            txt = contents[n].decode('utf-8')
            cnt = txt.count('<x14:sparklineGroup')
            if cnt:
                txt = re.sub(r'<x14:sparklineGroups.*?</x14:sparklineGroups>', '', txt, flags=re.S)
                txt = re.sub(r'<extLst>\s*<ext[^>]*>\s*</ext>\s*</extLst>', '', txt, flags=re.S)
                changed[n] = txt
                removed += cnt
    if changed:
        _rewrite_parts(xlsx_path, changed)
    return removed
