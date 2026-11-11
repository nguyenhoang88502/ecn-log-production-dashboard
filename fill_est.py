# -*- coding: utf-8 -*-
"""Dry run of Prompt_Claude_Excel.md on a COPY of the workbook (same rules as the prompt, per ROW).
   Output: _filled_test.xlsx + printed summary. The real ECN Log.xlsx is never modified."""
import re, sys, warnings, collections, datetime as dt
import openpyxl
from openpyxl.styles import PatternFill
import ecn_parse as p

warnings.filterwarnings("ignore")
SRC = sys.argv[1] if len(sys.argv) > 1 else "_new_ecn_log.xlsx"
DST = sys.argv[2] if len(sys.argv) > 2 else "_filled_test.xlsx"
NL, CR = chr(10), chr(13)

wb = openpyxl.load_workbook(SRC)
ws = wb["ECN"]
hdr_row = next(i for i in range(1, 11) if ws.cell(i, 1).value == "ECN Number")
col = {p._norm(ws.cell(hdr_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(hdr_row, c).value}
cR, cS = col["New item (est MP week)"], col["New item (est MP year)"]
cA, cB, cN, cV = col["ECN Number"], col["Date"], col["Date Issued"], col["Note"]


def d_of(v):
    return v.date() if isinstance(v, dt.datetime) else (v if isinstance(v, dt.date) else None)


def lines_of(cell):
    return [x.strip() for x in str(cell).replace(CR, "").split(NL) if x.strip()] if cell else []


def first_move(cells, ref):
    """top-down over the cells: first line with a week AND change -> (week, year or None, line)."""
    for c in cells:
        for line in lines_of(c):
            n = p.parse_note_line(line, ref)
            im, mw, my = p.move_info(n["text"])
            if im and mw is not None:
                return mw, my, line
    return None


def note_year(cell, base_year):
    ys = [int(t) for line in lines_of(cell) for t in re.findall(r"20\d\d", line)]
    ys = [y for y in ys if y > base_year]
    return max(ys) if ys else None


rows_by_ecn = collections.OrderedDict()
for r in range(hdr_row + 1, ws.max_row + 1):
    k = ws.cell(r, cA).value
    if k:
        rows_by_ecn.setdefault(str(k).strip(), []).append(r)

yellow = PatternFill("solid", fgColor="FFF2CC")


def s_blank(r):
    return ws.cell(r, cS).value in (None, "")

log, counts = [], collections.Counter()
for k, rows in rows_by_ecn.items():
    ecn_date = next((d_of(ws.cell(r, cB).value) for r in rows if d_of(ws.cell(r, cB).value)), None)
    ecn_cells = list(dict.fromkeys(str(ws.cell(r, cV).value).strip() for r in rows if ws.cell(r, cV).value))
    for r in rows:
        r_old = str(ws.cell(r, cR).value).strip() if ws.cell(r, cR).value is not None else None
        valid = bool(r_old and re.fullmatch(r"W\d\d", r_old))
        date = d_of(ws.cell(r, cB).value) or ecn_date or d_of(ws.cell(r, cN).value)
        own = [ws.cell(r, cV).value] if ws.cell(r, cV).value else []
        wk = source = line = None
        if valid:
            wk, source = int(r_old[1:]), "Giữ nguyên"
        else:
            mv = first_move(own, date or dt.date(2026, 1, 1)) or first_move(ecn_cells, date or dt.date(2026, 1, 1))
            if mv:
                wk, source, line = mv[0], "Note", mv[2]
            elif note_year(ws.cell(r, cV).value, (date or dt.date(2026, 1, 1)).isocalendar()[0]):
                source = "Note (chỉ năm)"          # Note says it is later than this year, no week given
            elif date:
                wk, source = date.isocalendar()[1], "Date"
        yr = None
        if source == "Note (chỉ năm)":
            yr = note_year(ws.cell(r, cV).value, (date or dt.date(2026, 1, 1)).isocalendar()[0])
            if s_blank(r):
                ws.cell(r, cS).value = yr
                ws.cell(r, cS).fill = yellow
        if wk is not None:
            base_year = date.isocalendar()[0] if date else (d_of(ws.cell(r, cN).value) or dt.date(2026, 1, 1)).isocalendar()[0]
            yr = base_year
            if date and date.isocalendar()[1] - wk > 26:
                yr = base_year + 1
            ny = note_year(ws.cell(r, cV).value, base_year)
            if ny:
                yr = ny
        s_old = ws.cell(r, cS).value
        if wk is not None:
            if not valid:
                ws.cell(r, cR).value = "W%02d" % wk
                ws.cell(r, cR).number_format = "@"
                ws.cell(r, cR).fill = yellow
            if s_old in (None, ""):
                ws.cell(r, cS).value = yr
                ws.cell(r, cS).fill = yellow
        counts[source or "Để trống"] += 1
        log.append([r, k, r_old, "W%02d" % wk if wk else None, yr, source or "Để trống", line])

ls = wb.create_sheet("ECN fill log")
ls.append(["Dòng", "ECN Number", "R cũ", "R mới", "S mới", "Nguồn", "Dòng Note đã dùng"])
for x in log:
    ls.append(x)
wb.save(DST)

print("rows:", len(log), dict(counts))
print("from Note:", [(x[1], x[3], x[4], x[6][:45]) for x in log if x[5] == "Note"])
print("year 2027:", sorted({(x[1], x[3], x[4]) for x in log if x[4] and x[4] > 2026}))
print("Note year only:", sorted({(x[1], x[4]) for x in log if x[5] == "Note (chỉ năm)"}))
print("blank:", [(x[0], x[1]) for x in log if x[5] == "Để trống"])
