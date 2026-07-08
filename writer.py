"""
writer.py — Turn Claude's free-form markdown reconciliation into a neat Excel.

No schema. We just parse whatever Claude wrote — headings, paragraphs, bullet
points and markdown tables — and lay it out on a single clean sheet, rendering
any tables as proper bordered Excel tables.
"""
from __future__ import annotations
import io
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BLUE       = "1F4E78"
LIGHT_BLUE = "DDEBF7"
INK        = "272220"
BORDER_GREY = "BFBFBF"
NUM_FMT    = '#,##0.00;[Red]-#,##0.00'


def _font(size=11, bold=False, color=None, italic=False):
    kw = dict(name="Calibri", size=size, bold=bold, italic=italic)
    if color:
        kw["color"] = color
    return Font(**kw)

def _thin():
    s = Side(style="thin", color=BORDER_GREY)
    return Border(left=s, right=s, top=s, bottom=s)

def _clean(text: str) -> str:
    """Strip common markdown inline markup."""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)   # bold
    text = re.sub(r"(?<!\*)\*(?!\*)(.*?)\*", r"\1", text)  # italic
    text = re.sub(r"`(.*?)`", r"\1", text)          # code
    text = re.sub(r"\[(.*?)\]\(.*?\)", r"\1", text) # links
    return text.strip()

def _num(v):
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("₹", "").replace("$", "").replace("%", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", s):
            try:
                return float(s)
            except ValueError:
                return None
    return None

def _is_separator(line: str) -> bool:
    """A markdown table separator row, e.g. |---|:--:|---|."""
    return bool(re.fullmatch(r"\s*\|?[\s:|-]*-[\s:|-]*\|?\s*", line)) and "-" in line

def _split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [_clean(c) for c in line.split("|")]


def _write_table(ws, start_row, header, rows):
    ncols = max([len(header)] + [len(r) for r in rows])
    header = header + [""] * (ncols - len(header))
    r = start_row

    for ci, h in enumerate(header, 1):
        c = ws.cell(r, ci, h)
        c.font = _font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=BLUE, end_color=BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _thin()
    ws.row_dimensions[r].height = 20
    r += 1

    for row in rows:
        row = row + [""] * (ncols - len(row))
        for ci, v in enumerate(row[:ncols], 1):
            c = ws.cell(r, ci)
            c.border = _thin()
            n = _num(v)
            if n is not None:
                c.value = n
                c.number_format = NUM_FMT
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.font = _font(color="9C0006") if n < 0 else _font()
            else:
                c.value = v
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c.font = _font()
        r += 1
    return r, ncols


def report_to_excel(report: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    ws.sheet_view.showGridLines = False

    lines = (report or "").split("\n")
    r, i, maxcol = 1, 0, 1

    while i < len(lines):
        line = lines[i].rstrip()

        # Markdown table: a row with "|" followed by a separator row.
        if "|" in line and i + 1 < len(lines) and _is_separator(lines[i + 1]):
            header = _split_row(line)
            i += 2
            body = []
            while i < len(lines) and "|" in lines[i] and lines[i].strip():
                body.append(_split_row(lines[i]))
                i += 1
            r, used = _write_table(ws, r, header, body)
            maxcol = max(maxcol, used)
            r += 1
            continue

        heading = re.match(r"^(#{1,6})\s+(.*)", line)
        bullet  = re.match(r"^\s*[-*+]\s+(.*)", line)

        if heading:
            level, text = len(heading.group(1)), _clean(heading.group(2))
            size = 16 if level == 1 else 13 if level == 2 else 12
            c = ws.cell(r, 1, text)
            c.font = _font(size=size, bold=True, color=BLUE if level <= 2 else INK)
            ws.row_dimensions[r].height = 22 if level <= 2 else 18
            r += 1
        elif re.fullmatch(r"\s*([-*_])\1{2,}\s*", line):  # horizontal rule
            r += 1
        elif not line.strip():
            r += 1
        else:
            text = ("•  " + _clean(bullet.group(1))) if bullet else _clean(line)
            c = ws.cell(r, 1, text)
            c.font = _font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        i += 1

    # Auto-fit columns from actual content.
    for ci in range(1, maxcol + 1):
        longest = 12
        for row in ws.iter_rows(min_col=ci, max_col=ci):
            v = row[0].value
            if v is not None:
                longest = max(longest, len(str(v)))
        ws.column_dimensions[get_column_letter(ci)].width = min(max(longest + 3, 14), 70)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
