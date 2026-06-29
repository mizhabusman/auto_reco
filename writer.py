"""
writer.py — Professional Excel formatter for reconciliation reports.

Claude returns the DATA (sheets + rows).
This file turns that data into a beautiful, readable Excel file.

Palette: warm charcoal headers, category-coded sheet tabs,
         zebra rows, auto-width columns, freeze panes, SUM formulas.
"""

from __future__ import annotations
import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ── Palette ───────────────────────────────────────────────────────────────────
FONT = "Arial"

# Header row
HDR_BG   = "2B2622"   # warm charcoal
HDR_FG   = "FFFFFF"

# Sheet-type accent colours (tab + header accent strip)
TONE_COLORS = {
    "summary":  "B5793A",   # warm amber
    "matched":  "3A7D52",   # forest green
    "diff":     "C0392B",   # warm red
    "tds":      "E67E22",   # orange
    "onlya":    "8E44AD",   # purple
    "onlyb":    "2980B9",   # blue
    "other":    "546E7A",   # slate
}

# Zebra row fills
ZEBRA_LIGHT = "FAF9F6"
ZEBRA_DARK  = "F0EBE3"

# Number / currency format
NUM_FMT = '#,##0.00;(#,##0.00);"-"'


def _fill(hex6: str) -> PatternFill:
    return PatternFill("solid", start_color=hex6, end_color=hex6)

def _font(bold=False, color="000000", size=10, italic=False) -> Font:
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)

def _border(style="thin", color="D9D2C7") -> Border:
    s = Side(border_style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _center(wrap=False) -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _left(wrap=True) -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

def _right() -> Alignment:
    return Alignment(horizontal="right", vertical="center")


# ── Tone detection ─────────────────────────────────────────────────────────────
def _detect_tone(sheet_name: str) -> str:
    n = sheet_name.lower()
    if any(k in n for k in ("summary", "overview", "dashboard")):  return "summary"
    if any(k in n for k in ("match",)):                            return "matched"
    if any(k in n for k in ("diff", "mismatch", "discrepan")):    return "diff"
    if any(k in n for k in ("tds", "tax", "withhold")):           return "tds"
    if any(k in n for k in ("only in a", "only a", "missing b", "unmatched a")): return "onlya"
    if any(k in n for k in ("only in b", "only b", "missing a", "unmatched b")): return "onlyb"
    return "other"


# ── Column auto-width ──────────────────────────────────────────────────────────
def _autowidth(ws, min_w=8, max_w=52):
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        best = min_w
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.__class__.__name__ == "MergedCell" or cell.value is None:
                continue
            best = max(best, min(len(str(cell.value)) + 3, max_w))
        ws.column_dimensions[letter].width = best


# ── Is a value numeric? ───────────────────────────────────────────────────────
def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def _coerce(v):
    """Try to coerce string numbers; leave everything else as-is."""
    if isinstance(v, str):
        cleaned = v.replace(",", "").replace("₹", "").strip()
        try:    return float(cleaned) if "." in cleaned else int(cleaned)
        except: return v
    return v


# ── Sheet writers ──────────────────────────────────────────────────────────────
def _write_cover(wb: Workbook, all_sheets: list[dict], meta: dict):
    """First sheet: a visual cover / table of contents."""
    ws = wb.create_sheet("📋 Report", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TONE_COLORS["summary"]

    # Title block
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "LEDGER RECONCILIATION REPORT"
    c.font  = Font(name=FONT, size=16, bold=True, color=HDR_FG)
    c.fill  = _fill(HDR_BG)
    c.alignment = _center()
    ws.row_dimensions[1].height = 38

    # Accent strip
    for col in range(1, 6):
        ws.cell(row=2, column=col).fill = _fill(TONE_COLORS["summary"])
    ws.row_dimensions[2].height = 5

    # Meta rows
    r = 4
    for label, key in [("Party A", "party_a"), ("Party B", "party_b"),
                        ("Period",  "period"),  ("Prepared", "prepared")]:
        val = meta.get(key, "")
        if not val: continue
        lc = ws.cell(row=r, column=1, value=label)
        vc = ws.cell(row=r, column=2, value=val)
        lc.font = _font(bold=True, color=HDR_BG)
        vc.font = _font(color="2B2622")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        r += 1
    r += 1

    # Contents table
    ws.cell(row=r, column=1, value="Sheet").font   = _font(bold=True, color=HDR_FG)
    ws.cell(row=r, column=1).fill = _fill(HDR_BG)
    ws.cell(row=r, column=2, value="Rows").font    = _font(bold=True, color=HDR_FG)
    ws.cell(row=r, column=2).fill = _fill(HDR_BG)
    ws.cell(row=r, column=3, value="Description").font = _font(bold=True, color=HDR_FG)
    ws.cell(row=r, column=3).fill = _fill(HDR_BG)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    r += 1

    DESCRIPTIONS = {
        "matched":  "Transactions that agree in both ledgers",
        "diff":     "Amount or value differences between books",
        "tds":      "Entries matched only after TDS adjustment",
        "onlya":    "Present in Party A's books only",
        "onlyb":    "Present in Party B's books only",
        "summary":  "Key figures and totals",
        "other":    "Additional reconciliation details",
    }
    for sh in all_sheets:
        tone  = _detect_tone(sh["name"])
        n_rows = max(0, len(sh.get("rows", [])) - 1)
        color  = TONE_COLORS[tone]
        c1 = ws.cell(row=r, column=1, value=sh["name"])
        c2 = ws.cell(row=r, column=2, value=n_rows)
        c3 = ws.cell(row=r, column=3, value=DESCRIPTIONS.get(tone, ""))
        c1.font = _font(bold=True, color=color)
        c2.font = _font(color="2B2622")
        c2.alignment = _center()
        c3.font = _font(italic=True, color="6D6560")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        if r % 2 == 0:
            for col in range(1, 6):
                ws.cell(row=r, column=col).fill = _fill(ZEBRA_DARK)
        r += 1

    for col, w in [(1, 28), (2, 10), (3, 50)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def _write_data_sheet(wb: Workbook, sh: dict):
    rows = sh.get("rows", [])
    if not rows:
        return

    name = sh["name"]
    tone = _detect_tone(name)
    color = TONE_COLORS[tone]

    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = color

    headers = [str(h) for h in rows[0]]
    data    = rows[1:]
    n_cols  = len(headers)

    # ── Title bar ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws["A1"]
    tc.value     = name.upper()
    tc.font      = Font(name=FONT, size=13, bold=True, color=HDR_FG)
    tc.fill      = _fill(HDR_BG)
    tc.alignment = _center()
    ws.row_dimensions[1].height = 30

    # Accent strip under title
    for col in range(1, n_cols + 1):
        ws.cell(row=2, column=col).fill = _fill(color)
    ws.row_dimensions[2].height = 4

    # ── Header row ──
    HDR_ROW = 3
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=HDR_ROW, column=ci, value=h)
        c.font      = _font(bold=True, color=HDR_FG, size=10)
        c.fill      = _fill("4A4440")   # slightly lighter than title
        c.alignment = _center(wrap=True)
        c.border    = _border()
    ws.row_dimensions[HDR_ROW].height = 24
    ws.freeze_panes = ws.cell(row=HDR_ROW + 1, column=1)

    # ── Data rows ──
    num_cols = set()   # track which columns are numeric for totals
    for ri, row in enumerate(data):
        excel_row = HDR_ROW + 1 + ri
        bg = ZEBRA_LIGHT if ri % 2 == 0 else ZEBRA_DARK
        padded = list(row) + [""] * max(0, n_cols - len(row))

        for ci, raw_val in enumerate(padded[:n_cols], 1):
            val = _coerce(raw_val)
            c   = ws.cell(row=excel_row, column=ci, value=val)
            c.fill   = _fill(bg)
            c.border = _border()
            c.font   = _font(size=10)

            if _is_num(val):
                c.number_format = NUM_FMT
                c.alignment     = _right()
                num_cols.add(ci)
            else:
                c.alignment = _left()

    # ── Totals row (only if there are numeric cols and data rows) ──
    if data and num_cols:
        total_row = HDR_ROW + 1 + len(data)
        ws.merge_cells(start_row=total_row, start_column=1,
                       end_row=total_row, end_column=max(1, min(num_cols) - 1))
        tc2 = ws.cell(row=total_row, column=1, value="TOTAL")
        tc2.font      = _font(bold=True, color=HDR_FG)
        tc2.fill      = _fill(HDR_BG)
        tc2.alignment = _right()
        tc2.border    = _border()

        first_data = HDR_ROW + 1
        last_data  = HDR_ROW + len(data)
        for ci in range(1, n_cols + 1):
            c = ws.cell(row=total_row, column=ci)
            if ci in num_cols:
                col_letter = get_column_letter(ci)
                c.value         = f"=SUM({col_letter}{first_data}:{col_letter}{last_data})"
                c.number_format = NUM_FMT
                c.font          = _font(bold=True, color=HDR_FG)
                c.fill          = _fill(HDR_BG)
                c.alignment     = _right()
                c.border        = _border()
            elif ci > 1:
                c.fill   = _fill(HDR_BG)
                c.border = _border()

    # Empty-sheet placeholder
    if not data:
        ws.merge_cells(start_row=HDR_ROW + 1, start_column=1,
                       end_row=HDR_ROW + 1, end_column=n_cols)
        ec = ws.cell(row=HDR_ROW + 1, column=1, value="— No entries in this category —")
        ec.font      = _font(italic=True, color="9E9590")
        ec.alignment = _center()

    _autowidth(ws)


# ── Main entry point ───────────────────────────────────────────────────────────
def sheets_to_excel(sheets: list[dict], summary_text: str = "", meta: dict = None) -> bytes:
    """
    sheets      : list of {"name": str, "rows": [[header...], [row...], ...]}
    summary_text: plain-text CA summary (written into first sheet notes)
    meta        : {"party_a":..., "party_b":..., "period":..., "prepared":...}
    """
    meta = meta or {}
    wb   = Workbook()
    wb.remove(wb.active)   # remove default blank sheet

    # Cover sheet first
    _write_cover(wb, sheets, meta)

    # Data sheets
    seen: dict[str, int] = {}
    for sh in sheets:
        raw_name = str(sh.get("name", "Sheet"))
        for ch in r'/\?*[]:\x00': raw_name = raw_name.replace(ch, "-")
        name = raw_name[:31].strip() or "Sheet"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = name[:28] + f" {seen[name]}"
        sh = dict(sh); sh["name"] = name   # use sanitised name
        _write_data_sheet(wb, sh)

    # Summary note tab (CA's plain-English summary)
    if summary_text:
        ws = wb.create_sheet("📝 Notes")
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = TONE_COLORS["other"]
        ws.merge_cells("A1:F1")
        t = ws["A1"]
        t.value     = "RECONCILIATION NOTES"
        t.font      = Font(name=FONT, size=13, bold=True, color=HDR_FG)
        t.fill      = _fill(HDR_BG)
        t.alignment = _center()
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A3:F3")
        h = ws["A3"]
        h.value     = "Summary prepared by AI Reconciliation Assistant"
        h.font      = _font(bold=True, color=HDR_BG)

        # Split summary into lines and write
        lines = [l.strip() for l in summary_text.strip().splitlines() if l.strip()]
        for i, line in enumerate(lines, start=5):
            ws.merge_cells(f"A{i}:F{i}")
            c = ws.cell(row=i, column=1, value=line)
            c.font      = _font(size=10)
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.row_dimensions[i].height = max(18, 14 * (1 + len(line) // 90))

        for col, w in [(1,12),(2,14),(3,14),(4,14),(5,14),(6,16)]:
            ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()