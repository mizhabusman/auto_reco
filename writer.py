"""
writer.py — Render the flexible workbook JSON Claude returns into a polished .xlsx.

Claude decides the sheets/columns; this renderer is generic. It builds:
  • a Summary sheet (from meta + summary.metrics + insights)
  • one sheet per entry in data["sheets"], rendered as a formatted table.
"""

from __future__ import annotations
import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- premium-ish palette (matches the app theme) --------------------------
FONT_NAME = "Calibri"
INK = "2B2622"          # warm near-black
HEADER_BG = "3C3530"    # warm charcoal
TITLE_BG = "2B2622"
ACCENT = "B5793A"       # warm amber/terracotta
GOOD_BG = "E7EFE3"
WARN_BG = "FBF0DA"
BAD_BG = "F6E1D9"
NEUTRAL_BG = "EDE8E1"
GRID = "D9D2C7"

TITLE_FONT = Font(name=FONT_NAME, size=15, bold=True, color="FFFFFF")
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
SECTION_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10, color=INK)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True, color=INK)
MUTED_FONT = Font(name=FONT_NAME, size=9, italic=True, color="8A8178")

THIN = Side(border_style="thin", color=GRID)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
CURRENCY_FMT = '#,##0.00;[Red](#,##0.00);"-"'

TONE_FILL = {
    "good": GOOD_BG, "warn": WARN_BG, "bad": BAD_BG, "neutral": NEUTRAL_BG,
}


def _fill(hexcolor: str) -> PatternFill:
    return PatternFill("solid", start_color=hexcolor, end_color=hexcolor)


def _autosize(ws, max_width: int = 55):
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        longest = 0
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.__class__.__name__ == "MergedCell" or cell.value is None:
                continue
            longest = max(longest, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(longest + 3, 12), max_width)


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _coerce(v: Any) -> Any:
    if isinstance(v, str):
        s = v.replace(",", "").strip()
        try:
            return float(s) if ("." in s) else int(s)
        except (ValueError, TypeError):
            return v
    return v


# ---------------------------------------------------------------------------
def write_reco_excel(reco: dict[str, Any]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    _write_summary(wb, reco)

    used_names = {"Summary"}
    for sheet in reco.get("sheets", []) or []:
        name = (sheet.get("name") or "Sheet").strip()[:28] or "Sheet"
        base, n = name, 2
        while name in used_names:
            name = f"{base[:25]} {n}"
            n += 1
        used_names.add(name)
        _write_table_sheet(wb, name, sheet)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _write_summary(wb: Workbook, reco: dict):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    meta = reco.get("meta", {}) or {}
    summary = reco.get("summary", {}) or {}
    metrics = summary.get("metrics", []) or []
    insights = summary.get("insights", []) or []

    SPAN = 4
    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=SPAN)
    t = ws.cell(row=1, column=1, value=reco.get("title", "Reconciliation Report"))
    t.font = TITLE_FONT
    t.fill = _fill(TITLE_BG)
    t.alignment = CENTER
    ws.row_dimensions[1].height = 34

    # accent strip
    for c in range(1, SPAN + 1):
        ws.cell(row=2, column=c).fill = _fill(ACCENT)
    ws.row_dimensions[2].height = 4

    r = 4
    info = [
        ("Party A", meta.get("party_a", "")),
        ("Party B", meta.get("party_b", "")),
        ("Period", meta.get("period", "")),
        ("Currency", meta.get("currency", "INR")),
        ("Note", meta.get("prepared_note", "")),
    ]
    for label, val in info:
        if not val:
            continue
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = BOLD_FONT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=SPAN)
        vc = ws.cell(row=r, column=2, value=val)
        vc.font = BODY_FONT
        vc.alignment = LEFT
        r += 1
    r += 1

    # Metrics block
    if metrics:
        _section(ws, r, "Balances & Totals", SPAN)
        r += 1
        head = ["Metric", meta.get("party_a", "Party A"), meta.get("party_b", "Party B"), "Difference"]
        for i, h in enumerate(head, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.font = HEADER_FONT
            c.fill = _fill(HEADER_BG)
            c.alignment = CENTER
            c.border = BORDER
        r += 1
        for m in metrics:
            vals = [m.get("label", ""), _coerce(m.get("value_a", "")),
                    _coerce(m.get("value_b", "")), _coerce(m.get("difference", ""))]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                c.border = BORDER
                c.font = BOLD_FONT if i == 1 else BODY_FONT
                if i > 1 and _is_number(v):
                    c.number_format = CURRENCY_FMT
                    c.alignment = Alignment(horizontal="right", vertical="center")
            r += 1
        r += 1

    # Insights block
    if insights:
        _section(ws, r, "Key Observations", SPAN)
        r += 1
        for ins in insights:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=SPAN)
            c = ws.cell(row=r, column=1, value=f"•  {ins}")
            c.font = BODY_FONT
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = max(18, 15 * (1 + len(str(ins)) // 90))
            r += 1

    for col, w in [(1, 26), (2, 26), (3, 26), (4, 22)]:
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.sheet_properties.tabColor = ACCENT


def _section(ws, row: int, text: str, span: int):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = SECTION_FONT
    c.fill = _fill(ACCENT)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def _write_table_sheet(wb: Workbook, name: str, sheet: dict):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    columns = sheet.get("columns", []) or []
    rows = sheet.get("rows", []) or []
    subtitle = sheet.get("subtitle")
    tone = (sheet.get("tone") or "neutral").lower()
    span = max(len(columns), 1)

    ws.sheet_properties.tabColor = TONE_FILL.get(tone, NEUTRAL_BG)

    # Title bar
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    t = ws.cell(row=1, column=1, value=name)
    t.font = TITLE_FONT
    t.fill = _fill(TITLE_BG)
    t.alignment = CENTER
    ws.row_dimensions[1].height = 30

    header_row = 2
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        s = ws.cell(row=2, column=1, value=subtitle)
        s.font = MUTED_FONT
        s.fill = _fill(TONE_FILL.get(tone, NEUTRAL_BG))
        s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        header_row = 3

    # Column headers
    for i, col in enumerate(columns, 1):
        c = ws.cell(row=header_row, column=i, value=col)
        c.font = HEADER_FONT
        c.fill = _fill(HEADER_BG)
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[header_row].height = 26

    if not rows:
        ws.merge_cells(start_row=header_row + 1, start_column=1,
                       end_row=header_row + 1, end_column=span)
        c = ws.cell(row=header_row + 1, column=1, value="— No items in this category —")
        c.font = MUTED_FONT
        c.alignment = CENTER
        _autosize(ws)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
        return

    band = _fill("FBFAF7")
    for ridx, row in enumerate(rows):
        excel_row = header_row + 1 + ridx
        cells = list(row) + [""] * (len(columns) - len(row))  # pad short rows
        for i, v in enumerate(cells[:len(columns)], 1):
            val = _coerce(v)
            c = ws.cell(row=excel_row, column=i, value=val)
            c.border = BORDER
            c.font = BODY_FONT
            if _is_number(val):
                c.number_format = CURRENCY_FMT
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = LEFT
            if ridx % 2 == 1:
                c.fill = band

    _autosize(ws)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)