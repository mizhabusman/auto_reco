"""
writer.py — Professional Excel formatter for reconciliation reports.
Claude returns the data. This makes it beautiful.
"""
from __future__ import annotations
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Palette ───────────────────────────────────────────────────────────────────
F = "Arial"

# Warm professional palette
C = {
    "title_bg":   "1C1916",
    "title_fg":   "FFFFFF",
    "hdr_bg":     "3C3530",
    "hdr_fg":     "FFFFFF",
    "total_bg":   "2B2622",
    "total_fg":   "FFFFFF",
    "accent":     "B5793A",  # amber
    "green":      "2D6A4F",
    "red":        "9B2226",
    "orange":     "CA6702",
    "purple":     "5E548E",
    "blue":       "1D6FA4",
    "slate":      "495867",
    "z1":         "FFFFFF",
    "z2":         "F7F3EE",
    "cover_bg":   "FAF7F2",
    "line":       "DDD5C8",
}

# tone → accent colour for tab + header strip
TONE = {
    "summary":  C["accent"],
    "matched":  C["green"],
    "diff":     C["red"],
    "tds":      C["orange"],
    "onlya":    C["purple"],
    "onlyb":    C["blue"],
    "other":    C["slate"],
}

NUM_FMT  = '#,##0.00;(#,##0.00);"-"'
INT_FMT  = '#,##0;(#,##0);"-"'
DATE_FMT = "DD-MMM-YYYY"

def _f(hex6): return PatternFill("solid", start_color=hex6, end_color=hex6)
def _c(bold=False, color="000000", size=10, italic=False):
    return Font(name=F, bold=bold, color=color, size=size, italic=italic)
def _b(color=C["line"]):
    s = Side(border_style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def _a(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _tone(name):
    n = name.lower()
    if any(k in n for k in ("summary","overview","dashboard","reconciliation summary")): return "summary"
    if any(k in n for k in ("matched","match","agree","payment","invoice","credit note")): return "matched"
    if any(k in n for k in ("diff","mismatch","discrepan","issue","varianc","key diff")): return "diff"
    if any(k in n for k in ("tds","tax","withhold")): return "tds"
    if any(k in n for k in ("only in a","only a","missing b","unmatched a")): return "onlya"
    if any(k in n for k in ("only in b","only b","missing a","unmatched b")): return "onlyb"
    return "other"

def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def _coerce(v):
    if isinstance(v, str):
        s = v.replace(",","").replace("₹","").replace("$","").strip()
        if s in ("-",""):  return v
        try:    return int(s) if "." not in s else float(s)
        except: return v
    return v

def _autowidth(ws, min_w=10, max_w=48):
    for ci in range(1, ws.max_column + 1):
        col = get_column_letter(ci)
        best = min_w
        for ri in range(1, ws.max_row + 1):
            cell = ws.cell(ri, ci)
            if cell.__class__.__name__ == "MergedCell" or cell.value is None: continue
            best = max(best, min(len(str(cell.value)) + 3, max_w))
        ws.column_dimensions[col].width = best

# ── Cover sheet ───────────────────────────────────────────────────────────────
def _cover(wb, sheets, meta):
    ws = wb.create_sheet("📋 Report", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C["accent"]
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 42

    # Title
    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = "LEDGER RECONCILIATION REPORT"
    t.font  = Font(name=F, size=16, bold=True, color=C["title_fg"])
    t.fill  = _f(C["title_bg"])
    t.alignment = _a("center","center")
    ws.row_dimensions[1].height = 40

    # Amber accent strip
    for c in "ABC":
        ws[f"{c}2"].fill = _f(C["accent"])
    ws.row_dimensions[2].height = 5

    # Meta info
    r = 4
    for label, key in [("Party A","party_a"),("Party B","party_b"),
                        ("Period","period"),("Prepared","prepared")]:
        val = meta.get(key,"")
        if not val: continue
        lc = ws.cell(r, 1, label)
        vc = ws.cell(r, 2, val)
        lc.font = _c(bold=True, color=C["accent"])
        vc.font = _c(color=C["title_bg"])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1
    r += 1

    # Contents table header
    for ci, (col, label) in enumerate([(1,"Sheet"),(2,"Rows"),(3,"Category")], 1):
        c = ws.cell(r, ci, label)
        c.font  = _c(bold=True, color=C["hdr_fg"])
        c.fill  = _f(C["hdr_bg"])
        c.alignment = _a("center")
        c.border = _b()
    ws.row_dimensions[r].height = 22
    r += 1

    CATS = {
        "summary": "Key figures & totals",
        "matched": "Transactions matched in both books",
        "diff":    "Differences & discrepancies",
        "tds":     "TDS / tax-adjusted entries",
        "onlya":   "Only in Party A's books",
        "onlyb":   "Only in Party B's books",
        "other":   "Additional details",
    }
    for i, sh in enumerate(sheets):
        tone = _tone(sh["name"])
        color = TONE[tone]
        n_rows = max(0, len(sh.get("rows",[])) - 1)
        c1 = ws.cell(r, 1, sh["name"])
        c2 = ws.cell(r, 2, n_rows)
        c3 = ws.cell(r, 3, CATS.get(tone,""))
        c1.font = _c(bold=True, color=color)
        c2.font = _c(); c2.alignment = _a("center")
        c3.font = _c(italic=True, color="6D6560")
        for c in (c1,c2,c3): c.border = _b()
        if i % 2 == 0:
            for ci in range(1,4): ws.cell(r,ci).fill = _f(C["z2"])
        ws.row_dimensions[r].height = 17
        r += 1

# ── Data sheet ────────────────────────────────────────────────────────────────
def _sheet(wb, sh):
    rows = sh.get("rows", [])
    if not rows: return

    name  = sh["name"]
    tone  = _tone(name)
    color = TONE[tone]

    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = color

    headers = [str(h) for h in rows[0]]
    data    = rows[1:]
    ncols   = len(headers)

    # ── Title bar ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws["A1"]
    t.value     = name.upper()
    t.font      = Font(name=F, size=13, bold=True, color=C["title_fg"])
    t.fill      = _f(C["title_bg"])
    t.alignment = _a("center","center")
    ws.row_dimensions[1].height = 32

    # ── Colour accent strip ──
    for ci in range(1, ncols+1):
        ws.cell(2, ci).fill = _f(color)
    ws.row_dimensions[2].height = 5

    # ── Header row ──
    HR = 3
    for ci, h in enumerate(headers, 1):
        c = ws.cell(HR, ci, h)
        c.font      = _c(bold=True, color=C["hdr_fg"], size=10)
        c.fill      = _f(C["hdr_bg"])
        c.alignment = _a("center","center", wrap=True)
        c.border    = _b("AAAAAA")
    ws.row_dimensions[HR].height = 26
    ws.freeze_panes = ws.cell(HR+1, 1)

    # ── Detect numeric columns from header keywords ──
    num_kw = ("amount","debit","credit","balance","tds","total","variance",
              "diff","impact","payment","invoice","seq","value","qty")
    def is_num_col(h):
        return any(k in h.lower() for k in num_kw)

    # ── Data rows ──
    num_cols = set()
    for ri, row in enumerate(data):
        er  = HR + 1 + ri
        bg  = C["z1"] if ri % 2 == 0 else C["z2"]
        pad = list(row) + [""] * max(0, ncols - len(row))

        for ci, raw in enumerate(pad[:ncols], 1):
            val = _coerce(raw)
            c   = ws.cell(er, ci, val)
            c.fill   = _f(bg)
            c.border = _b()
            c.font   = _c(size=10)

            if _is_num(val):
                num_cols.add(ci)
                fmt = INT_FMT if isinstance(val, int) else NUM_FMT
                c.number_format = fmt
                c.alignment = _a("right")
                # colour negative values red
                if val < 0:
                    c.font = _c(color=C["red"], size=10)
            else:
                c.alignment = _a("left", wrap=True)
                # status/risk colouring
                sv = str(val).strip().upper()
                if sv in ("MATCHED","CLEAR","LOW"):
                    c.font = _c(bold=True, color=C["green"])
                elif sv in ("UNRESOLVED","HIGH"):
                    c.font = _c(bold=True, color=C["red"])
                elif sv in ("MINOR","MEDIUM","ACCEPTABLE","SUBSTANTIALLY RECONCILED"):
                    c.font = _c(bold=True, color=C["orange"])

    # ── Totals row (SUM formula — only numeric cols, only if data exists) ──
    if data and num_cols:
        tr = HR + 1 + len(data)
        # Merge non-numeric leading columns as "TOTAL" label
        first_num = min(num_cols)
        if first_num > 1:
            ws.merge_cells(start_row=tr, start_column=1,
                           end_row=tr, end_column=first_num - 1)
        label = ws.cell(tr, 1, "TOTAL")
        label.font      = _c(bold=True, color=C["total_fg"])
        label.fill      = _f(C["total_bg"])
        label.alignment = _a("right")
        label.border    = _b()

        first_data = HR + 1
        last_data  = HR + len(data)
        for ci in range(1, ncols+1):
            c = ws.cell(tr, ci)
            if ci in num_cols:
                col_l = get_column_letter(ci)
                # Use Excel SUM formula — calculated by Excel, not Python
                c.value         = f"=SUM({col_l}{first_data}:{col_l}{last_data})"
                c.number_format = NUM_FMT
                c.font          = _c(bold=True, color=C["total_fg"])
                c.fill          = _f(C["total_bg"])
                c.alignment     = _a("right")
                c.border        = _b()
            elif ci > 1:
                c.fill   = _f(C["total_bg"])
                c.border = _b()

    # ── Empty placeholder ──
    if not data:
        ws.merge_cells(start_row=HR+1, start_column=1, end_row=HR+1, end_column=ncols)
        ec = ws.cell(HR+1, 1, "— No entries in this category —")
        ec.font      = _c(italic=True, color="9E9590")
        ec.alignment = _a("center")

    _autowidth(ws)

# ── Notes sheet ───────────────────────────────────────────────────────────────
def _notes(wb, summary_text):
    ws = wb.create_sheet("📝 Notes")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C["slate"]

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value     = "RECONCILIATION NOTES"
    t.font      = Font(name=F, size=13, bold=True, color=C["title_fg"])
    t.fill      = _f(C["title_bg"])
    t.alignment = _a("center","center")
    ws.row_dimensions[1].height = 32

    for c in "ABCDE": ws[f"{c}2"].fill = _f(C["slate"])
    ws.row_dimensions[2].height = 5

    ws.cell(4, 1, "Summary prepared by AI Reconciliation Assistant"
            ).font = _c(bold=True, color=C["accent"])

    r = 6
    for line in summary_text.strip().splitlines():
        line = line.strip()
        if not line: continue
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, line)
        c.font      = _c(size=10)
        c.alignment = _a("left","top", wrap=True)
        ws.row_dimensions[r].height = max(18, 13 * (1 + len(line)//90))
        r += 1

    for col, w in [(1,14),(2,14),(3,14),(4,14),(5,18)]:
        ws.column_dimensions[get_column_letter(col)].width = w

# ── Main ──────────────────────────────────────────────────────────────────────
def sheets_to_excel(sheets: list[dict],
                    summary_text: str = "",
                    meta: dict = None) -> bytes:
    meta = meta or {}
    wb   = Workbook()
    wb.remove(wb.active)

    _cover(wb, sheets, meta)

    seen = {}
    for sh in sheets:
        raw = str(sh.get("name","Sheet"))
        for ch in r'/\?*[]:\x00': raw = raw.replace(ch,"-")
        name = raw[:31].strip() or "Sheet"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1: name = name[:28] + f" {seen[name]}"
        _sheet(wb, {**sh, "name": name})

    if summary_text:
        _notes(wb, summary_text)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()