"""
writer.py — Renders the 7-sheet reconciliation report exactly matching
the example template (Summary / TDS Reconciliation / Matched / Amount
Mismatches / Missing in Their Books / Missing in Our Books / Timing Differences).

Styling: Excel-classic — deep blue title #1F4E78, amber alert #FFEB9C,
Calibri body, white table headers on blue.
"""
from __future__ import annotations
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Palette ───────────────────────────────────────────────────────────────────
BLUE       = "1F4E78"
LIGHT_BLUE = "DDEBF7"
AMBER_BG   = "FFEB9C"   # warning yellow
AMBER_FG   = "9C5700"
GREEN_BG   = "C6EFCE"   # excel good
GREEN_FG   = "006100"
RED_BG     = "FFC7CE"   # excel bad
RED_FG     = "9C0006"
BORDER_GREY = "BFBFBF"

NUM_FMT  = '#,##0.00;[Red](#,##0.00);"-"'
INT_FMT  = '#,##0'

def _f(hex6):
    return PatternFill("solid", start_color=hex6, end_color=hex6)

def _font(size=11, bold=False, color=None, italic=False):
    kw = dict(name="Calibri", size=size, bold=bold, italic=italic)
    if color: kw["color"] = color
    return Font(**kw)

def _thin():
    s = Side(style="thin", color=BORDER_GREY)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

# ── Common helpers ────────────────────────────────────────────────────────────
def _title(ws, row, text, span=1):
    """Big section title — size 16 bold dark blue."""
    if span > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1, text)
    c.font = _font(size=16, bold=True, color=BLUE)
    ws.row_dimensions[row].height = 24

def _section(ws, row, text, span=1):
    """Sub-section header — size 12 bold."""
    if span > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1, text)
    c.font = _font(size=12, bold=True)
    ws.row_dimensions[row].height = 20

def _banner(ws, row, text, kind="warn", span=7):
    """Coloured alert banner row."""
    bg, fg = (AMBER_BG, AMBER_FG) if kind == "warn" else \
             (GREEN_BG, GREEN_FG) if kind == "good" else \
             (RED_BG, RED_FG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1, text)
    c.font = _font(size=12, bold=True, color=fg)
    c.fill = _f(bg)
    c.alignment = _align(h="left", wrap=True)
    ws.row_dimensions[row].height = 30

def _table_header(ws, row, headers):
    """White-on-blue table header row."""
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row, ci, h)
        c.font = _font(bold=True, color="FFFFFF")
        c.fill = _f(BLUE)
        c.alignment = _align(h="center")
        c.border = _thin()
    ws.row_dimensions[row].height = 20

def _write_row(ws, row, values, num_cols=None):
    """Write a data row with thin borders, currency format on numeric cols."""
    num_cols = num_cols or set()
    for ci, v in enumerate(values, 1):
        c = ws.cell(row, ci, v)
        c.font = _font()
        c.border = _thin()
        if ci in num_cols or isinstance(v, (int, float)) and not isinstance(v, bool):
            c.number_format = NUM_FMT
            c.alignment = _align(h="right")
            if isinstance(v, (int, float)) and v < 0:
                c.font = _font(color="9C0006")
        else:
            c.alignment = _align(h="left")

def _set_cols(ws, widths):
    """widths = {'A': 45, 'B': 22, ...}"""
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ── Sheet 1: Summary ──────────────────────────────────────────────────────────
def _summary_sheet(wb, data):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 50, "B": 22, "C": 22, "D": 22})

    _title(ws, 1, "Reconciliation Summary", span=4)

    stats = data.get("stats", {})
    bal   = data.get("balances", {})

    # Match Statistics block
    _section(ws, 3, "Match Statistics", span=4)
    stat_rows = [
        ("Total records (Our books)",                stats.get("total_ours", 0)),
        ("Total records (Their books)",              stats.get("total_theirs", 0)),
        ("L1 matches (Date + Ref + Amount)",         stats.get("l1_matches", 0)),
        ("L2 matches (Ref + Amount, dates differ)",  stats.get("l2_matches", 0)),
        ("L3 matches (Date + Amount, weak)",         stats.get("l3_matches", 0)),
        ("Amount mismatches",                         stats.get("amount_mismatches", 0)),
        ("Missing in their books",                    stats.get("missing_their_books", 0)),
        ("Missing in our books",                      stats.get("missing_our_books", 0)),
        ("TDS journal entries (see TDS Reconciliation sheet)", stats.get("tds_journal_entries", 0)),
    ]
    r = 4
    for label, val in stat_rows:
        ws.cell(r, 1, label).font = _font()
        c = ws.cell(r, 2, val)
        c.font = _font(bold=True); c.alignment = _align(h="right")
        c.number_format = INT_FMT
        r += 1
    r += 1   # blank

    # Closing Balance Walk block
    _section(ws, r, "Closing Balance Walk", span=4); r += 1
    walk_rows = [
        ("Opening Balance (Ours)",      bal.get("opening_ours", 0)),
        ("+ Sum of our transactions",   bal.get("sum_ours", 0)),
        ("= Closing Balance (Ours)",    bal.get("closing_ours", 0)),
        (None, None),
        ("Opening Balance (Theirs)",    bal.get("opening_theirs", 0)),
        ("+ Sum of their transactions", bal.get("sum_theirs", 0)),
        ("= Closing Balance (Theirs)",  bal.get("closing_theirs", 0)),
        (None, None),
        ("Difference (Ours − Theirs)",  bal.get("difference", 0)),
        ("Reconciling items (one-sided)", bal.get("reconciling_items", 0)),
    ]
    for label, val in walk_rows:
        if label is None:
            r += 1; continue
        c1 = ws.cell(r, 1, label)
        c2 = ws.cell(r, 2, val)
        # Bold the "=" and "Difference" rows
        is_total = label.startswith("=") or label.startswith("Difference")
        c1.font = _font(bold=is_total)
        c2.font = _font(bold=is_total)
        c2.alignment = _align(h="right")
        c2.number_format = NUM_FMT
        if is_total and isinstance(val, (int, float)) and val < 0:
            c2.font = _font(bold=True, color="9C0006")
        r += 1

# ── Sheet 2: TDS Reconciliation ───────────────────────────────────────────────
def _tds_sheet(wb, data):
    ws = wb.create_sheet("TDS Reconciliation")
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 50, "B": 22, "C": 22, "D": 55, "E": 18, "F": 14})

    tds = data.get("tds", {}) or {}

    _title(ws, 1, "TDS Reconciliation", span=7)

    # Banner
    status = (tds.get("status") or "CLEAR").upper()
    msg    = tds.get("message", "")
    if status == "CLEAR":
        prefix, kind = "✓ CLEAR", "good"
    elif status == "PARTIAL":
        prefix, kind = "⚠ PARTIAL", "warn"
    elif status == "EXCESS":
        prefix, kind = "✗ EXCESS", "bad"
    else:
        prefix, kind = "⚠ MISMATCH", "warn"
    banner_text = f"{prefix}  —  {msg}" if msg else prefix
    _banner(ws, 3, banner_text, kind=kind, span=7)

    # TDS Totals Comparison
    _section(ws, 6, "TDS Totals Comparison", span=7)
    ws.cell(7, 2, "Our Records").font = _font(bold=True); ws.cell(7,2).alignment = _align(h="right")
    ws.cell(7, 3, "Their Records").font = _font(bold=True); ws.cell(7,3).alignment = _align(h="right")
    rows = [
        ("TDS column total (sum of TDS Amount):",
         tds.get("our_tds_column_total", 0), tds.get("their_tds_column_total", 0)),
        ("TDS journal entries (descr-flagged):",
         tds.get("our_tds_journal_total", 0), tds.get("their_tds_journal_total", 0)),
    ]
    r = 8
    for label, ours, theirs in rows:
        ws.cell(r, 1, label).font = _font()
        for ci, v in [(2, ours), (3, theirs)]:
            c = ws.cell(r, ci, v)
            c.font = _font(); c.alignment = _align(h="right"); c.number_format = NUM_FMT
        r += 1

    # Cross-Comparison
    _section(ws, 11, "Cross-Comparison", span=7)
    cross_rows = [
        ("Our TDS column  vs  Their TDS journal:", tds.get("our_col_vs_their_journal", 0)),
        ("Their TDS column  vs  Our TDS journal:", tds.get("their_col_vs_our_journal", 0)),
    ]
    r = 12
    for label, val in cross_rows:
        ws.cell(r, 1, label).font = _font()
        c = ws.cell(r, 2, val)
        c.font = _font(); c.alignment = _align(h="right"); c.number_format = NUM_FMT
        r += 1

    # Individual TDS journal entries
    journal = tds.get("journal_entries", []) or []
    if journal:
        _section(ws, 16, "Individual TDS Journal Entries Detected", span=7)
        _table_header(ws, 17, ["Source", "Date", "Voucher No", "Description", "Amount", "Status"])
        r = 18
        for je in journal:
            _write_row(ws, r, [
                je.get("source",""), je.get("date",""), je.get("voucher_no",""),
                je.get("description",""), je.get("amount", 0), je.get("status","")
            ], num_cols={5})
            # Status badge colour
            s = (je.get("status") or "").upper()
            if s:
                sc = ws.cell(r, 6)
                if s == "CLEAR":     sc.fill = _f(GREEN_BG); sc.font = _font(bold=True, color=GREEN_FG)
                elif s == "PARTIAL": sc.fill = _f(AMBER_BG); sc.font = _font(bold=True, color=AMBER_FG)
                elif s == "EXCESS":  sc.fill = _f(RED_BG);   sc.font = _font(bold=True, color=RED_FG)
                sc.alignment = _align(h="center")
            r += 1

# ── Sheet 3: Matched ──────────────────────────────────────────────────────────
def _matched_sheet(wb, data):
    ws = wb.create_sheet("Matched")
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 12, "B": 13, "C": 21, "D": 21, "E": 13, "F": 47, "G": 50, "H": 14, "I": 14, "J": 14})

    rows = data.get("matched", [])
    headers = ["Rec ID", "Match Level", "Our Date", "Their Date", "Invoice Ref",
               "Our Description", "Their Description", "Our Amount", "Their Amount", "TDS Amount"]
    _table_header(ws, 1, headers)
    ws.freeze_panes = "A2"

    if not rows:
        ws.merge_cells("A2:J2")
        c = ws.cell(2, 1, "(No matched records found)")
        c.font = _font(italic=True, color="808080"); c.alignment = _align(h="center")
        return

    r = 2
    for row in rows:
        _write_row(ws, r, [
            row.get("rec_id",""), row.get("match_level",""),
            row.get("our_date",""), row.get("their_date",""),
            row.get("invoice_ref",""),
            row.get("our_description",""), row.get("their_description",""),
            row.get("our_amount", 0), row.get("their_amount", 0), row.get("tds_amount", 0),
        ], num_cols={8, 9, 10})

        # Colour the match level
        ml = (row.get("match_level") or "").upper()
        lc = ws.cell(r, 2)
        if ml == "L1":   lc.fill = _f(GREEN_BG); lc.font = _font(bold=True, color=GREEN_FG)
        elif ml == "L2": lc.fill = _f(AMBER_BG); lc.font = _font(bold=True, color=AMBER_FG)
        elif ml == "L3": lc.fill = _f(LIGHT_BLUE); lc.font = _font(bold=True, color=BLUE)
        lc.alignment = _align(h="center")
        r += 1

# ── Sheet 4: Amount Mismatches ────────────────────────────────────────────────
def _amount_mismatches_sheet(wb, data):
    ws = wb.create_sheet("Amount Mismatches")
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 14, "B": 13, "C": 14, "D": 45, "E": 16, "F": 16, "G": 14})

    rows = data.get("amount_mismatches", [])
    if not rows:
        ws["A1"] = "(No amount mismatches — clean!)"
        ws["A1"].font = _font(size=12, italic=True, color=GREEN_FG)
        return

    headers = ["Rec ID", "Date", "Invoice Ref", "Description",
               "Our Amount", "Their Amount", "Difference"]
    _table_header(ws, 1, headers)
    ws.freeze_panes = "A2"

    r = 2
    for row in rows:
        _write_row(ws, r, [
            row.get("rec_id",""), row.get("date",""), row.get("invoice_ref",""),
            row.get("description",""),
            row.get("our_amount", 0), row.get("their_amount", 0), row.get("difference", 0),
        ], num_cols={5, 6, 7})
        # Highlight non-zero difference
        diff = row.get("difference", 0)
        if isinstance(diff, (int, float)) and diff != 0:
            dc = ws.cell(r, 7)
            dc.fill = _f(RED_BG); dc.font = _font(bold=True, color=RED_FG)
            dc.number_format = NUM_FMT; dc.alignment = _align(h="right")
        r += 1

# ── Sheets 5 & 6: Missing in Their/Our Books ──────────────────────────────────
def _missing_sheet(wb, sheet_name, rows, empty_msg):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 13, "B": 14, "C": 16, "D": 13, "E": 36, "F": 14, "G": 14, "H": 14, "I": 40})

    if not rows:
        ws["A1"] = empty_msg
        ws["A1"].font = _font(size=12, italic=True, color=GREEN_FG)
        return

    headers = ["Date", "Voucher Type", "Voucher No", "Invoice Ref", "Description",
               "Gross Amount", "TDS Amount", "Net Amount", "Note"]
    _table_header(ws, 1, headers)
    ws.freeze_panes = "A2"

    r = 2
    for row in rows:
        _write_row(ws, r, [
            row.get("date",""), row.get("voucher_type",""), row.get("voucher_no",""),
            row.get("invoice_ref",""), row.get("description",""),
            row.get("gross_amount", 0), row.get("tds_amount", 0), row.get("net_amount", 0),
            row.get("note","")
        ], num_cols={6, 7, 8})
        r += 1

# ── Sheet 7: Timing Differences ───────────────────────────────────────────────
def _timing_sheet(wb, data):
    ws = wb.create_sheet("Timing Differences")
    ws.sheet_view.showGridLines = False
    _set_cols(ws, {"A": 12, "B": 13, "C": 21, "D": 21, "E": 13, "F": 47, "G": 50, "H": 14, "I": 14, "J": 14, "K": 12})

    rows = data.get("timing_differences", [])
    if not rows:
        ws["A1"] = "(No timing differences — all matched dates agree)"
        ws["A1"].font = _font(size=12, italic=True, color=GREEN_FG)
        return

    headers = ["Rec ID", "Match Level", "Our Date", "Their Date", "Invoice Ref",
               "Our Description", "Their Description", "Our Amount", "Their Amount",
               "TDS Amount", "Days Diff"]
    _table_header(ws, 1, headers)
    ws.freeze_panes = "A2"

    r = 2
    for row in rows:
        _write_row(ws, r, [
            row.get("rec_id",""), row.get("match_level","L2"),
            row.get("our_date",""), row.get("their_date",""),
            row.get("invoice_ref",""),
            row.get("our_description",""), row.get("their_description",""),
            row.get("our_amount", 0), row.get("their_amount", 0), row.get("tds_amount", 0),
            row.get("days_diff", 0)
        ], num_cols={8, 9, 10, 11})

        ml = (row.get("match_level") or "L2").upper()
        lc = ws.cell(r, 2)
        if ml == "L2": lc.fill = _f(AMBER_BG); lc.font = _font(bold=True, color=AMBER_FG)
        lc.alignment = _align(h="center")
        r += 1

# ── Main entry point ──────────────────────────────────────────────────────────
def reco_to_excel(reco_data: dict) -> bytes:
    """Build the 7-sheet reconciliation report from Claude's structured data."""
    wb = Workbook()
    wb.remove(wb.active)

    _summary_sheet(wb, reco_data)
    _tds_sheet(wb, reco_data)
    _matched_sheet(wb, reco_data)
    _amount_mismatches_sheet(wb, reco_data)
    _missing_sheet(wb, "Missing in Their Books",
                   reco_data.get("missing_their_books", []),
                   "(Nothing missing on their side)")
    _missing_sheet(wb, "Missing in Our Books",
                   reco_data.get("missing_our_books", []),
                   "(Nothing missing on our side)")
    _timing_sheet(wb, reco_data)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()