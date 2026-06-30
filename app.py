"""
app.py — UI. Upload → Run → Download. That's it.
"""
from __future__ import annotations
import os, traceback
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

from reco import (
    run_reconciliation,
    MODEL_NAMES, MODEL_CAPTIONS, NAME_TO_LABEL, DEFAULT_INDEX, DEFAULT_USD_INR,
)

load_dotenv()
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
USD_INR = DEFAULT_USD_INR

st.set_page_config(page_title="Ledger Reconciliation", page_icon="📘", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap');
:root{ --bg:#F3EEE7; --bg2:#FBF8F3; --card:#FFF; --ink:#272220; --muted:#8C8379;
       --accent:#B5793A; --adk:#8C5C2B; --line:#EAE3D8; }
header[data-testid="stHeader"],#MainMenu,footer,div[data-testid="stToolbar"],
section[data-testid="stSidebar"]{ display:none !important; }
.stApp{ background:radial-gradient(900px 420px at 50% -8%,var(--bg2) 0%,rgba(251,248,243,0) 70%),
  linear-gradient(180deg,var(--bg) 0%,#EFE9E1 100%) fixed; }
.block-container{ padding:2.4rem 1rem 4rem; max-width:720px; }
html,body,[class*="css"]{ font-family:'Inter',sans-serif; color:var(--ink); }
h1,h2,h3{ font-family:'Fraunces',serif !important; color:var(--ink); letter-spacing:-.015em; }
div[data-testid="stVerticalBlock"]{ gap:.85rem; }

.app-header{ text-align:center; margin-bottom:1.6rem; }
.badge{ display:inline-flex; align-items:center; gap:6px; font-size:.7rem; font-weight:600;
  letter-spacing:.14em; text-transform:uppercase; color:var(--adk);
  background:rgba(181,121,58,.08); border:1px solid rgba(181,121,58,.22);
  border-radius:999px; padding:5px 14px; margin-bottom:14px; }
.dot{ width:6px; height:6px; border-radius:50%; background:var(--accent); display:inline-block; }
.app-header h1{ font-size:2.35rem; font-weight:600; margin:0; }
.app-header p{ color:var(--muted); margin:.55rem 0 0; font-size:1rem; }

div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--card) !important; border:1px solid var(--line) !important;
  border-radius:18px; padding:1.4rem 1.5rem;
  box-shadow:0 1px 2px rgba(39,34,32,.04),0 12px 30px rgba(39,34,32,.05); }
.card-h{ display:flex; align-items:center; gap:10px; margin-bottom:.4rem; }
.num{ background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  width:26px; height:26px; border-radius:8px; display:inline-flex; align-items:center;
  justify-content:center; font-size:.82rem; font-weight:600; box-shadow:0 4px 10px rgba(181,121,58,.3); }
.card-t{ font-family:'Fraunces',serif; font-weight:600; font-size:1.12rem; }
.card-sub{ color:var(--muted); font-size:.86rem; margin:-.1rem 0 .4rem 36px; }
.summary-box{ background:var(--bg2); border:1px solid var(--line); border-radius:12px;
  padding:14px 16px; font-size:.95rem; line-height:1.7; color:var(--ink); margin-bottom:.6rem; }

div[data-testid="stFileUploader"] section{
  background:var(--bg2); border:1.5px dashed var(--line); border-radius:12px; padding:.6rem .8rem; }
div[data-testid="stFileUploader"] section:hover{ border-color:var(--accent); }

div[data-testid="stRadio"]{ width:100%; }
div[data-testid="stRadio"] > div{ width:100%; }
div[role="radiogroup"]{ display:flex; width:100%; gap:10px; flex-wrap:nowrap; }
div[role="radiogroup"] > label{ flex:1 1 0; min-width:0; min-height:64px; box-sizing:border-box;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px;
  background:var(--bg2); border:1.5px solid var(--line); border-radius:12px; padding:8px 10px;
  margin:0 !important; cursor:pointer; transition:all .15s; }
div[role="radiogroup"] > label > div:first-child{ display:none !important; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{
  margin:0; white-space:nowrap; font-size:.98rem; font-weight:600; color:var(--ink); text-align:center; }
div[role="radiogroup"] > label div[data-testid="stCaptionContainer"] p{
  margin:0; font-size:.72rem !important; font-weight:500; color:var(--muted) !important; text-align:center; }
div[role="radiogroup"] > label:hover{ border-color:var(--accent); background:#FCFAF6; }
div[role="radiogroup"] > label:has(input:checked){ border-color:var(--accent);
  background:rgba(181,121,58,.09); box-shadow:0 4px 14px rgba(181,121,58,.16); }
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p{ color:var(--adk); }

.stButton>button{ background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  border:0; border-radius:13px; padding:.85rem 1.2rem; font-weight:600; font-size:.98rem;
  box-shadow:0 10px 24px rgba(181,121,58,.28); transition:transform .08s,filter .2s; }
.stButton>button:hover:not(:disabled){ transform:translateY(-1px); filter:brightness(1.04); }
.stButton>button:disabled{ background:#D6CDC0; box-shadow:none; }
[data-testid="stDownloadButton"]>button{ background:var(--ink); color:#fff; border:0;
  border-radius:13px; padding:.9rem 1.2rem; font-weight:600; width:100%;
  box-shadow:0 10px 24px rgba(39,34,32,.18); }
[data-testid="stDownloadButton"]>button:hover{ background:#1c1815; }

.chip{ display:inline-block; background:var(--bg2); color:var(--adk); border:1px solid var(--line);
  border-radius:999px; padding:4px 12px; font-size:.8rem; font-weight:600; margin:0 6px 6px 0; }
.note{ color:var(--muted); font-size:.82rem; margin-top:.4rem; }
.stExpander{ border:1px solid var(--line) !important; border-radius:12px !important; }
</style>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────────────────────
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

def start_new():
    st.session_state.pop("result", None)
    st.session_state["uploader_key"] += 1

# ── Loader ────────────────────────────────────────────────────────────────────
def loader_html():
    return """
<div style="font-family:'Inter',sans-serif;display:flex;flex-direction:column;align-items:center;
     justify-content:center;padding:36px 16px;background:#fff;border:1px solid #EAE3D8;
     border-radius:18px;box-shadow:0 12px 30px rgba(39,34,32,.06);">
  <div class="ring"></div>
  <div id="ld" style="margin-top:20px;font-weight:600;color:#272220;font-size:1.02rem;min-height:1.4em;">Reading the ledgers…</div>
  <div style="margin-top:6px;color:#8C8379;font-size:.84rem;">Usually 30–90 seconds.</div>
  <div class="bar"><div class="fill"></div></div>
</div>
<style>
.ring{width:58px;height:58px;border-radius:50%;border:5px solid #F0E7DA;border-top-color:#B5793A;
  border-right-color:#D9A766;animation:spin 1s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
.bar{width:240px;height:6px;background:#F0E7DA;border-radius:99px;margin-top:22px;overflow:hidden;}
.fill{height:100%;width:40%;border-radius:99px;background:linear-gradient(90deg,#B5793A,#D9A766);
  animation:slide 1.6s ease-in-out infinite;}
@keyframes slide{0%{margin-left:-40%;}100%{margin-left:100%;}}
</style>
<script>
const m=["Reading the ledgers…","Understanding both books…","Matching invoices…",
"Checking TDS differences…","Spotting unmatched entries…","Building the report…","Almost done…"];
let i=0;const e=document.getElementById('ld');
setInterval(()=>{i=(i+1)%m.length;if(e)e.textContent=m[i];},2600);
</script>"""

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <span class="badge"><span class="dot"></span>AI-Powered</span>
  <h1>Ledger Reconciliation</h1>
  <p>Upload two ledgers — Claude reconciles and returns a ready Excel report.</p>
</div>""", unsafe_allow_html=True)

# ── Results view ──────────────────────────────────────────────────────────────
import json as _json

def _reco_to_sheets(reco_data):
    """Convert reco JSON to the list-of-sheets format used by the dashboard."""
    sheets = []
    s = reco_data.get("stats", {}) or {}
    b = reco_data.get("balances", {}) or {}
    tds = reco_data.get("tds", {}) or {}

    # Summary as a key/value list
    summary_rows = [["Metric", "Value"]]
    summary_rows += [
        ["Total records (Our books)",      s.get("total_ours", 0)],
        ["Total records (Their books)",    s.get("total_theirs", 0)],
        ["L1 matches",                     s.get("l1_matches", 0)],
        ["L2 matches",                     s.get("l2_matches", 0)],
        ["L3 matches",                     s.get("l3_matches", 0)],
        ["Amount mismatches",              s.get("amount_mismatches", 0)],
        ["Missing in their books",         s.get("missing_their_books", 0)],
        ["Missing in our books",           s.get("missing_our_books", 0)],
        ["Closing Balance (Ours)",         b.get("closing_ours", 0)],
        ["Closing Balance (Theirs)",       b.get("closing_theirs", 0)],
        ["Difference (Ours − Theirs)", b.get("difference", 0)],
    ]
    sheets.append({"name": "Reconciliation Summary", "rows": summary_rows})

    # TDS
    if tds.get("status") or tds.get("journal_entries"):
        tds_rows = [["Item", "Our Records", "Their Records"]]
        tds_rows += [
            ["TDS column total",   tds.get("our_tds_column_total", 0),  tds.get("their_tds_column_total", 0)],
            ["TDS journal total",  tds.get("our_tds_journal_total", 0), tds.get("their_tds_journal_total", 0)],
        ]
        for je in tds.get("journal_entries", []) or []:
            tds_rows.append([
                f"{je.get('source','')} {je.get('voucher_no','')}",
                je.get("amount", 0) if je.get("source") == "Ours" else "",
                je.get("amount", 0) if je.get("source") == "Theirs" else "",
            ])
        sheets.append({"name": "TDS Reconciliation", "rows": tds_rows,
                       "_tds_status": tds.get("status",""), "_tds_msg": tds.get("message","")})

    # Matched
    matched = reco_data.get("matched", []) or []
    if matched:
        rows = [["Rec ID", "Level", "Our Date", "Their Date", "Invoice", "Description", "Our Amount", "Their Amount"]]
        for m in matched:
            rows.append([
                m.get("rec_id",""), m.get("match_level",""),
                m.get("our_date",""), m.get("their_date",""),
                m.get("invoice_ref",""), m.get("our_description","") or m.get("their_description",""),
                m.get("our_amount", 0), m.get("their_amount", 0)
            ])
        sheets.append({"name": "Matched", "rows": rows})

    # Amount mismatches
    am = reco_data.get("amount_mismatches", []) or []
    if am:
        rows = [["Rec ID", "Date", "Invoice", "Description", "Our Amount", "Their Amount", "Difference"]]
        for m in am:
            rows.append([m.get("rec_id",""), m.get("date",""), m.get("invoice_ref",""),
                         m.get("description",""), m.get("our_amount", 0),
                         m.get("their_amount", 0), m.get("difference", 0)])
        sheets.append({"name": "Amount Mismatches", "rows": rows})

    # Missing in their books
    mt = reco_data.get("missing_their_books", []) or []
    if mt:
        rows = [["Date", "Voucher", "Invoice", "Description", "Gross", "TDS", "Net"]]
        for x in mt:
            rows.append([x.get("date",""), x.get("voucher_no",""), x.get("invoice_ref",""),
                         x.get("description",""), x.get("gross_amount", 0),
                         x.get("tds_amount", 0), x.get("net_amount", 0)])
        sheets.append({"name": "Missing in Their Books", "rows": rows})

    # Missing in our books
    mo = reco_data.get("missing_our_books", []) or []
    if mo:
        rows = [["Date", "Voucher", "Invoice", "Description", "Gross", "TDS", "Net"]]
        for x in mo:
            rows.append([x.get("date",""), x.get("voucher_no",""), x.get("invoice_ref",""),
                         x.get("description",""), x.get("gross_amount", 0),
                         x.get("tds_amount", 0), x.get("net_amount", 0)])
        sheets.append({"name": "Missing in Our Books", "rows": rows})

    # Timing differences
    td = reco_data.get("timing_differences", []) or []
    if td:
        rows = [["Rec ID", "Our Date", "Their Date", "Days", "Invoice", "Description", "Amount"]]
        for x in td:
            rows.append([x.get("rec_id",""), x.get("our_date",""), x.get("their_date",""),
                         x.get("days_diff", 0), x.get("invoice_ref",""),
                         x.get("our_description",""), x.get("our_amount", 0)])
        sheets.append({"name": "Timing Differences", "rows": rows})

    return sheets

def _detect_tone(name):
    n = name.lower()
    if any(k in n for k in ("summary","overview","reconciliation summary")): return "summary"
    if any(k in n for k in ("matched","match","agree","payment","invoice","credit","bank","receipt")): return "matched"
    if any(k in n for k in ("diff","mismatch","issue","varianc","key diff","discrepan","amount diff")): return "diff"
    if any(k in n for k in ("tds","tax","withhold")): return "tds"
    if any(k in n for k in ("only in a","only a","unmatched a")): return "onlya"
    if any(k in n for k in ("only in b","only b","unmatched b")): return "onlyb"
    if any(k in n for k in ("conclusion","auditor","finding")): return "conc"
    return "other"

TONE_STYLE = {
    "summary": dict(icon="SUMM", ibg="#F0FDF4", ifg="#16A34A", accent="#16A34A", alert=None),
    "matched": dict(icon="OK",   ibg="#F0FDF4", ifg="#16A34A", accent="#16A34A", alert="success"),
    "diff":    dict(icon="DIFF", ibg="#FEF2F2", ifg="#DC2626", accent="#DC2626", alert="danger"),
    "tds":     dict(icon="TDS",  ibg="#FFFBEB", ifg="#D97706", accent="#D97706", alert="warn"),
    "onlya":   dict(icon="A",    ibg="#F5F3FF", ifg="#7C3AED", accent="#7C3AED", alert="warn"),
    "onlyb":   dict(icon="B",    ibg="#EFF6FF", ifg="#2563EB", accent="#2563EB", alert="warn"),
    "conc":    dict(icon="CONC", ibg="#EFF6FF", ifg="#2563EB", accent="#2563EB", alert=None),
    "other":   dict(icon="INFO", ibg="#F9FAFB", ifg="#6B7280", accent="#6B7280", alert=None),
}
STATUS_WORDS = {"MATCHED","CLEAR","SUBSTANTIALLY RECONCILED","UNRESOLVED","HIGH","MINOR","MEDIUM","PARTIAL","ACCEPTABLE","LOW"}

def _extract_metrics(sheets):
    metrics = []
    for sh in sheets:
        if "summary" in sh.get("name","").lower():
            for row in sh.get("rows",[])[1:]:
                if len(row) >= 2:
                    label = str(row[0]).lower()
                    val   = row[-1] if len(row)>2 else row[1]
                    if any(k in label for k in ("closing diff","variance","difference","closing balance diff")):
                        try:
                            metrics.append(dict(label="CLOSING DIFFERENCE", value=float(str(val).replace(",","").replace("₹","").strip()), sub="Ledger A vs Ledger B", kind="diff"))
                        except: pass
    matched = sum(max(0,len(sh.get("rows",[]))-1) for sh in sheets if _detect_tone(sh.get("name",""))=="matched")
    if matched:
        metrics.append(dict(label="MATCHED ENTRIES", value=matched, sub="Invoices, CNs, payments", kind="good"))
    diff_rows = sum(max(0,len(sh.get("rows",[]))-1) for sh in sheets if _detect_tone(sh.get("name",""))=="diff")
    if diff_rows:
        metrics.append(dict(label="UNRESOLVED ITEMS", value=diff_rows, sub="Need investigation", kind="warn"))
    return metrics[:4]

def _badge(v):
    sv = str(v).strip().upper()
    if sv in ("MATCHED","CLEAR","SUBSTANTIALLY RECONCILED"):
        return '<span style="display:inline-flex;align-items:center;gap:3px;background:#DCFCE7;color:#166534;font-size:11px;font-weight:600;padding:3px 10px;border-radius:999px;white-space:nowrap">&#10003; ' + str(v) + '</span>'
    if sv in ("UNRESOLVED","HIGH"):
        return '<span style="display:inline-flex;align-items:center;gap:3px;background:#FEE2E2;color:#991B1B;font-size:11px;font-weight:600;padding:3px 10px;border-radius:999px;white-space:nowrap">&#10005; ' + str(v) + '</span>'
    if sv in ("MINOR","MEDIUM","PARTIAL","ACCEPTABLE","LOW"):
        return '<span style="display:inline-flex;align-items:center;gap:3px;background:#FEF3C7;color:#92400E;font-size:11px;font-weight:600;padding:3px 10px;border-radius:999px;white-space:nowrap">~ ' + str(v) + '</span>'
    return None

def _fmt_cell(v):
    sv = str(v).strip().upper()
    b = _badge(v)
    if b: return b
    if isinstance(v,(int,float)) and not isinstance(v,bool):
        color = "#DC2626" if v<0 else ("#9CA3AF" if v==0 else "#111827")
        w = "600" if v!=0 else "400"
        fv = "₹{:,.2f}".format(v) if isinstance(v,float) else "{:,}".format(v)
        return '<span style="color:' + color + ';font-weight:' + w + ';font-variant-numeric:tabular-nums">' + fv + '</span>'
    if v is None or str(v).strip()=="":
        return '<span style="color:#D1D5DB">&mdash;</span>'
    return str(v)

def _is_num_col(ci, data):
    for row in data[:5]:
        if ci<len(row) and isinstance(row[ci],(int,float)) and not isinstance(row[ci],bool): return True
    return False

def _build_html(sheets, model_label, input_tokens, output_tokens, cost_inr, cost_usd, metrics):
    # metric cards
    mc_html = ""
    for m in metrics:
        color = "#DC2626" if m["kind"]=="diff" else ("#16A34A" if m["kind"]=="good" else "#D97706")
        val = m["value"]
        if m["kind"]=="diff":
            fmt_val = "₹{:,.0f}".format(abs(val))
        elif isinstance(val, float) and val==int(val):
            fmt_val = str(int(val))
        else:
            fmt_val = str(val)
        mc_html += (
            '<div style="flex:1;min-width:140px">'
            '<div style="font-size:10px;font-weight:700;letter-spacing:.1em;color:#9CA3AF;margin-bottom:6px">' + m["label"] + '</div>'
            '<div style="font-size:28px;font-weight:700;line-height:1;color:' + color + '">' + fmt_val + '</div>'
            '<div style="font-size:12px;color:#9CA3AF;margin-top:5px">' + m["sub"] + '</div>'
            '</div>'
        )

    pills = (
        '<span style="font-size:11px;font-weight:600;background:#F3F4F6;border:1px solid #E5E7EB;border-radius:999px;padding:3px 11px;color:#6B7280">' + model_label.split(" (")[0] + '</span> '
        '<span style="font-size:11px;font-weight:600;background:#F3F4F6;border:1px solid #E5E7EB;border-radius:999px;padding:3px 11px;color:#6B7280">In {:,} &middot; Out {:,} tokens</span> '.format(input_tokens, output_tokens) +
        '<span style="font-size:11px;font-weight:600;background:#F3F4F6;border:1px solid #E5E7EB;border-radius:999px;padding:3px 11px;color:#6B7280">₹ {:,.2f} ($ {:.4f})</span>'.format(cost_inr, cost_usd)
    )

    ALERT_CONFIGS = {
        "danger":  dict(bg="#FEF2F2",  border="#FECACA", lborder="#DC2626", icon="&#9888;", title="Differences found",           body="These items need investigation."),
        "warn":    dict(bg="#FFFBEB",  border="#FDE68A", lborder="#D97706", icon="!",        title="Attention required",           body="Verify these entries with the counterparty."),
        "success": dict(bg="#F0FDF4",  border="#BBF7D0", lborder="#16A34A", icon="&#10003;", title="All entries reconciled",       body="Amounts agree on both sides."),
    }

    sections_html = ""
    for idx, sh in enumerate(sheets):
        rows   = sh.get("rows", [])
        n_data = max(0, len(rows)-1)
        if n_data == 0: continue
        tone = _detect_tone(sh["name"])
        ts   = TONE_STYLE[tone]
        headers = [str(h) for h in rows[0]] if rows else []
        data    = rows[1:] if len(rows)>1 else []

        # alert banner
        alert_html = ""
        ak = ts["alert"]
        if ak and ak in ALERT_CONFIGS:
            c = dict(ALERT_CONFIGS[ak])
            if tone in ("onlya","onlyb"):
                side = "Party A" if tone=="onlya" else "Party B"
                c["title"] = "Present only in " + side + "'s books"
                c["body"]  = "No matching entry in the other ledger."
            elif tone == "tds":
                c["title"] = "TDS / withholding tax"
                c["body"]  = "Verify TDS amounts are correctly booked on both sides."
            elif tone == "matched":
                c["title"] = "All entries reconciled"
                c["body"]  = "Amounts agree on both sides including TDS base and TDS amounts."
            alert_html = (
                '<div style="display:flex;gap:12px;align-items:flex-start;padding:12px 16px;border-radius:8px;margin-bottom:12px;'
                'border-left:4px solid ' + c["lborder"] + ';background:' + c["bg"] + ';'
                'border-top:1px solid ' + c["border"] + ';border-right:1px solid ' + c["border"] + ';border-bottom:1px solid ' + c["border"] + '">'
                '<span style="font-size:16px;color:' + c["lborder"] + ';flex-shrink:0;font-weight:700">' + c["icon"] + '</span>'
                '<div><div style="font-size:13px;font-weight:600;color:#111827;margin-bottom:2px">' + c["title"] + '</div>'
                '<div style="font-size:12px;color:#6B7280">' + c["body"] + '</div></div></div>'
            )

        # table header row
        th_row = ""
        for ci, h in enumerate(headers):
            align = "right" if _is_num_col(ci, data) else "left"
            th_row += (
                '<th style="padding:10px 14px;text-align:' + align + ';font-size:11px;font-weight:700;'
                'letter-spacing:.06em;text-transform:uppercase;color:#6B7280;background:#F9FAFB;'
                'border-bottom:2px solid ' + ts["accent"] + ';white-space:nowrap">' + h + '</th>'
            )

        # table data rows
        tr_rows = ""
        for ri, row in enumerate(data):
            pad = list(row) + [""] * max(0, len(headers)-len(row))
            bg  = "#FFFFFF" if ri%2==0 else "#F9FAFB"
            tds = ""
            for ci, v in enumerate(pad[:len(headers)]):
                align = "right" if _is_num_col(ci, data) else "left"
                tds += (
                    '<td style="padding:10px 14px;text-align:' + align + ';vertical-align:middle;'
                    'background:' + bg + ';border-bottom:1px solid #F3F4F6;font-size:13px">'
                    + _fmt_cell(v) + '</td>'
                )
            tr_rows += "<tr>" + tds + "</tr>"

        table_html = (
            '<div style="overflow-x:auto;border:1px solid #E5E7EB;border-radius:10px">'
            '<table style="width:100%;border-collapse:collapse;min-width:360px">'
            '<thead><tr>' + th_row + '</tr></thead><tbody>' + tr_rows + '</tbody></table></div>'
        )

        open_state = tone in ("diff","tds","onlya","onlyb")
        sid = "sec_" + str(idx)
        icon_box = (
            '<div style="width:32px;height:32px;border-radius:8px;background:' + ts["ibg"] + ';'
            'display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;'
            'color:' + ts["ifg"] + ';flex-shrink:0;letter-spacing:.02em">' + ts["icon"] + '</div>'
        )
        sections_html += (
            '<div style="background:#FFFFFF;border:1px solid #E5E7EB;border-radius:14px;overflow:hidden;margin-bottom:10px">'
            '<div onclick="var p=this.parentNode;var b=p.querySelector(\'.acc-body\');var c=this.querySelector(\'.acc-chev\');if(b.style.display===\'none\'){b.style.display=\'block\';c.innerHTML=\'&#8744;\';}else{b.style.display=\'none\';c.innerHTML=\'&#8250;\';}" style="display:flex;align-items:center;gap:12px;padding:14px 18px;cursor:pointer;user-select:none;background:#FFFFFF">'
            + icon_box +
            '<span style="font-size:14px;font-weight:600;color:#111827;flex:1">' + sh["name"] + '</span>'
            '<span style="font-size:12px;color:#9CA3AF;margin-right:10px">' + str(n_data) + ' ' + ("entry" if n_data==1 else "entries") + '</span>'
            '<span class="acc-chev" style="font-size:22px;color:#9CA3AF;font-weight:300;flex-shrink:0">' + ('&#8744;' if open_state else '&#8250;') + '</span>'
            '</div>'
            '<div class="acc-body" style="display:' + ('block' if open_state else 'none') + ';padding:0 18px 16px">'
            + alert_html + table_html +
            '</div></div>'
        )

    return (
        '<div style="font-family:-apple-system,BlinkMacSystemFont,Inter,sans-serif;color:#111827;padding:4px 0 8px">'
        '<div style="display:flex;gap:32px;flex-wrap:wrap;padding:20px 24px;background:#FFFFFF;border:1px solid #E5E7EB;border-radius:14px;margin-bottom:14px">'
        + mc_html +
        '</div>'
        '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px">' + pills + '</div>'
        + sections_html +
        '</div>'
        ''
    )

if "result" in st.session_state:
    r = st.session_state["result"]
    sheets_view = _reco_to_sheets(r.sheets) if isinstance(r.sheets, dict) else r.sheets
    metrics = _extract_metrics(sheets_view)

    if r.summary:
        st.markdown(
            '<p style="color:#4B5563;font-size:13.5px;line-height:1.75;margin-bottom:.6rem">' + r.summary + '</p>',
            unsafe_allow_html=True)

    n_rows_total = sum(max(0,len(sh.get("rows",[]))-1) for sh in sheets_view)
    h = max(700, 260 + n_rows_total * 42)
    dash = _build_html(sheets_view, r.model_label, r.input_tokens, r.output_tokens, r.cost_inr, r.cost_usd, metrics)
    components.html(dash, height=min(h, 3200), scrolling=True)

    fname = f"Reco_{datetime.now():%Y%m%d_%H%M}.xlsx"
    st.download_button(
        "\u2b07\ufe0f  Download full reconciliation report (.xlsx)",
        data=r.excel_bytes, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

    with st.expander("\U0001f50d Raw Claude response"):
        st.text(r.raw_response[:6000] + ("\u2026" if len(r.raw_response) > 6000 else ""))

    st.button("\u21ba  Start a new reconciliation", on_click=start_new, use_container_width=True)
    st.stop()

# ── Form view ─────────────────────────────────────────────────────────────────
uk = st.session_state["uploader_key"]

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">1</span>'
                '<span class="card-t">Upload the two ledgers</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-sub">Both parties\' books — Excel or CSV</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        file_a = st.file_uploader("Ledger A", type=["xls","xlsx","xlsm","csv"], key=f"a{uk}")
    with c2:
        file_b = st.file_uploader("Ledger B", type=["xls","xlsx","xlsm","csv"], key=f"b{uk}")

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">2</span>'
                '<span class="card-t">Choose model & run</span></div>', unsafe_allow_html=True)
    choice = st.radio("Model", MODEL_NAMES, captions=MODEL_CAPTIONS,
                      index=DEFAULT_INDEX, horizontal=True, label_visibility="collapsed")
    model_label = NAME_TO_LABEL[choice]
    run_btn = st.button("✨  Run Reconciliation", type="primary",
                        disabled=not (file_a and file_b and API_KEY),
                        use_container_width=True)
    if not API_KEY:
        st.markdown('<div class="note">⚠️ Set <b>ANTHROPIC_API_KEY</b> in Streamlit secrets.</div>',
                    unsafe_allow_html=True)

# ── Run ───────────────────────────────────────────────────────────────────────
if run_btn:
    slot = st.empty()
    with slot: components.html(loader_html(), height=270)
    try:
        result = run_reconciliation(
            file_a.getvalue(), file_a.name,
            file_b.getvalue(), file_b.name,
            model_label=model_label, usd_inr=USD_INR, api_key=API_KEY,
        )
        st.session_state["result"] = result
        slot.empty()
        st.rerun()
    except Exception as e:
        slot.empty()
        st.error(f"Something went wrong: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())