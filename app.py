"""
app.py — Reconciliation dashboard.
Upload two ledgers → Claude reconciles → dashboard + neat Excel download.
"""
from __future__ import annotations
import os, traceback
from datetime import datetime

import pandas as pd
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

st.set_page_config(page_title="Ledger Reconciliation", page_icon="📘", layout="wide")

# ── Tone palette (shared with the Excel writer's intent) ───────────────────────
TONES = {
    "success": ("#E7F5EE", "#A9D9BF", "#1E7D4F"),
    "warning": ("#FBF1DC", "#EAD3A0", "#8A5A0F"),
    "danger":  ("#F9E4E3", "#E7B7B4", "#B3261E"),
    "info":    ("#E4F0F8", "#AECEE4", "#1D6FA4"),
    "neutral": ("#F1EEE9", "#D9D2C7", "#55504A"),
}
def tone(name):
    return TONES.get(str(name or "").lower(), TONES["neutral"])

st.markdown("""
<style>
:root{ --bg:#F5F6F8; --bg2:#FBFCFD; --card:#FFF; --ink:#1F2937; --muted:#6B7280;
       --accent:#2563EB; --adk:#1D4ED8; --line:#E5E7EB; }
header[data-testid="stHeader"],#MainMenu,footer,div[data-testid="stToolbar"]{ display:none !important; }
.stApp{ background:linear-gradient(180deg,#F7F8FA 0%,#EEF1F5 100%) fixed; }
.block-container{ padding:2.2rem 2.2rem 4rem; max-width:1180px; }
html,body,[class*="css"]{
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  color:var(--ink); }
h1,h2,h3{ font-family:inherit !important; color:var(--ink); letter-spacing:-.01em; }

.app-header{ text-align:center; margin:.4rem 0 1.6rem; }
.app-header h1{ font-size:2rem; font-weight:700; margin:0; }
.app-header p{ color:var(--muted); margin:.5rem 0 0; font-size:1rem; }

div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--card) !important; border:1px solid var(--line) !important;
  border-radius:18px; padding:1.3rem 1.5rem;
  box-shadow:0 1px 2px rgba(39,34,32,.04),0 12px 30px rgba(39,34,32,.05); }
.card-h{ display:flex; align-items:center; gap:10px; margin-bottom:.3rem; }
.num{ background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  width:26px; height:26px; border-radius:8px; display:inline-flex; align-items:center;
  justify-content:center; font-size:.82rem; font-weight:700; box-shadow:0 4px 10px rgba(37,99,235,.3); }
.card-t{ font-weight:700; font-size:1.1rem; }
.card-sub{ color:var(--muted); font-size:.86rem; margin:-.1rem 0 .5rem 36px; }

/* Model choice — pill buttons */
div[role="radiogroup"]{ display:flex; width:100%; gap:10px; flex-wrap:nowrap; }
div[role="radiogroup"] > label{ flex:1 1 0; min-width:0; min-height:60px; box-sizing:border-box;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:2px;
  background:var(--bg2); border:1.5px solid var(--line); border-radius:12px; padding:8px 10px;
  margin:0 !important; cursor:pointer; transition:all .15s; }
div[role="radiogroup"] > label > div:first-child{ display:none !important; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{
  margin:0; white-space:nowrap; font-size:.98rem; font-weight:600; color:var(--ink); text-align:center; }
div[role="radiogroup"] > label div[data-testid="stCaptionContainer"] p{
  margin:0; font-size:.72rem !important; font-weight:500; color:var(--muted) !important; text-align:center; }
div[role="radiogroup"] > label:hover{ border-color:var(--accent); background:#F5F8FF; }
div[role="radiogroup"] > label:has(input:checked){ border-color:var(--accent);
  background:rgba(37,99,235,.08); box-shadow:0 4px 14px rgba(37,99,235,.16); }
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p{ color:var(--adk); }

/* KPI cards */
.kpi-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin:.2rem 0 1.1rem; }
.kpi{ background:#fff; border:1px solid var(--line); border-radius:16px; padding:16px 18px;
  box-shadow:0 1px 2px rgba(39,34,32,.04),0 10px 26px rgba(39,34,32,.05); position:relative; overflow:hidden; }
.kpi .accent{ position:absolute; left:0; top:0; bottom:0; width:5px; }
.kpi .label{ font-size:.72rem; font-weight:700; letter-spacing:.09em; text-transform:uppercase; color:var(--muted); }
.kpi .value{ font-size:1.7rem; font-weight:700; margin-top:6px; line-height:1.1; font-variant-numeric:tabular-nums; }

.summary-box{ background:#fff; border:1px solid var(--line); border-radius:16px;
  padding:1.2rem 1.4rem; font-size:1rem; line-height:1.75; color:var(--ink); margin-bottom:1rem;
  box-shadow:0 1px 2px rgba(39,34,32,.04),0 10px 26px rgba(39,34,32,.05); }
.sec-note{ border-radius:12px; padding:11px 15px; font-size:.92rem; font-weight:600; margin:.2rem 0 .8rem; }

.stButton>button{ background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  border:0; border-radius:12px; padding:.85rem 1.2rem; font-weight:600; font-size:.98rem;
  box-shadow:0 8px 20px rgba(37,99,235,.28); transition:transform .08s,filter .2s; }
.stButton>button:hover:not(:disabled){ transform:translateY(-1px); filter:brightness(1.05); }
.stButton>button:disabled{ background:#CBD2DB; box-shadow:none; }
[data-testid="stDownloadButton"]>button{ background:var(--ink); color:#fff; border:0;
  border-radius:12px; padding:.9rem 1.2rem; font-weight:600; width:100%;
  box-shadow:0 8px 20px rgba(31,41,55,.18); }
[data-testid="stDownloadButton"]>button:hover{ background:#111827; }
div[data-testid="stFileUploader"] section{
  background:var(--bg2); border:1.5px dashed var(--line); border-radius:12px; padding:.6rem .8rem; }
div[data-testid="stFileUploader"] section:hover{ border-color:var(--accent); }
.stTabs [data-baseweb="tab-list"]{ gap:6px; flex-wrap:wrap; }
.stTabs [data-baseweb="tab"]{ background:var(--bg2); border:1px solid var(--line);
  border-radius:10px; padding:6px 14px; font-weight:600; }
.stTabs [aria-selected="true"]{ background:rgba(37,99,235,.10); border-color:var(--accent); color:var(--adk); }
.note{ color:var(--muted); font-size:.82rem; margin-top:.4rem; }
.pill{ font-size:11px; font-weight:600; background:#EEF2FF; border:1px solid #DBE3F5;
  border-radius:999px; padding:3px 11px; color:#3B4A63; }
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
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;
     display:flex;flex-direction:column;align-items:center;
     justify-content:center;padding:34px 16px;background:#fff;border:1px solid #E5E7EB;
     border-radius:18px;box-shadow:0 12px 30px rgba(31,41,55,.06);">
  <div class="ring"></div>
  <div id="ld" style="margin-top:18px;font-weight:600;color:#1F2937;font-size:1.02rem;min-height:1.4em;">Reading the ledgers…</div>
  <div style="margin-top:6px;color:#6B7280;font-size:.84rem;">Usually 30–90 seconds.</div>
  <div class="bar"><div class="fill"></div></div>
</div>
<style>
.ring{width:54px;height:54px;border-radius:50%;border:5px solid #E5EAF2;border-top-color:#2563EB;
  border-right-color:#60A5FA;animation:spin 1s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
.bar{width:240px;height:6px;background:#E5EAF2;border-radius:99px;margin-top:20px;overflow:hidden;}
.fill{height:100%;width:40%;border-radius:99px;background:linear-gradient(90deg,#2563EB,#60A5FA);
  animation:slide 1.6s ease-in-out infinite;}
@keyframes slide{0%{margin-left:-40%;}100%{margin-left:100%;}}
</style>
<script>
const m=["Reading the ledgers…","Understanding both books…","Matching entries…",
"Reconciling balances…","Sorting the findings…","Building your report…","Almost done…"];
let i=0;const e=document.getElementById('ld');
setInterval(()=>{i=(i+1)%m.length;if(e)e.textContent=m[i];},2500);
</script>"""

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>AI Powered Ledger Reconciliation</h1>
  <p>Upload 2 ledgers and let the model reconcile them for you.</p>
</div>""", unsafe_allow_html=True)

# ── Results dashboard ─────────────────────────────────────────────────────────
def render_kpis(metrics):
    if not metrics:
        return
    cards = ""
    for m in metrics:
        _, _, fg = tone(m.get("tone"))
        cards += (
            '<div class="kpi"><div class="accent" style="background:' + fg + '"></div>'
            '<div class="label">' + str(m.get("label", "")) + '</div>'
            '<div class="value" style="color:' + fg + '">' + str(m.get("value", "")) + '</div></div>'
        )
    st.markdown('<div class="kpi-grid">' + cards + '</div>', unsafe_allow_html=True)

def render_section(sec):
    columns = sec.get("columns", []) or []
    rows    = sec.get("rows", []) or []
    note    = sec.get("note")
    bg, bd, fg = tone(sec.get("tone"))

    if note:
        st.markdown(
            f'<div class="sec-note" style="background:{bg};border:1px solid {bd};color:{fg}">{note}</div>',
            unsafe_allow_html=True)

    if not rows:
        st.caption("No entries in this section.")
        return

    # Size to the widest of (declared columns, any row) so a mismatch never crashes.
    ncol = max([len(columns)] + [len(r) for r in rows])
    cols = [str(c) for c in columns] + [f"Col {i + 1}" for i in range(len(columns), ncol)]

    # De-duplicate headers (Arrow/st.dataframe rejects duplicate column names).
    seen, cols_unique = {}, []
    for c in cols:
        seen[c] = seen.get(c, 0) + 1
        cols_unique.append(c if seen[c] == 1 else f"{c} ({seen[c]})")

    norm = [list(r) + [""] * (ncol - len(r)) for r in rows]
    df = pd.DataFrame(norm, columns=cols_unique)
    st.dataframe(df, use_container_width=True, hide_index=True)

if "result" in st.session_state:
    r = st.session_state["result"]

    top = st.columns([3, 1])
    with top[0]:
        st.markdown(
            f'<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:.4rem">'
            f'<span class="pill">{r.model_label.split(" (")[0]}</span>'
            f'<span class="pill">In {r.input_tokens:,} · Out {r.output_tokens:,} tokens</span>'
            f'<span class="pill">₹ {r.cost_inr:,.2f} ($ {r.cost_usd:.4f})</span>'
            f'</div>', unsafe_allow_html=True)
    with top[1]:
        fname = f"Reconciliation_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.download_button(
            "⬇️  Download Excel", data=r.excel_bytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    render_kpis(r.metrics)

    if r.summary:
        st.markdown(f'<div class="summary-box">{r.summary}</div>', unsafe_allow_html=True)

    if r.sections:
        titles = [s.get("title", f"Section {i+1}") for i, s in enumerate(r.sections)]
        for tab, sec in zip(st.tabs(titles), r.sections):
            with tab:
                render_section(sec)
    elif not r.summary:
        st.warning("Claude did not return structured sections. Raw response below.")
        st.code(r.raw_response[:6000])

    with st.expander("🔍  Raw AI response"):
        st.code(r.raw_response[:8000] + ("…" if len(r.raw_response) > 8000 else ""))

    st.button("↺  Start a new reconciliation", on_click=start_new)
    st.stop()

# ── Form view ─────────────────────────────────────────────────────────────────
uk = st.session_state["uploader_key"]
_, mid, _ = st.columns([1, 2.4, 1])

with mid:
    with st.container(border=True):
        st.markdown('<div class="card-t">Upload the two ledgers</div>'
                    '<div class="card-sub" style="margin-left:0">Both parties\' books — Excel, CSV or PDF</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            file_a = st.file_uploader("Ledger 1", type=["xls","xlsx","xlsm","csv","pdf"], key=f"a{uk}")
        with c2:
            file_b = st.file_uploader("Ledger 2", type=["xls","xlsx","xlsm","csv","pdf"], key=f"b{uk}")

    with st.container(border=True):
        st.markdown('<div class="card-t">Choose model</div>', unsafe_allow_html=True)
        choice = st.radio("Model", MODEL_NAMES, captions=MODEL_CAPTIONS,
                          index=DEFAULT_INDEX, horizontal=True, label_visibility="collapsed")
        model_label = NAME_TO_LABEL[choice]

    run_btn = st.button("Run Reconciliation", type="primary",
                        disabled=not (file_a and file_b and API_KEY),
                        use_container_width=True)
    if not API_KEY:
        st.markdown('<div class="note">⚠️ Set <b>ANTHROPIC_API_KEY</b> in your .env or Streamlit secrets.</div>',
                    unsafe_allow_html=True)

# ── Run ───────────────────────────────────────────────────────────────────────
if run_btn:
    slot = st.empty()
    with slot:
        components.html(loader_html(), height=260)
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
