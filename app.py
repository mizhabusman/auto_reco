"""
app.py — The platform. Upload two ledgers, pick a model, click Run.

Claude reconciles them like a CA in chat — the report streams onto the
screen live as it's written — then a downloadable Excel follows.
All logic lives in Claude (see prompts.py); this file is just the surface.
"""
from __future__ import annotations

import os
import traceback
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv

from engine import (
    DEFAULT_INDEX, MODEL_NAMES, RecoError, RecoSession,
)

UPLOAD_TYPES = ["xls", "xlsx", "xlsm", "csv", "tsv", "txt", "pdf", "docx"]

load_dotenv()
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

RESULT_VERSION = 3  # bump when the result shape changes → old session results are discarded

st.set_page_config(page_title="Ledger Reconciliation", page_icon="📘", layout="centered")

st.markdown("""
<style>
:root{ --bg:#F5F5F5; --bg2:#FAFAFA; --card:#FFF; --ink:#272220; --muted:#8C8379;
       --accent:#B5793A; --adk:#8C5C2B; --line:#EAE3D8;
       --font:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; }
header[data-testid="stHeader"],#MainMenu,footer,div[data-testid="stToolbar"],
section[data-testid="stSidebar"]{ display:none !important; }
.stApp{ background:var(--bg); }
.block-container{ padding:2.4rem 1rem 4rem; max-width:780px; }
html,body,[class*="css"]{ font-family:var(--font); color:var(--ink); }
h1,h2,h3{ font-family:var(--font) !important; color:var(--ink); letter-spacing:-.01em; }
div[data-testid="stVerticalBlock"]{ gap:.85rem; }

.app-header{ text-align:center; margin-bottom:1.6rem; }
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
.card-t{ font-weight:600; font-size:1.12rem; }
.card-sub{ color:var(--muted); font-size:.86rem; margin:-.1rem 0 .4rem 36px; }

div[data-testid="stFileUploader"] section{
  background:var(--bg2); border:1.5px dashed var(--line); border-radius:12px; padding:.6rem .8rem; }
div[data-testid="stFileUploader"] section:hover{ border-color:var(--accent); }

/* Model picker — compact segmented pill toggle (centered, not full width) */
div[role="radiogroup"]{ display:flex; width:100%; max-width:340px; margin:.15rem auto 0;
  gap:4px; flex-wrap:nowrap; align-items:stretch;
  background:linear-gradient(135deg,var(--accent),var(--adk)); border-radius:999px; padding:4px;
  box-shadow:0 6px 16px rgba(181,121,58,.24), inset 0 1px 2px rgba(255,255,255,.18); }
div[role="radiogroup"] > label{ flex:1 1 0; min-width:0;
  display:flex; align-items:center; justify-content:center;
  background:transparent; border:0; border-radius:999px; padding:6px 12px; margin:0 !important;
  cursor:pointer; transition:background-color .2s, box-shadow .2s, transform .1s; }
div[role="radiogroup"] > label > div:first-child{ display:none !important; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{
  margin:0; white-space:nowrap; font-size:.86rem; font-weight:600; text-align:center;
  color:rgba(255,255,255,.94); transition:color .2s; }
div[role="radiogroup"] > label:hover:not(:has(input:checked)){ background:rgba(255,255,255,.14); }
div[role="radiogroup"] > label:active{ transform:scale(.96); }
div[role="radiogroup"] > label:has(input:checked){ background:#fff; box-shadow:0 2px 8px rgba(39,34,32,.2); }
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p{ color:var(--adk); }

.cost-hint{ text-align:center; color:var(--muted); font-size:.78rem; margin:.55rem 0 0; }
.cost-hint b{ color:var(--adk); font-weight:600; }

.stButton>button{ background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  border:0; border-radius:13px; padding:.85rem 1.2rem; font-weight:600; font-size:.98rem;
  box-shadow:0 10px 24px rgba(181,121,58,.28); transition:transform .08s,filter .2s; }
.stButton>button:hover:not(:disabled){ transform:translateY(-1px); filter:brightness(1.04); }
.stButton>button:disabled{ background:#D6CDC0; box-shadow:none; }
[data-testid="stDownloadButton"]>button{ background:var(--ink); color:#fff; border:0;
  border-radius:13px; padding:.9rem 1.2rem; font-weight:600; width:100%;
  box-shadow:0 10px 24px rgba(39,34,32,.18); }
[data-testid="stDownloadButton"]>button:hover{ background:#1c1815; }

.pill{ display:inline-block; font-size:11px; font-weight:600; background:#F4EFE8;
  border:1px solid #EAE3D8; border-radius:999px; padding:3px 11px; color:#6B5E52;
  margin:0 5px 5px 0; }
.note{ color:var(--muted); font-size:.82rem; margin-top:.4rem; }
.stExpander{ border:1px solid var(--line) !important; border-radius:12px !important; }

/* ── Reconciliation report — visual polish (presentation only) ───────────── */
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h1,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h2{
  font-size:1.32rem; font-weight:600; margin:1.5rem 0 .7rem; padding-bottom:.4rem;
  border-bottom:2px solid var(--line); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h3{
  font-size:1.06rem; color:var(--adk); margin:1.15rem 0 .45rem;
  padding-left:.6rem; border-left:3px solid var(--accent); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h1:first-child,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h2:first-child{ margin-top:.2rem; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] p,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] li{ line-height:1.7; }

[class*="st-key-report"] [data-testid="stMarkdownContainer"] table{
  width:100%; border-collapse:separate; border-spacing:0; margin:1rem 0; font-size:.9rem;
  border:1px solid var(--line); border-radius:12px; overflow:hidden;
  box-shadow:0 1px 2px rgba(39,34,32,.04),0 6px 18px rgba(39,34,32,.04); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] thead th{
  background:linear-gradient(135deg,var(--accent),var(--adk)); color:#fff;
  font-weight:600; text-align:left; padding:.62rem .85rem; white-space:nowrap; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody td{
  padding:.55rem .85rem; border-top:1px solid var(--line); vertical-align:top; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody tr:nth-child(even){ background:var(--bg2); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody tr:hover{ background:rgba(181,121,58,.07); }

[class*="st-key-report"] [data-testid="stMarkdownContainer"] blockquote{
  background:rgba(181,121,58,.07); border-left:4px solid var(--accent);
  border-radius:0 10px 10px 0; padding:.75rem 1rem; margin:1.1rem 0; color:var(--ink); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] blockquote p{ margin:0; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] strong{ color:var(--adk); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] code{
  background:var(--bg2); border:1px solid var(--line); border-radius:6px;
  padding:.05rem .38rem; font-size:.85em; color:var(--adk); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] hr{
  border:0; border-top:1px solid var(--line); margin:1.5rem 0; }
</style>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────────────────────
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# Drop results left over from an older code version
_res = st.session_state.get("result")
if _res is not None and not (isinstance(_res, dict) and _res.get("version") == RESULT_VERSION):
    st.session_state.pop("result", None)


def start_new():
    st.session_state.pop("result", None)
    st.session_state["uploader_key"] += 1


def pills(r):
    st.markdown(
        '<div>'
        f'<span class="pill">{r["model_key"]}</span>'
        f'<span class="pill">In {r["in_tok"]:,} · Out {r["out_tok"]:,} tokens</span>'
        f'<span class="pill">₹ {r["cost_inr"]:,.2f} ($ {r["cost_usd"]:.4f})</span>'
        '</div>', unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>Ledger Reconciliation</h1>
  <p>Upload two ledgers — Claude reconciles them live, just like in chat.</p>
</div>""", unsafe_allow_html=True)

# ── Results view ──────────────────────────────────────────────────────────────
if "result" in st.session_state:
    r = st.session_state["result"]
    pills(r)
    with st.container(border=True, key="report"):
        st.markdown(r["report"])

    st.download_button(
        "⬇️  Download reconciliation (.xlsx)",
        data=r["excel_bytes"],
        file_name=f"Reco_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

    with st.expander("🔍 Excel structure Claude chose"):
        st.json(r["workbook"])

    st.button("↺  Start a new reconciliation", on_click=start_new, use_container_width=True)
    st.stop()

# ── Form view ─────────────────────────────────────────────────────────────────
uk = st.session_state["uploader_key"]

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">1</span>'
                '<span class="card-t">Upload the two ledgers</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-sub">Both parties\' books — Excel, CSV, PDF or Word</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        file_a = st.file_uploader("Ledger A", type=UPLOAD_TYPES, key=f"a{uk}")
    with c2:
        file_b = st.file_uploader("Ledger B", type=UPLOAD_TYPES, key=f"b{uk}")

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">2</span>'
                '<span class="card-t">Choose model & run</span></div>', unsafe_allow_html=True)
    model_key = st.radio("Model", MODEL_NAMES,
                         index=DEFAULT_INDEX, horizontal=True, label_visibility="collapsed")
    st.markdown(
        '<div class="cost-hint">Cost rises left → right — '
        '<b>Haiku</b> cheapest · <b>Opus</b> most capable</div>',
        unsafe_allow_html=True)
    run_btn = st.button("✨  Run Reconciliation", type="primary",
                        disabled=not (file_a and file_b and API_KEY),
                        use_container_width=True)
    if not API_KEY:
        st.markdown('<div class="note">⚠️ Set <b>ANTHROPIC_API_KEY</b> in your .env or Streamlit secrets.</div>',
                    unsafe_allow_html=True)

# ── Run: stream the report live, then build the Excel ────────────────────────
if run_btn:
    try:
        with st.spinner("Reading the ledgers…"):
            sess = RecoSession(
                API_KEY, model_key,
                file_a.getvalue(), file_a.name,
                file_b.getvalue(), file_b.name,
            )

        st.markdown("### 📄 Reconciliation report")
        st.caption("Streaming live as Claude writes it — a short pause at the start is Claude thinking.")
        with st.container(border=True, key="reportlive"):
            st.write_stream(sess.stream_report())

        with st.spinner("Building the Excel workbook…"):
            excel_bytes = sess.build_workbook()

        st.session_state["result"] = {
            "version": RESULT_VERSION,
            "report": sess.report,
            "excel_bytes": excel_bytes,
            "workbook": sess.workbook,
            "model_key": model_key,
            "in_tok": sess.input_tokens,
            "out_tok": sess.output_tokens,
            "cost_usd": sess.cost_usd,
            "cost_inr": sess.cost_inr,
        }
        st.rerun()
    except RecoError as e:
        st.error(f"**{e.message}**")
        if e.hint:
            st.info(e.hint)
    except Exception:
        st.error("**Something unexpected happened.** Please try running it again.")
        with st.expander("Technical details (for debugging)"):
            st.code(traceback.format_exc())
