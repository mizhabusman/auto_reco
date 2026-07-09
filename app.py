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
/* ══════════════════════════ Design tokens ══════════════════════════ */
:root{
  /* palette */
  --bg:#F4F3F1; --bg-elev:#FBFAF8; --card:#FFFFFF;
  --ink:#221E1B; --ink-soft:#4A423C; --muted:#8A8079;
  --accent:#B5793A; --accent-dark:#8C5C2B; --accent-soft:rgba(181,121,58,.08);
  --line:#E9E2D7; --line-soft:#F1ECE3;
  --ok:#3E7D53; --warn:#B5793A; --bad:#B4433B;
  /* aliases kept so scoped report CSS below keeps reading the same names */
  --bg2:#FAF8F5; --adk:#8C5C2B;
  /* spacing scale (4 / 8) */
  --s1:4px; --s2:8px; --s3:12px; --s4:16px; --s5:20px; --s6:24px; --s7:32px; --s8:48px;
  /* radii */
  --r-sm:8px; --r-md:12px; --r-lg:16px; --r-xl:22px; --r-pill:999px;
  /* shadows */
  --sh-sm:0 1px 2px rgba(34,30,27,.05);
  --sh-md:0 1px 2px rgba(34,30,27,.04), 0 10px 28px rgba(34,30,27,.06);
  --sh-accent:0 8px 20px rgba(181,121,58,.26);
  /* type + motion + sizing */
  --font:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;
  --t:.2s cubic-bezier(.4,0,.2,1);
  --control-h:48px;
}

/* ══════════════════════════ App chrome / base ══════════════════════════ */
header[data-testid="stHeader"],#MainMenu,footer,div[data-testid="stToolbar"],
section[data-testid="stSidebar"]{ display:none !important; }
.stApp{ background:
  radial-gradient(1100px 460px at 50% -12%, rgba(181,121,58,.06), transparent 68%),
  var(--bg); }
.block-container{ max-width:820px; padding:clamp(24px,4vw,56px) clamp(16px,4vw,24px) 72px; }
html,body,[class*="css"]{ font-family:var(--font); color:var(--ink);
  -webkit-font-smoothing:antialiased; text-rendering:optimizeLegibility; }
h1,h2,h3,h4{ font-family:var(--font) !important; color:var(--ink); letter-spacing:-.02em; }
p{ line-height:1.6; }

/* Vertical rhythm: 24px between top-level sections (cards tighten themselves below) */
[data-testid="stVerticalBlock"]{ gap:var(--s6); }

/* ══════════════════════════ Header ══════════════════════════ */
.app-header{ text-align:center; margin:0; }
.brand{ width:56px; height:56px; margin:0 auto var(--s4); border-radius:var(--r-lg);
  background:linear-gradient(135deg,var(--accent),var(--accent-dark));
  display:flex; align-items:center; justify-content:center; box-shadow:var(--sh-accent); }
.brand svg{ width:28px; height:28px; }
.app-header h1{ font-size:clamp(2rem,4.4vw,2.6rem); font-weight:600; line-height:1.08; margin:0; }
.app-header p{ color:var(--muted); font-size:1.02rem; line-height:1.5;
  margin:var(--s3) auto 0; max-width:34rem; }

/* ══════════════════════════ Cards ══════════════════════════ */
/* st.container(border=True, key=...) → stable .st-key-<key> class on the block */
[class*="st-key-card"], [class*="st-key-report"]{
  background:var(--card) !important; border:1px solid var(--line) !important;
  border-radius:var(--r-xl) !important; padding:var(--s7) !important;
  box-shadow:var(--sh-md); gap:var(--s5) !important; }
/* tighten blocks nested inside a card (e.g. an uploader's label + dropzone) */
[class*="st-key-card"] [data-testid="stVerticalBlock"],
[class*="st-key-report"] [data-testid="stVerticalBlock"]{ gap:var(--s2); }
.card-h{ display:flex; align-items:center; gap:var(--s3); }
.num{ width:32px; height:32px; flex:0 0 auto; border-radius:var(--r-sm);
  background:linear-gradient(135deg,var(--accent),var(--accent-dark)); color:#fff;
  display:inline-flex; align-items:center; justify-content:center;
  font-size:.9rem; font-weight:700; box-shadow:var(--sh-accent); }
.card-t{ font-size:1.2rem; font-weight:600; letter-spacing:-.01em; }
.card-sub{ color:var(--muted); font-size:.9rem; line-height:1.45;
  margin:var(--s2) 0 0 calc(32px + var(--s3)); }

/* ══════════════════════════ File uploaders ══════════════════════════ */
[data-testid="stWidgetLabel"] p{ font-size:.85rem; font-weight:600; color:var(--ink-soft);
  margin:0 0 var(--s2); }
[data-testid="stFileUploaderDropzone"]{
  background:var(--bg-elev); border:1.5px dashed var(--line); border-radius:var(--r-md);
  padding:var(--s5) var(--s4); min-height:104px; align-items:center;
  transition:border-color var(--t), background var(--t); }
[data-testid="stFileUploaderDropzone"]:hover{ border-color:var(--accent); background:var(--accent-soft); }
[data-testid="stFileUploaderDropzone"] button{
  border:1px solid var(--line) !important; background:var(--card) !important; color:var(--ink) !important;
  border-radius:var(--r-sm) !important; font-weight:600 !important; box-shadow:var(--sh-sm) !important;
  min-height:36px !important; transition:border-color var(--t), color var(--t) !important; }
[data-testid="stFileUploaderDropzone"] button:hover{ border-color:var(--accent) !important; color:var(--accent-dark) !important; }
[data-testid="stFileUploaderFile"]{ border-radius:var(--r-sm); }

/* ══════════════════════════ Model toggle (segmented pill) ══════════════════════════ */
div[role="radiogroup"]{ position:relative; display:flex; width:100%; gap:var(--s1); flex-wrap:nowrap;
  background:linear-gradient(135deg,var(--accent),var(--accent-dark)); border-radius:var(--r-pill); padding:5px;
  box-shadow:var(--sh-accent), inset 0 1px 2px rgba(255,255,255,.18); }
div[role="radiogroup"] > label{ flex:1 1 0; min-width:0;
  display:flex; align-items:center; justify-content:center;
  background:transparent; border:0; border-radius:var(--r-pill); padding:10px 8px; margin:0 !important;
  cursor:pointer; transition:background-color var(--t), box-shadow var(--t), transform .12s var(--t); }
div[role="radiogroup"] > label > div:first-child{ display:none !important; }  /* hide the radio dot */
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{
  margin:0; white-space:nowrap; font-size:1rem; font-weight:600; text-align:center;
  color:rgba(255,255,255,.94); transition:color var(--t); }
div[role="radiogroup"] > label:hover:not(:has(input:checked)){ background:rgba(255,255,255,.14); }
div[role="radiogroup"] > label:active{ transform:scale(.97); }
div[role="radiogroup"] > label:has(input:checked){ background:#fff; box-shadow:0 2px 10px rgba(34,30,27,.22); }
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p{ color:var(--adk); }

.cost-hint{ display:flex; align-items:center; justify-content:center; flex-wrap:wrap; gap:var(--s2);
  text-align:center; color:var(--muted); font-size:.8rem; margin:0; }
.cost-hint b{ color:var(--adk); font-weight:600; }
.cost-scale{ display:inline-flex; align-items:center; gap:5px; }
.cost-scale i{ width:5px; height:5px; border-radius:50%; background:var(--accent); opacity:.35; }
.cost-scale i:nth-child(2){ opacity:.6; } .cost-scale i:nth-child(3){ opacity:1; }

/* ══════════════════════════ Buttons ══════════════════════════ */
.stButton>button, [data-testid="stDownloadButton"]>button{
  width:100%; min-height:var(--control-h); border-radius:var(--r-md);
  font-weight:600; font-size:1rem; letter-spacing:.01em; border:1px solid transparent;
  transition:transform .08s var(--t), filter var(--t), background var(--t),
             border-color var(--t), box-shadow var(--t); }
.stButton>button:active, [data-testid="stDownloadButton"]>button:active{ transform:translateY(1px); }
/* primary — Run */
.stButton button[kind="primary"]{
  background:linear-gradient(135deg,var(--accent),var(--accent-dark)); color:#fff;
  box-shadow:var(--sh-accent); }
.stButton button[kind="primary"]:hover:not(:disabled){ transform:translateY(-1px); filter:brightness(1.05); }
.stButton button[kind="primary"]:disabled{ background:#DAD3C8; color:#fff; box-shadow:none; opacity:1; }
/* secondary — Start a new reconciliation */
.stButton button[kind="secondary"]{
  background:var(--card); color:var(--ink); border-color:var(--line); box-shadow:var(--sh-sm); }
.stButton button[kind="secondary"]:hover:not(:disabled){ border-color:var(--accent); color:var(--accent-dark); }
/* download — dark */
[data-testid="stDownloadButton"]>button{ background:var(--ink); color:#fff; box-shadow:var(--sh-md); }
[data-testid="stDownloadButton"]>button:hover{ background:#15110F; transform:translateY(-1px); }

/* ══════════════════════════ Results meta chips ══════════════════════════ */
.meta{ display:flex; flex-wrap:wrap; gap:var(--s2); }
.pill{ display:inline-flex; align-items:center; gap:var(--s2);
  font-size:.82rem; font-weight:600; color:var(--ink-soft);
  background:var(--card); border:1px solid var(--line); border-radius:var(--r-pill);
  padding:7px 14px; box-shadow:var(--sh-sm); }
.pill .k{ font-size:.66rem; font-weight:700; letter-spacing:.09em; text-transform:uppercase;
  color:var(--muted); }

.note{ display:flex; align-items:center; gap:var(--s2); color:var(--muted); font-size:.84rem; margin:0; }
.note b{ color:var(--ink-soft); }

/* ══════════════════════════ Expander & alerts ══════════════════════════ */
[data-testid="stExpander"]{ border:1px solid var(--line) !important; border-radius:var(--r-md) !important;
  background:var(--card); box-shadow:var(--sh-sm); overflow:hidden; }
[data-testid="stExpander"] summary{ padding:var(--s4) var(--s5); font-weight:600; }
[data-testid="stExpander"] summary:hover{ color:var(--accent-dark); }
[data-testid="stAlert"]{ border-radius:var(--r-md); border:1px solid var(--line-soft); }

/* Run-phase heading + caption */
.run-title{ font-size:1.25rem; font-weight:600; letter-spacing:-.01em; margin:0; }
[data-testid="stCaptionContainer"] p{ color:var(--muted); }

/* ══════════════════════════ Footer ══════════════════════════ */
.app-footer{ text-align:center; color:var(--muted); font-size:.8rem; line-height:1.5;
  margin-top:var(--s7); padding-top:var(--s5); border-top:1px solid var(--line-soft); }
.app-footer b{ color:var(--ink-soft); font-weight:600; }

/* ══════════════════════════ Reconciliation report (presentation only) ══════════════════════════ */
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h1,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h2{
  font-size:1.32rem; font-weight:600; margin:var(--s6) 0 var(--s3); padding-bottom:var(--s2);
  border-bottom:2px solid var(--line); letter-spacing:-.01em; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h3{
  font-size:1.06rem; color:var(--adk); margin:var(--s5) 0 var(--s2);
  padding-left:var(--s2); border-left:3px solid var(--accent); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h1:first-child,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] h2:first-child{ margin-top:var(--s1); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] p,
[class*="st-key-report"] [data-testid="stMarkdownContainer"] li{ line-height:1.7; }

[class*="st-key-report"] [data-testid="stMarkdownContainer"] table{
  width:100%; border-collapse:separate; border-spacing:0; margin:var(--s4) 0; font-size:.9rem;
  border:1px solid var(--line); border-radius:var(--r-md); overflow:hidden;
  box-shadow:var(--sh-sm); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] thead th{
  background:linear-gradient(135deg,var(--accent),var(--accent-dark)); color:#fff;
  font-weight:600; text-align:left; padding:var(--s3) var(--s4); white-space:nowrap; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody td{
  padding:var(--s3) var(--s4); border-top:1px solid var(--line); vertical-align:top; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody tr:nth-child(even){ background:var(--bg2); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] tbody tr:hover{ background:rgba(181,121,58,.07); }

[class*="st-key-report"] [data-testid="stMarkdownContainer"] blockquote{
  background:var(--accent-soft); border-left:4px solid var(--accent);
  border-radius:0 var(--r-md) var(--r-md) 0; padding:var(--s3) var(--s4); margin:var(--s4) 0; color:var(--ink); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] blockquote p{ margin:0; }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] strong{ color:var(--adk); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] code{
  background:var(--bg2); border:1px solid var(--line); border-radius:var(--r-sm);
  padding:.05rem .4rem; font-size:.85em; color:var(--adk); }
[class*="st-key-report"] [data-testid="stMarkdownContainer"] hr{
  border:0; border-top:1px solid var(--line); margin:var(--s6) 0; }

/* ══════════════════════════ Responsive ══════════════════════════ */
@media (max-width:640px){
  [class*="st-key-card"], [class*="st-key-report"]{ padding:var(--s5) !important; }
  [data-testid="stHorizontalBlock"]{ flex-wrap:wrap; gap:var(--s4) !important; }
  [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]{ flex:1 1 100% !important; min-width:100% !important; }
  div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{ font-size:.9rem; }
  [class*="st-key-report"] [data-testid="stMarkdownContainer"] table{ display:block; overflow-x:auto; white-space:nowrap; }
}
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
        '<div class="meta">'
        f'<span class="pill"><span class="k">Model</span>{r["model_key"]}</span>'
        f'<span class="pill"><span class="k">Tokens</span>In {r["in_tok"]:,} · Out {r["out_tok"]:,}</span>'
        f'<span class="pill"><span class="k">Cost</span>₹ {r["cost_inr"]:,.2f} · $ {r["cost_usd"]:.4f}</span>'
        '</div>', unsafe_allow_html=True)


def footer():
    st.markdown(
        '<div class="app-footer">Powered by <b>Claude</b> · '
        'Files are processed in memory and never stored.</div>',
        unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="brand">
    <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8"
         stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 3v18"/><path d="M4 8h16"/>
      <circle cx="7.5" cy="14" r="3.2"/><circle cx="16.5" cy="14" r="3.2"/>
    </svg>
  </div>
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
    footer()
    st.stop()

# ── Form view ─────────────────────────────────────────────────────────────────
uk = st.session_state["uploader_key"]

with st.container(border=True, key="card_upload"):
    st.markdown(
        '<div class="card-head">'
        '<div class="card-h"><span class="num">1</span>'
        '<span class="card-t">Upload the two ledgers</span></div>'
        '<div class="card-sub">Both parties\' books — Excel, CSV, PDF or Word</div>'
        '</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        file_a = st.file_uploader("Ledger A", type=UPLOAD_TYPES, key=f"a{uk}")
    with c2:
        file_b = st.file_uploader("Ledger B", type=UPLOAD_TYPES, key=f"b{uk}")

with st.container(border=True, key="card_model"):
    st.markdown(
        '<div class="card-head">'
        '<div class="card-h"><span class="num">2</span>'
        '<span class="card-t">Choose model &amp; run</span></div>'
        '<div class="card-sub">Higher models reason more deeply — and cost more.</div>'
        '</div>', unsafe_allow_html=True)
    model_key = st.radio("Model", MODEL_NAMES,
                         index=DEFAULT_INDEX, horizontal=True, label_visibility="collapsed")
    st.markdown(
        '<div class="cost-hint">Cost &amp; capability rise left → right — '
        '<b>Haiku</b> is cheapest, <b>Opus</b> the most capable'
        '<span class="cost-scale"><i></i><i></i><i></i></span></div>',
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

        st.markdown('<p class="run-title">📄 Reconciliation report</p>', unsafe_allow_html=True)
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
else:
    footer()
