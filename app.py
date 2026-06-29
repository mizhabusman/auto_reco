"""
app.py — AI-Powered Ledger Reconciliation (Streamlit).

Run:  streamlit run app.py
API key: put ANTHROPIC_API_KEY in a .env file (see .env.example).
"""

from __future__ import annotations
import os
import traceback
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

from reco import run_reconciliation, MODELS, DEFAULT_USD_INR
from writer import write_reco_excel

load_dotenv()
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
USD_INR = DEFAULT_USD_INR

MODEL_NAMES = ["Haiku", "Sonnet", "Opus"]
MODEL_CAPTIONS = ["Low cost · Fast", "Balanced · Recommended", "High cost · Most capable"]
NAME_TO_LABEL = {
    "Haiku": "Haiku (fastest, lowest cost)",
    "Sonnet": "Sonnet (balanced — recommended)",
    "Opus": "Opus (most capable, highest cost)",
}
DEFAULT_INDEX = 1  # Sonnet

st.set_page_config(page_title="Ledger Reconciliation", page_icon="📘", layout="centered")

# ===========================================================================
# Theme
# ===========================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap');

:root{
  --bg:#F3EEE7; --bg2:#FBF8F3; --card:#FFFFFF; --ink:#272220; --muted:#8C8379;
  --accent:#B5793A; --accent-dark:#8C5C2B; --line:#EAE3D8; --line-soft:#F1EBE1;
}

/* strip default Streamlit chrome */
header[data-testid="stHeader"], #MainMenu, footer, div[data-testid="stToolbar"]{ display:none !important; }
section[data-testid="stSidebar"]{ display:none !important; }

.stApp{
  background:
    radial-gradient(900px 420px at 50% -8%, var(--bg2) 0%, rgba(251,248,243,0) 70%),
    linear-gradient(180deg, var(--bg) 0%, #EFE9E1 100%) fixed;
}
.block-container{ padding:2.4rem 1rem 4rem; max-width:720px; }
html, body, [class*="css"]{ font-family:'Inter',sans-serif; color:var(--ink); }
h1,h2,h3{ font-family:'Fraunces',serif !important; color:var(--ink); letter-spacing:-.015em; }

/* tighten vertical rhythm */
div[data-testid="stVerticalBlock"]{ gap:.85rem; }

/* ---------- header ---------- */
.app-header{ text-align:center; margin-bottom:1.6rem; }
.app-header .badge{
  display:inline-flex; align-items:center; gap:6px; font-size:.7rem; font-weight:600;
  letter-spacing:.14em; text-transform:uppercase; color:var(--accent-dark);
  background:rgba(181,121,58,.08); border:1px solid rgba(181,121,58,.22);
  border-radius:999px; padding:5px 14px; margin-bottom:14px;
}
.app-header .badge .dot{ width:6px; height:6px; border-radius:50%; background:var(--accent); }
.app-header h1{ font-size:2.35rem; font-weight:600; margin:0; line-height:1.08; }
.app-header p{ color:var(--muted); margin:.55rem 0 0; font-size:1rem; }

/* ---------- real bordered containers become cards ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--card); border:1px solid var(--line) !important; border-radius:18px;
  padding:1.4rem 1.5rem; box-shadow:0 1px 2px rgba(39,34,32,.04), 0 12px 30px rgba(39,34,32,.05);
  transition:box-shadow .25s, transform .25s;
}
div[data-testid="stVerticalBlockBorderWrapper"]:hover{
  box-shadow:0 1px 2px rgba(39,34,32,.04), 0 18px 40px rgba(39,34,32,.08);
}

/* card heading row */
.card-h{ display:flex; align-items:center; gap:10px; margin-bottom:.4rem; }
.card-h .num{ background:linear-gradient(135deg,var(--accent),var(--accent-dark)); color:#fff;
  width:26px; height:26px; border-radius:8px; display:inline-flex; align-items:center;
  justify-content:center; font-size:.82rem; font-weight:600; font-family:'Inter';
  box-shadow:0 4px 10px rgba(181,121,58,.3); }
.card-h .t{ font-family:'Fraunces',serif; font-weight:600; font-size:1.12rem; }
.card-sub{ color:var(--muted); font-size:.86rem; margin:-.1rem 0 .2rem 36px; }

/* ---------- uploaders ---------- */
div[data-testid="stFileUploader"] section{
  background:var(--bg2); border:1.5px dashed var(--line); border-radius:12px;
  padding:.6rem .8rem; transition:border-color .2s, background .2s;
}
div[data-testid="stFileUploader"] section:hover{ border-color:var(--accent); background:#FCFAF6; }
div[data-testid="stFileUploader"] label{ font-weight:600; color:var(--ink); font-size:.9rem; }
div[data-testid="stFileUploader"] button{
  border:1px solid var(--line) !important; border-radius:9px !important;
  color:var(--ink) !important; background:#fff !important; font-weight:600 !important;
}

/* ---------- model toggle (segmented cards: name + lighter caption) ---------- */
div[data-testid="stRadio"]{ width:100%; }
div[data-testid="stRadio"] > div{ width:100%; }
div[role="radiogroup"]{ display:flex; width:100%; gap:10px; flex-wrap:nowrap; }
div[role="radiogroup"] > label{
  flex:1 1 0; min-width:0; min-height:64px; box-sizing:border-box; overflow:hidden;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px;
  background:var(--bg2); border:1.5px solid var(--line); border-radius:12px;
  padding:8px 10px; margin:0 !important; transition:all .15s; cursor:pointer;
}
/* hide the native radio circle for a clean card look */
div[role="radiogroup"] > label > div:first-child{ display:none !important; }
/* model name (main label) */
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"]{ width:100%; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p{
  margin:0; white-space:nowrap; font-size:.98rem; font-weight:600;
  color:var(--ink); line-height:1.15; text-align:center;
}
/* cost descriptor (caption) — lighter, smaller, below; wraps if tight */
div[role="radiogroup"] > label div[data-testid="stCaptionContainer"]{ width:100%; }
div[role="radiogroup"] > label div[data-testid="stCaptionContainer"] p,
div[role="radiogroup"] > label div[data-testid="stCaptionContainer"]{
  margin:0; font-size:.72rem !important; font-weight:500;
  color:var(--muted) !important; line-height:1.15; text-align:center;
}
div[role="radiogroup"] > label:hover{ border-color:var(--accent); background:#FCFAF6; }
div[role="radiogroup"] > label:has(input:checked){
  border-color:var(--accent); background:rgba(181,121,58,.09);
  box-shadow:0 4px 14px rgba(181,121,58,.16);
}
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p{
  color:var(--accent-dark);
}

/* ---------- primary button ---------- */
.stButton>button{
  background:linear-gradient(135deg,var(--accent),var(--accent-dark)); color:#fff;
  border:0; border-radius:13px; padding:.85rem 1.2rem; font-weight:600; font-size:.98rem;
  box-shadow:0 10px 24px rgba(181,121,58,.28); transition:transform .08s, box-shadow .2s, filter .2s;
}
.stButton>button:hover:not(:disabled){ transform:translateY(-1px); filter:brightness(1.04);
  box-shadow:0 14px 30px rgba(181,121,58,.36); }
.stButton>button:disabled{ background:#D6CDC0; color:#fff; box-shadow:none; }

/* secondary (start new) button */
.stButton>button[kind="secondary"]{
  background:#fff; color:var(--ink); border:1.5px solid var(--line); box-shadow:none;
}
.stButton>button[kind="secondary"]:hover{ border-color:var(--accent); color:var(--accent-dark);
  background:#FCFAF6; transform:none; }

/* ---------- metrics ---------- */
div[data-testid="stMetric"]{
  background:var(--bg2); border:1px solid var(--line-soft); border-radius:12px; padding:12px 14px;
}
div[data-testid="stMetricLabel"] p{ color:var(--muted); font-size:.78rem; font-weight:600; }
div[data-testid="stMetricValue"]{ font-family:'Fraunces',serif; }

/* ---------- download ---------- */
[data-testid="stDownloadButton"]>button{
  background:var(--ink); color:#fff; border:0; border-radius:13px; padding:.9rem 1.2rem;
  font-weight:600; width:100%; box-shadow:0 10px 24px rgba(39,34,32,.18);
}
[data-testid="stDownloadButton"]>button:hover{ background:#1c1815; }

/* chips */
.chips{ margin-top:.2rem; }
.chip{ display:inline-block; background:var(--bg2); color:var(--accent-dark);
  border:1px solid var(--line); border-radius:999px; padding:4px 12px; font-size:.8rem;
  font-weight:600; margin:0 6px 6px 0; }

.stExpander{ border:1px solid var(--line) !important; border-radius:12px !important; background:var(--card); }
.note{ color:var(--muted); font-size:.82rem; margin-top:.4rem; }
</style>
""", unsafe_allow_html=True)

# ===========================================================================
# State helpers
# ===========================================================================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

def start_new():
    st.session_state.pop("result", None)
    st.session_state.pop("excel_bytes", None)
    st.session_state["uploader_key"] += 1

# ===========================================================================
# Loader (client-side animation; runs during the blocking API call)
# ===========================================================================
def loader_html() -> str:
    return """
<div style="font-family:'Inter',sans-serif;display:flex;flex-direction:column;align-items:center;
     justify-content:center;padding:36px 16px;background:#fff;border:1px solid #EAE3D8;
     border-radius:18px;box-shadow:0 12px 30px rgba(39,34,32,.06);">
  <div class="ring"></div>
  <div id="ld-msg" style="margin-top:20px;font-weight:600;color:#272220;font-size:1.02rem;
       min-height:1.4em;font-family:'Inter';">Reading the ledgers…</div>
  <div style="margin-top:6px;color:#8C8379;font-size:.84rem;">Usually takes 30–90 seconds.</div>
  <div class="bar"><div class="bar-fill"></div></div>
</div>
<style>
.ring{ width:58px;height:58px;border-radius:50%;border:5px solid #F0E7DA;
  border-top-color:#B5793A;border-right-color:#D9A766;animation:spin 1s linear infinite; }
@keyframes spin{ to{ transform:rotate(360deg); } }
.bar{ width:240px;height:6px;background:#F0E7DA;border-radius:99px;margin-top:22px;overflow:hidden; }
.bar-fill{ height:100%;width:40%;border-radius:99px;background:linear-gradient(90deg,#B5793A,#D9A766);
  animation:slide 1.6s ease-in-out infinite; }
@keyframes slide{ 0%{ margin-left:-40%; } 100%{ margin-left:100%; } }
</style>
<script>
const m=["Reading the ledgers…","Detecting headers and cleaning rows…","Matching invoices across both books…",
"Checking amounts and TDS differences…","Spotting timing and missing entries…",
"Compiling the reconciliation…","Writing up key observations…"];
let i=0;const e=document.getElementById('ld-msg');
setInterval(()=>{i=(i+1)%m.length;if(e)e.textContent=m[i];},2600);
</script>
"""

# ===========================================================================
# Header
# ===========================================================================
st.markdown("""
<div class="app-header">
  <span class="badge"><span class="dot"></span>AI-Powered</span>
  <h1>Ledger Reconciliation</h1>
  <p>Upload two ledgers and download a ready reconciliation in seconds.</p>
</div>
""", unsafe_allow_html=True)

# ===========================================================================
# RESULTS VIEW
# ===========================================================================
if "result" in st.session_state:
    result = st.session_state["result"]
    data = result.data
    summary = data.get("summary", {}) or {}
    meta = data.get("meta", {}) or {}
    sheets = data.get("sheets", []) or []

    with st.container(border=True):
        st.markdown('<div class="card-h"><span class="num">✓</span>'
                    '<span class="t">Reconciliation ready</span></div>', unsafe_allow_html=True)
        if meta.get("party_a") or meta.get("party_b"):
            st.markdown(f'<div class="card-sub">{meta.get("party_a","Party A")} '
                        f'&nbsp;↔&nbsp; {meta.get("party_b","Party B")}'
                        f'{" · " + meta.get("period","") if meta.get("period") else ""}</div>',
                        unsafe_allow_html=True)

        if sheets:
            cols = st.columns(min(len(sheets), 4))
            for i, sh in enumerate(sheets[:4]):
                with cols[i]:
                    st.metric(sh.get("name", f"Sheet {i+1}"), len(sh.get("rows", []) or []))

        insights = summary.get("insights", []) or []
        if insights:
            with st.expander("💡 Key Observations", expanded=True):
                for ins in insights:
                    st.markdown(f"- {ins}")

        st.markdown(
            f'<div class="chips"><span class="chip">{result.model_label.split(" (")[0]}</span>'
            f'<span class="chip">Tokens · in {result.input_tokens:,} / out {result.output_tokens:,}</span>'
            f'<span class="chip">This run: ₹{result.cost_inr:,.2f} (${result.cost_usd:.4f})</span></div>',
            unsafe_allow_html=True,
        )

        fname = f"Reco_{meta.get('party_a','A')}_vs_{meta.get('party_b','B')}_{datetime.now():%Y%m%d_%H%M}.xlsx"
        fname = "".join(c if c.isalnum() or c in "._-" else "_" for c in fname)
        st.download_button("⬇️  Download Reconciliation Report (.xlsx)",
                           data=st.session_state["excel_bytes"], file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    with st.expander("🔍 Raw output from Claude (debug)"):
        st.json(data)

    st.button("↺  Start a New Reconciliation", on_click=start_new,
              use_container_width=True, type="secondary")
    st.stop()

# ===========================================================================
# FORM VIEW
# ===========================================================================
uk = st.session_state["uploader_key"]

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">1</span>'
                '<span class="t">Upload the two ledgers</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-sub">Both parties\' books — Excel or CSV.</div>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        file_a = st.file_uploader("Ledger A", type=["xls", "xlsx", "xlsm", "csv"], key=f"a{uk}")
    with c2:
        file_b = st.file_uploader("Ledger B", type=["xls", "xlsx", "xlsm", "csv"], key=f"b{uk}")

with st.container(border=True):
    st.markdown('<div class="card-h"><span class="num">2</span>'
                '<span class="t">Choose model & run</span></div>', unsafe_allow_html=True)
    choice = st.radio("Model", MODEL_NAMES, captions=MODEL_CAPTIONS,
                      index=DEFAULT_INDEX, horizontal=True, label_visibility="collapsed")
    model_label = NAME_TO_LABEL[choice]
    run = st.button("✨  Run Reconciliation", type="primary",
                    disabled=not (file_a and file_b and API_KEY), use_container_width=True)
    if not API_KEY:
        st.markdown('<div class="note">⚠️ Add <b>ANTHROPIC_API_KEY</b> to your <code>.env</code> '
                    'file to enable runs.</div>', unsafe_allow_html=True)

if run:
    slot = st.empty()
    with slot:
        components.html(loader_html(), height=270)
    try:
        result = run_reconciliation(
            file_a.getvalue(), file_a.name,
            file_b.getvalue(), file_b.name,
            model_label=model_label, usd_inr=USD_INR, api_key=API_KEY,
        )
        st.session_state["result"] = result
        st.session_state["excel_bytes"] = write_reco_excel(result.data)
        slot.empty()
        st.rerun()
    except Exception as e:
        slot.empty()
        st.error(f"Reconciliation failed: {e}")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())