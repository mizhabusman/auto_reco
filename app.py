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
if "result" in st.session_state:
    r = st.session_state["result"]

    with st.container(border=True):
        st.markdown('<div class="card-h"><span class="num">✓</span>'
                    '<span class="card-t">Reconciliation complete</span></div>',
                    unsafe_allow_html=True)

        # CA's plain-English summary
        if r.summary:
            st.markdown(f'<div class="summary-box">{r.summary}</div>',
                        unsafe_allow_html=True)

        # Sheet count chips
        if r.sheets:
            st.markdown(
                " ".join(f'<span class="chip">📄 {sh["name"]} '
                         f'({max(0, len(sh.get("rows",[])) - 1)} rows)</span>'
                         for sh in r.sheets),
                unsafe_allow_html=True)

        # Cost
        st.markdown(
            f'<div style="margin:.6rem 0">'
            f'<span class="chip">{r.model_label.split(" (")[0]}</span>'
            f'<span class="chip">In {r.input_tokens:,} · Out {r.output_tokens:,} tokens</span>'
            f'<span class="chip">₹ {r.cost_inr:,.2f} &nbsp;($ {r.cost_usd:.4f})</span>'
            f'</div>', unsafe_allow_html=True)

        # Download
        pa = r.sheets[0]["rows"][1][1] if r.sheets and len(r.sheets[0].get("rows",[])) > 1 else "A"
        fname = f"Reco_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.download_button(
            "⬇️  Download Reconciliation Report (.xlsx)",
            data=r.excel_bytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    with st.expander("🔍 Raw Claude response"):
        st.text(r.raw_response[:6000] + ("…" if len(r.raw_response) > 6000 else ""))

    st.button("↺  Start a New Reconciliation", on_click=start_new, use_container_width=True)
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