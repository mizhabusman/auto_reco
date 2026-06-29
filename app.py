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
TONE_MAP = {
    "summary": ("#F0F7F0","#2D6A4F","#2D6A4F","📊"),
    "matched": ("#F0F7F0","#2D6A4F","#2D6A4F","✅"),
    "diff":    ("#FDF2F2","#9B2226","#9B2226","⚠️"),
    "tds":     ("#FEF9EC","#92400E","#B5793A","🧾"),
    "onlya":   ("#F3F0FA","#5E548E","#5E548E","🔵"),
    "onlyb":   ("#EFF6FD","#1D6FA4","#1D6FA4","🟣"),
    "conc":    ("#EFF6FD","#1D6FA4","#1D6FA4","📋"),
    "other":   ("#F7F5F2","#495867","#495867","📄"),
}

def _detect_tone(name):
    n = name.lower()
    if any(k in n for k in ("summary","overview","dashboard","reconciliation summary")): return "summary"
    if any(k in n for k in ("matched","match","agree","payment","invoice","credit","bank","receipt")): return "matched"
    if any(k in n for k in ("diff","mismatch","issue","varianc","key diff","discrepan","amount diff")): return "diff"
    if any(k in n for k in ("tds","tax","withhold")): return "tds"
    if any(k in n for k in ("only in a","only a","unmatched a")): return "onlya"
    if any(k in n for k in ("only in b","only b","unmatched b")): return "onlyb"
    if any(k in n for k in ("conclusion","auditor","finding")): return "conc"
    return "other"

def _fmt(v):
    if isinstance(v, float):  return f"₹{v:,.2f}"
    if isinstance(v, int) and not isinstance(v, bool): return f"{v:,}"
    if v is None or v == "": return "—"
    return str(v)

def _cell(v):
    sv = str(v).strip().upper()
    # Status badges
    if sv in ("MATCHED","CLEAR","SUBSTANTIALLY RECONCILED"):
        return f'<span style="background:#DCFCE7;color:#166534;font-size:11px;font-weight:600;padding:2px 9px;border-radius:999px;white-space:nowrap">✓ {v}</span>'
    if sv in ("UNRESOLVED","HIGH"):
        return f'<span style="background:#FEE2E2;color:#991B1B;font-size:11px;font-weight:600;padding:2px 9px;border-radius:999px;white-space:nowrap">✕ {v}</span>'
    if sv in ("MINOR","MEDIUM","PARTIAL","ACCEPTABLE","LOW"):
        return f'<span style="background:#FEF3C7;color:#92400E;font-size:11px;font-weight:600;padding:2px 9px;border-radius:999px;white-space:nowrap">~ {v}</span>'
    # Numbers
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        color = "#DC2626" if v < 0 else ("#6B7280" if v == 0 else "#111827")
        weight = "600" if v != 0 else "400"
        return f'<span style="color:{color};font-weight:{weight};font-variant-numeric:tabular-nums">{_fmt(v)}</span>'
    return _fmt(v)

def _render_table(headers, data, accent):
    TH = (f'background:#F9F7F4;padding:9px 14px;font-size:11px;font-weight:600;'
          f'letter-spacing:.05em;text-transform:uppercase;color:#6B7280;'
          f'border-bottom:2px solid {accent};white-space:nowrap')
    TD_EVEN = 'padding:9px 14px;font-size:13px;vertical-align:middle;background:#FFFFFF;border-bottom:1px solid #F0EBE3'
    TD_ODD  = 'padding:9px 14px;font-size:13px;vertical-align:middle;background:#FAF8F5;border-bottom:1px solid #F0EBE3'
    ths = "".join(f'<th style="{TH};text-align:{"right" if i>0 and isinstance((data[0][i] if data and len(data[0])>i else ""), (int,float)) else "left"}">{h}</th>' for i,h in enumerate(headers))
    trs = ""
    for ri, row in enumerate(data):
        pad = list(row) + [""]*(max(0,len(headers)-len(row)))
        bg = TD_EVEN if ri%2==0 else TD_ODD
        tds = ""
        for ci, v in enumerate(pad[:len(headers)]):
            is_num = isinstance(v,(int,float)) and not isinstance(v,bool)
            align = "right" if is_num else "left"
            tds += f'<td style="{bg};text-align:{align}">{_cell(v)}</td>'
        trs += f"<tr>{tds}</tr>"
    return (f'<div style="overflow-x:auto;border-radius:10px;border:1px solid #EAE3D8;margin-top:.5rem">'
            f'<table style="width:100%;border-collapse:collapse;min-width:400px">'
            f'<thead><tr>{ths}</tr></thead><tbody>{trs}</tbody></table></div>')

if "result" in st.session_state:
    r = st.session_state["result"]

    # ── Summary card ──────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown(
            '<div style="display:flex;align-items:center;gap:8px;margin-bottom:.7rem">'
            '<div style="width:24px;height:24px;background:#DCFCE7;border-radius:50%;display:flex;'
            'align-items:center;justify-content:center;font-size:13px;color:#166534;font-weight:700">✓</div>'
            '<span style="font-weight:600;font-size:16px;color:#1C1916">Reconciliation complete</span></div>',
            unsafe_allow_html=True)

        if r.summary:
            st.markdown(
                f'<p style="color:#4B4440;font-size:13.5px;line-height:1.75;margin-bottom:.9rem">{r.summary}</p>',
                unsafe_allow_html=True)

        # Cost pills
        st.markdown(
            f'<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:.9rem">'
            f'<span style="font-size:11px;font-weight:600;background:#F4EFE8;border:1px solid #EAE3D8;'
            f'border-radius:999px;padding:3px 11px;color:#6B5E52">{r.model_label.split(" (")[0]}</span>'
            f'<span style="font-size:11px;font-weight:600;background:#F4EFE8;border:1px solid #EAE3D8;'
            f'border-radius:999px;padding:3px 11px;color:#6B5E52">'
            f'In {r.input_tokens:,} · Out {r.output_tokens:,} tokens</span>'
            f'<span style="font-size:11px;font-weight:600;background:#F4EFE8;border:1px solid #EAE3D8;'
            f'border-radius:999px;padding:3px 11px;color:#6B5E52">'
            f'₹ {r.cost_inr:,.2f} ($ {r.cost_usd:.4f})</span>'
            f'</div>', unsafe_allow_html=True)

        fname = f"Reco_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.download_button(
            "⬇️  Download full reconciliation report (.xlsx)",
            data=r.excel_bytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    # ── Sheets as styled expanders ─────────────────────────────────────────────
    if r.sheets:
        st.markdown('<p style="font-weight:600;font-size:12px;color:#8C8379;letter-spacing:.06em;'
                    'text-transform:uppercase;margin:.8rem 0 .4rem">Details by category</p>',
                    unsafe_allow_html=True)

        for sh in r.sheets:
            rows   = sh.get("rows", [])
            n_data = max(0, len(rows) - 1)
            tone   = _detect_tone(sh["name"])
            bg, fg, accent, icon = TONE_MAP[tone]
            label  = f"{icon}  {sh['name']}  ({n_data} {'entry' if n_data==1 else 'entries'})"

            with st.expander(label, expanded=(tone in ("diff","tds","onlya","onlyb"))):
                if len(rows) < 2:
                    st.markdown('<p style="color:#9E9590;font-size:13px;padding:.5rem 0">No entries.</p>',
                                unsafe_allow_html=True)
                    continue

                headers = [str(h) for h in rows[0]]
                data    = rows[1:]

                # Alert banner
                if tone == "diff":
                    st.markdown(
                        '<div style="display:flex;gap:10px;padding:.8rem 1rem;border-radius:8px;'
                        'margin-bottom:.75rem;border-left:3px solid #DC2626;background:#FEF2F2">'
                        '<span style="color:#DC2626;font-size:16px;flex-shrink:0">⚠</span>'
                        '<div><strong style="font-size:13px;color:#1C1916;display:block">Differences found</strong>'
                        '<span style="font-size:12px;color:#6B7280">These items need investigation — '
                        'review each row and confirm with the counterparty.</span></div></div>',
                        unsafe_allow_html=True)
                elif tone in ("onlya","onlyb"):
                    side = "Party A" if tone=="onlya" else "Party B"
                    st.markdown(
                        f'<div style="display:flex;gap:10px;padding:.8rem 1rem;border-radius:8px;'
                        f'margin-bottom:.75rem;border-left:3px solid #D97706;background:#FFFBEB">'
                        f'<span style="color:#D97706;font-size:16px;flex-shrink:0">!</span>'
                        f'<div><strong style="font-size:13px;color:#1C1916;display:block">'
                        f'Present only in {side}\'s books</strong>'
                        f'<span style="font-size:12px;color:#6B7280">No matching entry found in the other ledger.</span>'
                        f'</div></div>', unsafe_allow_html=True)
                elif tone == "tds":
                    st.markdown(
                        '<div style="display:flex;gap:10px;padding:.8rem 1rem;border-radius:8px;'
                        'margin-bottom:.75rem;border-left:3px solid #D97706;background:#FFFBEB">'
                        '<span style="color:#D97706;font-size:16px;flex-shrink:0">🧾</span>'
                        '<div><strong style="font-size:13px;color:#1C1916;display:block">TDS / withholding tax</strong>'
                        '<span style="font-size:12px;color:#6B7280">Verify TDS amounts are correctly booked on both sides.</span>'
                        '</div></div>', unsafe_allow_html=True)

                st.markdown(_render_table(headers, data, accent), unsafe_allow_html=True)

    with st.expander("🔍 Raw Claude response"):
        st.text(r.raw_response[:6000] + ("…" if len(r.raw_response) > 6000 else ""))

    st.button("↺  Start a new reconciliation", on_click=start_new, use_container_width=True)
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