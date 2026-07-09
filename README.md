# AI Ledger Reconciliation — Claude as the CA

Upload two ledger exports (Excel, CSV, PDF or Word). Claude reconciles them exactly the way
it does in chat — no template, no rules imposed — with the report **streaming
live onto the screen** as it's written, followed by a downloadable Excel
workbook whose sheets and columns Claude designed itself.

## How it works

```
File A ─┐                       ┌─► Call 1: "You're a CA — reconcile these."
File B ─┴─► raw text dump ──────┤        └─► markdown report (streamed live)
                                └─► Call 2: "Lay that reconciliation out as
                                     an Excel workbook" → JSON → .xlsx
```

**No reconciliation logic in Python.** Call 1 is a pure chat mimic — Claude
gets the two ledgers and one instruction, so it reconciles at full quality,
finding things you might not. Call 2 structures Claude's *own* findings into
Excel. The app is only plumbing.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI: upload → stream → download. |
| `engine.py` | File reading, the two streamed Claude calls, cost tracking. |
| `prompts.py` | **The only "logic".** Two short prompts. Edit me. |
| `excel_writer.py` | Generic dumper: whatever sheets Claude returns → .xlsx. |
| `BUILD_PROMPT.md` | The full prompt to rebuild this app from scratch. |

## Setup

```bash
python -m venv venv
venv\Scripts\activate          # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
# put ANTHROPIC_API_KEY=... in .env
streamlit run app.py
```

## Models

| Pick | Model | Price (per MTok in/out) |
|---|---|---|
| Haiku | `claude-haiku-4-5` | $1 / $5 |
| Sonnet (default) | `claude-sonnet-4-6` | $3 / $15 |
| Opus | `claude-opus-4-8` | $5 / $25 |

Sonnet and Opus run with adaptive thinking for deeper reconciliation.
Actual cost (both calls combined) is shown after every run.

## Tweaking behaviour

Everything Claude does is driven by [`prompts.py`](prompts.py) — two short
prompts. Want a different focus (GST, forex, a certain tone)? Edit the text
there. Nothing else needs to change.

## Hosting

Stateless, no DB. Streamlit Community Cloud / Render / Railway / a VM —
set `ANTHROPIC_API_KEY` as a secret and run `streamlit run app.py`.
