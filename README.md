# AI Ledger Reconciliation

Upload two ledger exports (Excel/CSV). Claude does the entire reconciliation —
detecting headers, cleaning rows, matching invoices (fuzzy + TDS-aware),
categorizing, and producing insights. You get a formatted Excel report.

## Architecture

```
File A  ─┐
File B  ─┴─►  raw CSV-text dump  ──►  Claude (Sonnet 4.6)  ──►  JSON  ──►  Excel
                                       (prompt.py)                       (writer.py)
```

**No business logic in Python.** All reconciliation rules — matching priorities,
TDS handling, categories, summary insights, output schema — live entirely in
`prompt.py`. Edit that one file when rules change.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI (upload → run → download). |
| `reco.py` | Reads raw files, calls Claude, parses JSON. Thin orchestrator. |
| `prompt.py` | **The brain.** The reconciliation prompt. Edit me. |
| `writer.py` | Turns Claude's JSON into a formatted multi-sheet `.xlsx`. |
| `requirements.txt` | Python deps. |
| `.env.example` | Copy to `.env` and add your Anthropic API key. |

## Setup

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env                # then edit .env with your key
streamlit run app.py
```

Open http://localhost:8501.

## Output structure

The downloaded `.xlsx` has these sheets:

- **Summary** — opening / closing balances per side, totals, category counts,
  AI insights, closing-balance difference.
- **Matched** — invoice + amount agree.
- **Amount Mismatches** — invoice matched, amount differs.
- **TDS Differences** — matched only after TDS adjustment (verify TDS booking).
- **Date Mismatches** — invoice + amount agree, dates differ > 15 days.
- **Unmatched in A** — present only in Party A's books.
- **Unmatched in B** — present only in Party B's books.

## Tweaking behaviour

Open `prompt.py`. The prompt is annotated with sections:

- **STEP 1** — header / column detection rules.
- **STEP 2** — matching priorities (invoice fuzziness, TDS tolerance, date window).
- **STEP 3** — category definitions.
- **STEP 4** — summary computation + insight style.
- **OUTPUT** — strict JSON schema.

If your boss hands you a "standard reco template", reshape the JSON schema in
**OUTPUT** (and the matching sheet headers in `writer.py`). Everything else
stays the same.

## Cost

Sonnet 4.6 at $3 in / $15 out per million tokens.
- 100-row ledgers: ~$0.02–0.05 per reco
- 500-row ledgers: ~$0.10–0.20 per reco

Visible in the UI after each run.

## Hosting

It's stateless. No DB. Any of the following will work:

- **Streamlit Community Cloud** — free, push to GitHub, set the API key as a secret.
- **Render / Railway / Fly.io** — `streamlit run app.py --server.port $PORT`.
- **Internal VM** — same command behind nginx + auth.

## Limits / known caveats

- For ledgers far above 500 rows, the raw-dump approach may approach the model's
  input window. If/when that happens, switch to a chunked flow in `reco.py`
  (split by date range, reconcile chunks, merge JSON). The prompt does not
  need to change.
- Claude sometimes wraps JSON in code fences despite instructions; `reco.py`
  strips them defensively.
- `.xls` (old binary) files require `xlrd==2.0.1` and only support read.
  `.xlsx` is preferred.
