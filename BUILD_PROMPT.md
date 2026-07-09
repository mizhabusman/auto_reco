# Build Prompt — "Claude the CA" Ledger Reconciliation App

> This is the complete, self-contained prompt used to build this app from
> scratch. Hand it to Claude (or any capable AI coding assistant) verbatim to
> recreate or evolve the app.

---

Build me a small web app that mimics the experience of using Claude in chat
for ledger reconciliation.

## The core idea (do not deviate from this)

When I paste two ledgers into a Claude chat and just say "reconcile this",
Claude beautifully reconciles everything — it finds the ledger structure
itself, matches entries, spots TDS/timing/rounding issues, and even surfaces
things we wouldn't find ourselves. I want that exact experience as a tool:

- **Do NOT impose any reconciliation format, rules, matching logic, schema,
  or template on Claude.** No fixed sheets, no fixed categories, no fixed
  JSON structure for the reconciliation itself. The single instruction is
  essentially: *"You are an expert Chartered Accountant. Here are two
  ledgers. Reconcile them."* Claude decides everything.
- All "logic" must live in one small prompts file. Python is only plumbing.

## Stack

- Python + Streamlit (single-page app), `anthropic` SDK, `pandas`,
  `openpyxl` (+ `xlrd` for legacy .xls), `python-dotenv`.
- API key from `ANTHROPIC_API_KEY` in `.env` / Streamlit secrets.

## UI (one page, warm minimal design)

1. Card 1: two file uploaders side by side — "Ledger A" and "Ledger B"
   (xls/xlsx/xlsm/csv).
2. Card 2: model picker (radio, horizontal): Haiku / Sonnet / Opus, with
   cost captions; Sonnet preselected. A big "Run Reconciliation" button,
   disabled until both files and the API key are present.
3. On run: the reconciliation report **streams onto the screen live, token
   by token, exactly like chat** (use `st.write_stream` over the Anthropic
   streaming API). Then an Excel file is prepared and a download button
   appears.
4. Results view: cost pills (model, in/out tokens, ₹ and $ cost), the full
   markdown report in a card, the Excel download button, an expander showing
   the workbook structure Claude chose, and a "Start a new reconciliation"
   button that resets the uploaders.

## Architecture — two API calls

**Call 1 — the reconciliation (pure chat mimic, streamed):**
Convert both uploaded files to plain text (CSV dump; for Excel, dump every
sheet as CSV text with the sheet name — do NOT try to parse headers or
columns yourself; Claude reads the raw structure). Send one user message:

> You are an expert Chartered Accountant with deep experience in ledger and
> party reconciliation. Below are two ledgers from two different parties.
> [LEDGER A text] [LEDGER B text]
> Reconcile these two ledgers. Present your complete reconciliation as a
> well-organised markdown report — exactly as you would if these files were
> handed to you in a conversation. You decide the structure, the depth, and
> what matters for THIS pair of books. Be thorough: match what matches, dig
> into every difference (amounts, TDS/withholding, timing, missing entries,
> rounding), and point out anything unusual you notice — including things the
> parties themselves may have missed. Use headings and tables where they
> help. End with your professional conclusion.

Stream this response to the screen. No format constraints — that's the point.

**Call 2 — the Excel (structure Claude's own findings):**
Send the two ledgers + the finished report back, asking Claude to return
ONLY a JSON object: `{"sheets": [{"name", "columns", "rows"}]}` — Claude
chooses sheet names, columns, and rows itself; include every line item, not
just totals; amounts as JSON numbers; dates as YYYY-MM-DD; no fences.

Write that JSON to `.xlsx` with a **generic** writer: one sheet per entry,
bold header row (white on dark blue #1F4E78), thin borders, right-aligned
numbers with `#,##0.00` format (negatives red), auto-fit column widths
capped at ~55 chars, frozen header row, sanitised + deduplicated sheet names
(strip `[]:*?/\`, max 31 chars). If no valid sheets come back, fall back to
dumping the report text into a single sheet so the download is never empty.

## Models & API specifics (important — current as of mid-2026)

- Model IDs: `claude-haiku-4-5` ($1/$5 per MTok), `claude-sonnet-4-6`
  ($3/$15), `claude-opus-4-8` ($5/$25). Never invent date-suffixed IDs.
- Use `client.messages.stream(...)` for BOTH calls (long outputs — avoids
  HTTP timeouts); `max_tokens=32000`.
- Enable adaptive thinking (`thinking={"type": "adaptive"}`) on Sonnet 4.6
  and Opus 4.8 for call 1 (reconciliation quality). Haiku 4.5 does not
  support adaptive thinking — omit the parameter there. Call 2 (mechanical
  JSON conversion) runs without thinking.
- Do NOT pass `temperature`/`top_p`/`top_k` (rejected on Opus 4.8).
- Track `usage.input_tokens`/`output_tokens` from both calls; show combined
  cost in ₹ (USD→INR ≈ 85) and $.

## Robustness

- File decoding: try utf-8-sig, utf-8, cp1252, latin-1; CSV delimiter via
  `csv.Sniffer` with fallback to comma. Excel via pandas with openpyxl,
  falling back to xlrd for .xls.
- JSON parsing of call 2: strip markdown fences, fall back to the outermost
  `{...}` slice, and default to `{"sheets": []}` if unparseable.
- Guard `st.session_state` results with a version number so hot-reloads
  after code changes never crash on stale objects.
- Wrap the run in try/except: show `st.error` plus a traceback expander.

## Files

- `app.py` — Streamlit UI only.
- `engine.py` — file→text, the two streamed calls, usage/cost, session class.
- `prompts.py` — the two prompts above. **The only file with any "logic".**
- `excel_writer.py` — the generic workbook dumper.
- `requirements.txt`, `README.md`, `.env` (gitignored) with the API key.

## Acceptance criteria

1. Upload two small CSV ledgers → report streams live → Excel downloads and
   opens in Excel with Claude-chosen sheets.
2. No reconciliation rule, category name, or sheet name appears anywhere in
   the Python code.
3. Changing the CA's behaviour requires editing only `prompts.py`.
