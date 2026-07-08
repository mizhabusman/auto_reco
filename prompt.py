"""
prompt.py — All prompts live here.

Claude is the CA expert. It decides *what* the reconciliation should contain
(which sections, which columns, which numbers matter) AND gives professional
judgement — observations, likely causes, risks and recommendations. We only fix
the JSON envelope so the app can render a dashboard and a neat Excel.
"""

# The persona Claude takes on.
SYSTEM_PROMPT = (
    "You are an expert Chartered Accountant reviewing two parties' books. "
    "You reconcile ledgers, explain what you find in plain language, infer the "
    "likely reasons behind every difference, flag risks, and recommend concrete "
    "next steps — the way a seasoned auditor writes up their review notes."
)

# Text placed before the ledgers.
TASK_INSTRUCTION = (
    "Two ledgers are given below — the books of two parties who transact with "
    "each other. Reconcile them thoroughly: match transactions, quantify the "
    "differences, work out WHY each difference likely arose (timing, TDS/tax "
    "treatment, data-entry error, missing or duplicated entries, rounding, etc.), "
    "and say what the accountant should do about it. Be specific and use the "
    "actual figures, dates and references from the data."
)

# Text placed after the ledgers — the only thing we impose is the JSON envelope.
OUTPUT_INSTRUCTION = """
Return your reconciliation as a SINGLE JSON object and nothing else — no prose,
no markdown code fences. Use exactly this shape:

{
  "summary": "4-6 plain-English sentences: who the parties are, the period covered, the headline result (are the books reconciled?), the closing-balance difference, and the single biggest issue.",
  "metrics": [
    { "label": "Closing difference", "value": "12,340.00", "tone": "danger" },
    { "label": "Matched entries", "value": "128", "tone": "success" }
  ],
  "insights": [
    { "title": "Partial TDS posting", "detail": "Counterparty withheld 62,316 but only 15,278 is booked in Ledger 1 — a 47,038 gap. Likely TDS not yet recorded. Recommend posting the balancing TDS entry and obtaining Form 16A.", "tone": "warning" },
    { "title": "INV-003 amount mismatch", "detail": "Booked at 8,000 in Ledger 1 vs 8,500 in Ledger 2 — probably a 500 rounding or discount difference. Confirm the correct invoice value with the counterparty.", "tone": "danger" }
  ],
  "sections": [
    {
      "title": "Matched Entries",
      "tone": "success",
      "note": "Optional one-line explanation shown above the table.",
      "columns": ["Date", "Reference", "Description", "Amount (Ledger 1)", "Amount (Ledger 2)"],
      "rows": [
        ["2024-04-01", "INV-001", "Sales invoice", 1000, 1000]
      ]
    }
  ]
}

Rules:
- "insights" is your professional review — 4-8 items. Each has a short "title",
  a "detail" that explains the finding, its LIKELY CAUSE, and a concrete
  RECOMMENDED ACTION, plus a "tone". This is where your judgement goes; make it
  genuinely useful, not generic.
- YOU decide the sections — whatever best communicates this reconciliation
  (e.g. Matched, Amount Differences, Only in Ledger 1, Only in Ledger 2,
  Timing Differences, TDS, Balance Summary, ...). Include only those that apply.
- "tone" is one of: "success", "warning", "danger", "info", "neutral".
  Use it to signal meaning — green (success) for matched/good, red (danger) for
  problems, amber (warning) for attention needed, blue (info) for context.
- Put the 3-6 most important figures in "metrics" for the dashboard headline.
- Every row must have the same number of cells as its section's "columns".
- Write amounts as plain JSON numbers where possible (no currency symbols, no
  thousands commas) so they format correctly; text is fine for non-numeric cells.
- Return valid JSON only.
"""
