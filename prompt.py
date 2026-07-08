"""
prompt.py — All prompts live here.

Claude is the CA expert. It decides *what* the reconciliation should contain
(which sections, which columns, which numbers matter). We only fix the JSON
envelope so the app can render a dashboard and a neat Excel from the result.
"""

# The persona Claude takes on.
SYSTEM_PROMPT = "You are an expert Chartered Accountant."

# Text placed before the ledgers.
TASK_INSTRUCTION = (
    "Two ledgers are given below. Reconcile them as a Chartered Accountant would "
    "and present your findings clearly."
)

# Text placed after the ledgers — the only thing we impose is the JSON envelope.
OUTPUT_INSTRUCTION = """
Return your reconciliation as a SINGLE JSON object and nothing else — no prose,
no markdown code fences. Use exactly this shape:

{
  "summary": "3-6 plain-English sentences: the parties, the period covered, and the key findings (closing balance difference, biggest unmatched items, TDS gaps, etc.).",
  "metrics": [
    { "label": "Closing difference", "value": "12,340.00", "tone": "danger" },
    { "label": "Matched entries", "value": "128", "tone": "success" }
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
- YOU decide the sections — whatever best communicates this reconciliation
  (e.g. Matched, Amount Differences, Only in Ledger 1, Only in Ledger 2,
  Timing Differences, TDS, Balance Summary, ...). Include only those that apply.
- "tone" is one of: "success", "warning", "danger", "info", "neutral".
  Use it to signal meaning — green (success) for matched/good, red (danger) for
  problems/differences, amber (warning) for items needing attention.
- Put the 3-6 most important figures in "metrics" for the dashboard headline.
- Every row must have the same number of cells as its section's "columns".
- Write amounts as plain JSON numbers where possible (no currency symbols, no
  thousands commas) so they format correctly; text is fine for non-numeric cells.
- Return valid JSON only.
"""
