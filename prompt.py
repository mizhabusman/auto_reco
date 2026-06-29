RECO_PROMPT = r"""
You are a Chartered Accountant and an expert in accounts reconciliation.

I have given you TWO ledgers below — both parties' books covering the same
counterparty relationship. Reconcile them the way a CA would: figure out the
structure of each ledger yourself (ignore the company-info/header junk rows),
clean the data, match transactions across the two books, identify what agrees
and what doesn't, and present a clear, professional reconciliation.

Use your own judgement on the best format and the best set of sheets for THIS
data. Typical things a good reco shows — matched entries, amount mismatches,
TDS/withholding differences, timing/date differences, items present in only one
book, opening/closing balance comparison, and a summary with key observations —
but include only what the data actually warrants, and add anything else useful.

When matching, remember ledgers are mirrors (a sale in one is a purchase in the
other), invoice/bill numbers may be written differently on each side
(e.g. "25-26/GS-140" vs "GS-140"), and one side may book amounts net of TDS
while the other books gross. Match on the substance, not the surface.

----- LEDGER A: {file_a_name} -----
{file_a_content}
----- END LEDGER A -----

----- LEDGER B: {file_b_name} -----
{file_b_content}
----- END LEDGER B -----

============================================================
HOW TO RETURN YOUR ANSWER
============================================================
Return your reconciliation as ONE JSON object describing a workbook. You decide
the sheets and columns. Use this envelope exactly (no prose, no markdown
fences, JSON only):

{{
  "title": "string — report title",
  "meta": {{
    "party_a": "string — whose book Ledger A is",
    "party_b": "string — whose book Ledger B is",
    "period": "string",
    "currency": "INR",
    "prepared_note": "string — one line, e.g. 'Reconciled as on ...'"
  }},
  "summary": {{
    "metrics": [
      {{ "label": "string", "value_a": "number or string",
         "value_b": "number or string", "difference": "number or string or null" }}
    ],
    "insights": ["string", "..."]
  }},
  "sheets": [
    {{
      "name": "string — short sheet/tab name (<= 28 chars)",
      "subtitle": "string or null — optional one-line description",
      "tone": "good | warn | bad | neutral",
      "columns": ["Col 1", "Col 2", "..."],
      "rows": [
        ["cell", 123.45, "cell"],
        ["...",  0,       "..."]
      ]
    }}
  ]
}}

RULES:
  • JSON only. No commentary, no ``` fences.
  • You choose how many sheets, their names, their columns, and their order.
  • "tone" colours the sheet tab so the reader can scan severity at a glance:
       good = clean/matched, warn = needs attention, bad = problems, neutral = info.
  • Numbers must be JSON numbers (not strings) wherever a value is monetary or
    countable, so Excel can format and total them. Use 0, not "".
  • Dates as "YYYY-MM-DD" strings. Keep narration short; strip "_x000D_" and
    line breaks.
  • Every row in a sheet must have the same number of cells as that sheet's
    "columns".
  • In "summary.metrics", include at least opening balance, closing balance,
    and their differences between the two books, plus any totals you find useful.
  • Make "insights" specific and actionable (name invoices, amounts, dates).
"""


def build_prompt(file_a_name: str, file_a_content: str,
                 file_b_name: str, file_b_content: str) -> str:
    return RECO_PROMPT.format(
        file_a_name=file_a_name,
        file_a_content=file_a_content,
        file_b_name=file_b_name,
        file_b_content=file_b_content,
    )