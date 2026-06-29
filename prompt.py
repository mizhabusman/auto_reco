"""
prompt.py — The entire brain of the reconciliation.

Edit this file to change how Claude reconciles, what the Excel looks like,
what summary info is shown on screen, or how matching works.
Nothing else needs to change.
"""

RECO_PROMPT = """
You are a senior Chartered Accountant. I am giving you two ledger exports from
two companies for the same counterparty relationship.

Your job: perform a complete, professional reconciliation and produce a
downloadable Excel report in whatever format you think is best for this data.

===== LEDGER A: {file_a_name} =====
{file_a_content}
===== END LEDGER A =====

===== LEDGER B: {file_b_name} =====
{file_b_content}
===== END LEDGER B =====

============================================================
WHAT TO DO
============================================================
1. Figure out the structure of each ledger yourself — ignore company-name
   junk rows at the top, find the real header row, identify date/voucher/
   amount/invoice columns by reading the content.

2. Reconcile like a CA would:
   - Ledgers are mirrors: a sale in one = purchase in the other.
   - Match on invoice/bill number (fuzzy — "25-26/GS-140" = "GS-140").
   - One side may book gross, the other net of TDS — treat as a match.
   - Flag amount differences, timing differences, entries only in one book.
   - Compute opening/closing balances and the difference between the two books.

3. Write Python code (using openpyxl) that creates a professional Excel
   reconciliation report. You decide the sheets, columns, colors and layout —
   whatever communicates the reconciliation most clearly. The code must:
   - import openpyxl and any stdlib modules it needs (no other libraries)
   - end with:  excel_bytes = <io.BytesIO with the saved workbook>.getvalue()
   - NOT call wb.save() to a file path — only save to a BytesIO buffer.

4. Write a short summary of the reconciliation for display on screen.

============================================================
RESPONSE FORMAT — return EXACTLY this, nothing else:
============================================================

<summary>
{{
  "party_a": "name of party A",
  "party_b": "name of party B",
  "period": "date range",
  "closing_diff": <number — difference between closing balances, 0 if matched>,
  "sheet_count": <number of sheets in your Excel>,
  "insights": [
    "specific finding 1 — name invoice numbers and amounts",
    "specific finding 2",
    "specific finding 3"
  ]
}}
</summary>

<excel_code>
# your complete, runnable Python code here
# must end with: excel_bytes = buf.getvalue()
</excel_code>

Rules:
- Return ONLY the two tagged blocks above. No prose before or after.
- The JSON inside <summary> must be valid JSON.
- The code inside <excel_code> must run without errors using only
  openpyxl and Python stdlib (io, datetime, etc.).
- Sheet names must not contain / \\ ? * [ ] or : (use - instead).
- All monetary values in the Excel should be numbers, not strings.
"""


def build_prompt(file_a_name, file_a_content, file_b_name, file_b_content):
    return RECO_PROMPT.format(
        file_a_name=file_a_name,
        file_a_content=file_a_content,
        file_b_name=file_b_name,
        file_b_content=file_b_content,
    )