"""
prompts.py — The only "logic" in the whole app.

Call 1 mimics pasting two ledgers into a Claude chat and saying
"you're a CA — reconcile this". No format is imposed, so Claude reconciles
at full quality, exactly like in chat.

Call 2 takes the finished reconciliation and asks Claude to lay the same
findings out as an Excel workbook (sheets/columns of its own choosing).
"""


def reconcile_prompt(a_name, a_text, b_name, b_text):
    return f"""You are an expert Chartered Accountant with deep experience in \
ledger and party reconciliation.

Below are two ledgers from two different parties.

===== LEDGER A: {a_name} =====
{a_text}
===== END OF LEDGER A =====

===== LEDGER B: {b_name} =====
{b_text}
===== END OF LEDGER B =====

Reconcile these two ledgers.

Present your complete reconciliation as a well-organised markdown report — \
exactly as you would if these files were handed to you in a conversation. \
You decide the structure, the depth, and what matters for THIS pair of books. \
Be thorough: match what matches, dig into every difference (amounts, \
TDS / withholding, timing, missing entries, rounding), and point out anything \
unusual you notice — including things the parties themselves may have missed. \
Use headings and tables where they help.

When you show whether items match, lead the status with a small indicator so \
the report is easy to scan: ✅ for items that agree, ⚠️ for minor timing / \
rounding items to review, and ❌ for unmatched, missing, or disputed items. \
Keep this restrained and professional — a single marker per status, no heavy \
colour blocks. End with your professional conclusion."""


def workbook_prompt(a_name, a_text, b_name, b_text, report):
    return f"""A Chartered Accountant has reconciled two ledgers. Below are the \
two original ledgers and the finished reconciliation report.

===== LEDGER A: {a_name} =====
{a_text}
===== END OF LEDGER A =====

===== LEDGER B: {b_name} =====
{b_text}
===== END OF LEDGER B =====

===== RECONCILIATION REPORT =====
{report}
===== END OF REPORT =====

Convert this reconciliation into an Excel workbook. Return ONLY a JSON object \
in exactly this shape, and nothing else:

{{
  "sheets": [
    {{
      "name": "<short sheet name you choose>",
      "columns": ["<column headers you choose>"],
      "rows": [ ["<cell>", 123.45, "..."], ["..."] ]
    }}
  ]
}}

Rules:
- You decide the sheets, their names, and their columns — whatever presents
  this reconciliation best in Excel (a summary, matched items, differences,
  entries missing on one side, TDS, timing — as applicable to the report).
- Include every individual line item from the reconciliation, not just totals.
- Each row is an array of cells aligned with "columns".
- Amounts must be JSON numbers (not strings). Use "" for blank cells.
- Dates as "YYYY-MM-DD" strings.
- Valid JSON only: no markdown fences, no comments, no trailing commas."""
