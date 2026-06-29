"""
prompt.py — The only thing that matters. Edit this to change how the reco works.
"""

RECO_PROMPT = """
You are a senior Chartered Accountant. I have uploaded two ledger files for the same counterparty relationship.

Reconcile them completely — like you would in real CA practice. Figure out the structure yourself, match entries, find differences, and tell me everything.

===== LEDGER A: {file_a_name} =====
{file_a_content}
===== END LEDGER A =====

===== LEDGER B: {file_b_name} =====
{file_b_content}
===== END LEDGER B =====

Return your answer in this exact format — nothing before or after:

<summary>
Write 4-6 plain English sentences summarising the reconciliation. Name the parties, state the period, mention the closing balance difference, and call out the most important findings (unmatched invoices, TDS differences, etc.).
</summary>

<sheets>
[
  {{
    "name": "Summary",
    "rows": [
      ["Party A", "Techsmith Software Pvt Ltd"],
      ["Party B", "BigC Mobiles Pvt Ltd"],
      ["Period", "Apr 2025 – Jun 2026"],
      ["Closing Balance Diff", -81557.52],
      ["Total Matched", 3],
      ["Total Unmatched", 2]
    ]
  }},
  {{
    "name": "Matched Entries",
    "rows": [
      ["Invoice No", "Date (A)", "Date (B)", "Amount", "Narration"],
      ["INV-001", "24-Apr-2025", "24-Apr-2025", 1725384, "HDFC NEFT payment"]
    ]
  }},
  {{
    "name": "Differences",
    "rows": [
      ["Invoice No", "Amount (A)", "Amount (B)", "Difference", "Reason"],
      ["GS-140", 1694893, 1666166, 28727, "TDS deducted by Party B"]
    ]
  }},
  {{
    "name": "Only in A",
    "rows": [
      ["Date", "Voucher", "Invoice", "Debit", "Credit", "Narration"],
      ["25-Apr-2025", "JV-986", "", 0.84, 0, "Round off"]
    ]
  }},
  {{
    "name": "Only in B",
    "rows": [
      ["Date", "Voucher", "Invoice", "Debit", "Credit", "Narration"],
      ["01-Mar-2026", "TCV-6206", "GS-211", 0, 29206.18, "IT Consultancy"]
    ]
  }}
]
</sheets>

Rules:
- The example above shows the FORMAT only. Use whatever sheets and columns make sense for this actual data.
- You decide how many sheets, what to name them, what columns to include.
- First row of every sheet must be the column headers.
- All monetary values must be plain numbers (no ₹ symbol, no commas). Dates as strings.
- Sheet names: no special characters / \\ ? * [ ] :
- The JSON inside <sheets> must be valid JSON.
- The <summary> is plain text only, no JSON.
"""


def build_prompt(file_a_name, file_a_content, file_b_name, file_b_content):
    return RECO_PROMPT.format(
        file_a_name=file_a_name,
        file_a_content=file_a_content,
        file_b_name=file_b_name,
        file_b_content=file_b_content,
    )