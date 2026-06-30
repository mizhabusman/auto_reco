"""
prompt.py — Asks Claude for the EXACT 7-sheet reconciliation format.

Output schema is rigid (matches the example template):
  Summary, TDS Reconciliation, Matched, Amount Mismatches,
  Missing in Their Books, Missing in Our Books, Timing Differences
"""

RECO_PROMPT = """You are a senior Chartered Accountant performing a vendor / counterparty
ledger reconciliation between two parties' books.

Two ledger files are provided below. Ledger A = "Our books" (the entity
preparing the reconciliation). Ledger B = "Their books" (the counterparty).

===== LEDGER A (OUR BOOKS): {file_a_name} =====
{file_a_content}
===== END LEDGER A =====

===== LEDGER B (THEIR BOOKS): {file_b_name} =====
{file_b_content}
===== END LEDGER B =====

============================================================
HOW TO RECONCILE
============================================================
1. Figure out the structure of each ledger yourself — ignore company-name
   junk rows, find the real header row, identify date/voucher/amount/
   invoice/TDS columns.

2. Match transactions across both books at three confidence levels:
   - L1 (strongest)  : Date + Invoice/Ref + Amount all agree
   - L2 (timing)     : Invoice/Ref + Amount agree, dates differ
   - L3 (weak)       : Date + Amount agree (no ref to match on)

3. Categorize every entry into ONE of these buckets:
   - Matched               (any L1/L2/L3 match)
   - Amount Mismatches     (same ref but amounts differ > ₹1)
   - Missing in Their Books (in our books, no counterpart in theirs)
   - Missing in Our Books  (in their books, no counterpart in ours)
   - Timing Differences    (subset of L2 matches — show separately)

4. TDS treatment — separate journal entries (description contains "TDS"
   or "withholding") must be detected and reported separately, NOT
   counted as transactions.

5. Generate a unique Rec ID for every matched/mismatched entry in the
   form "REC-XXXXXX" (6 random hex characters, uppercase).

============================================================
OUTPUT — RETURN EXACTLY THIS, NOTHING ELSE
============================================================

<summary>
Write 4-6 plain English sentences naming the parties, period, key
findings (closing balance difference, TDS gap, biggest unmatched items).
</summary>

<reco>
{{
  "stats": {{
    "total_ours": <int>,
    "total_theirs": <int>,
    "l1_matches": <int>,
    "l2_matches": <int>,
    "l3_matches": <int>,
    "amount_mismatches": <int>,
    "missing_their_books": <int>,
    "missing_our_books": <int>,
    "tds_journal_entries": <int>
  }},
  "balances": {{
    "opening_ours": <number>,
    "sum_ours": <number>,
    "closing_ours": <number>,
    "opening_theirs": <number>,
    "sum_theirs": <number>,
    "closing_theirs": <number>,
    "difference": <number>,
    "reconciling_items": <number>
  }},
  "tds": {{
    "status": "CLEAR" | "PARTIAL" | "EXCESS" | "MISMATCH",
    "message": "Plain English banner text — e.g. 'Partial TDS posting: counterparty withheld ₹62,316.00, we have booked only ₹15,278.00. Gap: ₹47,038.00 not yet posted by us.'",
    "our_tds_column_total": <number>,
    "their_tds_column_total": <number>,
    "our_tds_journal_total": <number>,
    "their_tds_journal_total": <number>,
    "our_col_vs_their_journal": <number>,
    "their_col_vs_our_journal": <number>,
    "journal_entries": [
      {{ "source": "Ours" | "Theirs", "date": "YYYY-MM-DD", "voucher_no": "...",
         "description": "...", "amount": <number>, "status": "PARTIAL"|"CLEAR"|"EXCESS" }}
    ]
  }},
  "matched": [
    {{ "rec_id": "REC-XXXXXX", "match_level": "L1"|"L2"|"L3",
       "our_date": "YYYY-MM-DD", "their_date": "YYYY-MM-DD",
       "invoice_ref": "...", "our_description": "...", "their_description": "...",
       "our_amount": <number>, "their_amount": <number>, "tds_amount": <number> }}
  ],
  "amount_mismatches": [
    {{ "rec_id": "REC-XXXXXX", "date": "YYYY-MM-DD", "invoice_ref": "...",
       "description": "...", "our_amount": <number>, "their_amount": <number>,
       "difference": <number> }}
  ],
  "missing_their_books": [
    {{ "date": "YYYY-MM-DD", "voucher_type": "...", "voucher_no": "...",
       "invoice_ref": "...", "description": "...", "gross_amount": <number>,
       "tds_amount": <number>, "net_amount": <number>, "note": "..." }}
  ],
  "missing_our_books": [
    {{ "date": "YYYY-MM-DD", "voucher_type": "...", "voucher_no": "...",
       "invoice_ref": "...", "description": "...", "gross_amount": <number>,
       "tds_amount": <number>, "net_amount": <number>, "note": "..." }}
  ],
  "timing_differences": [
    {{ "rec_id": "REC-XXXXXX", "match_level": "L2",
       "our_date": "YYYY-MM-DD", "their_date": "YYYY-MM-DD",
       "invoice_ref": "...", "our_description": "...", "their_description": "...",
       "our_amount": <number>, "their_amount": <number>, "days_diff": <int> }}
  ]
}}
</reco>

RULES:
- Return ONLY the two tagged blocks. No prose, no markdown fences.
- JSON must be valid. Use 0 for missing numbers, "" for missing strings.
- Every amount is a JSON number (not string).
- Dates as "YYYY-MM-DD" strings.
- If a category has no entries, return an empty array [].
- "Timing Differences" is a SUBSET of "Matched" (the L2 entries) —
  include them in both arrays.
- TDS journal entries should NOT appear in matched/mismatched arrays —
  only in tds.journal_entries.
"""


def build_prompt(file_a_name, file_a_content, file_b_name, file_b_content):
    return RECO_PROMPT.format(
        file_a_name=file_a_name,
        file_a_content=file_a_content,
        file_b_name=file_b_name,
        file_b_content=file_b_content,
    )