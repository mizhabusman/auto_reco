"""
reco.py — Call Claude, get data back, write Excel with pandas. That's it.
"""

from __future__ import annotations
import csv as _csv
import io
import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd
from anthropic import Anthropic
from prompt import build_prompt
from writer import reco_to_excel

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
MODELS = {
    "Haiku (fastest, lowest cost)":     {"id": "claude-haiku-4-5-20251001", "in_per_m": 1.0,  "out_per_m": 5.0},
    "Sonnet (balanced — recommended)":  {"id": "claude-sonnet-4-6",         "in_per_m": 3.0,  "out_per_m": 15.0},
    "Opus (most capable, highest cost)":{"id": "claude-opus-4-8",           "in_per_m": 15.0, "out_per_m": 75.0},
}
DEFAULT_MODEL   = "Sonnet (balanced — recommended)"
DEFAULT_USD_INR = 85.0
MAX_TOKENS      = 16000

MODEL_NAMES    = ["Haiku", "Sonnet", "Opus"]
MODEL_CAPTIONS = ["Low cost · Fast", "Balanced · Recommended", "High cost · Most capable"]
NAME_TO_LABEL  = {
    "Haiku":  "Haiku (fastest, lowest cost)",
    "Sonnet": "Sonnet (balanced — recommended)",
    "Opus":   "Opus (most capable, highest cost)",
}
DEFAULT_INDEX = 1

# ---------------------------------------------------------------------------
# Read uploaded file → plain text (Claude reads the structure)
# ---------------------------------------------------------------------------
def file_to_text(file_bytes: bytes, filename: str) -> str:
    if filename.lower().endswith((".csv", ".tsv", ".txt")):
        return _csv_to_text(file_bytes)
    return _excel_to_text(file_bytes)

def _decode(b):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try: return b.decode(enc)
        except: pass
    return b.decode("utf-8", errors="replace")

def _csv_to_text(b: bytes) -> str:
    text = _decode(b)
    delim = ","
    try:
        d = _csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        delim = d.delimiter
    except: pass
    rows = [r for r in _csv.reader(io.StringIO(text), delimiter=delim)
            if any(c.strip() for c in r)]
    buf = io.StringIO()
    _csv.writer(buf).writerows(rows)
    return f"## CSV\n{buf.getvalue()}"

def _excel_to_text(b: bytes) -> str:
    for engine in ("openpyxl", "xlrd", None):
        try:
            xl = pd.ExcelFile(io.BytesIO(b), engine=engine) if engine else pd.ExcelFile(io.BytesIO(b))
            parts = []
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, header=None, dtype=str, keep_default_na=False)
                df = df.replace("", pd.NA).dropna(how="all").dropna(axis=1, how="all").fillna("")
                parts.append(f"## Sheet: {sheet}\n{df.to_csv(index=False, header=False)}")
            return "\n\n".join(parts)
        except Exception as e:
            last = e
    raise RuntimeError(f"Could not read file: {last}")

# ---------------------------------------------------------------------------
# Parse Claude's response
# ---------------------------------------------------------------------------
def _extract(text: str, tag: str) -> str:
    m = re.search(rf"<{tag}>(.*?)</{tag}>", text, re.DOTALL)
    return m.group(1).strip() if m else ""

def _parse_reco(raw: str) -> dict:
    """Parse the <reco>...</reco> JSON block. Robust to fences/whitespace."""
    txt = _extract(raw, "reco")
    if not txt:
        return {"stats": {}, "balances": {}, "tds": {}, "matched": [],
                "amount_mismatches": [], "missing_their_books": [],
                "missing_our_books": [], "timing_differences": []}

    txt = re.sub(r"^```[a-zA-Z]*\n?", "", txt.strip())
    txt = re.sub(r"\n?```$", "", txt.strip())

    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        start, end = txt.find("{"), txt.rfind("}")
        if start != -1 and end != -1:
            try:
                return json.loads(txt[start:end+1])
            except json.JSONDecodeError:
                pass
    # last resort: empty skeleton
    return {"stats": {}, "balances": {}, "tds": {}, "matched": [],
            "amount_mismatches": [], "missing_their_books": [],
            "missing_our_books": [], "timing_differences": []}

# ---------------------------------------------------------------------------
# Write Excel — pandas ExcelWriter, one sheet per item, no styling
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
@dataclass
class RecoResult:
    excel_bytes: bytes
    summary: str
    sheets: list[dict]
    raw_response: str
    model_label: str
    input_tokens: int
    output_tokens: int
    usd_inr: float = DEFAULT_USD_INR
    pricing: dict = field(default_factory=dict)

    @property
    def cost_usd(self):
        return (self.input_tokens * self.pricing.get("in_per_m", 0)
              + self.output_tokens * self.pricing.get("out_per_m", 0)) / 1_000_000

    @property
    def cost_inr(self):
        return self.cost_usd * self.usd_inr

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
def run_reconciliation(
    file_a_bytes, file_a_name,
    file_b_bytes, file_b_name,
    model_label=DEFAULT_MODEL,
    usd_inr=DEFAULT_USD_INR,
    api_key=None,
) -> RecoResult:
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    model  = MODELS[model_label]
    client = Anthropic(api_key=api_key)

    prompt = build_prompt(
        file_a_name, file_to_text(file_a_bytes, file_a_name),
        file_b_name, file_to_text(file_b_bytes, file_b_name),
    )

    msg = client.messages.create(
        model=model["id"], max_tokens=MAX_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = "".join(b.text for b in msg.content if b.type == "text")

    summary    = _extract(raw, "summary")
    reco_data  = _parse_reco(raw)
    excel      = reco_to_excel(reco_data)

    return RecoResult(
        excel_bytes=excel,
        summary=summary,
        sheets=reco_data,
        raw_response=raw,
        model_label=model_label,
        input_tokens=msg.usage.input_tokens,
        output_tokens=msg.usage.output_tokens,
        usd_inr=usd_inr,
        pricing={"in_per_m": model["in_per_m"], "out_per_m": model["out_per_m"]},
    )