"""
reco.py — Hand two ledgers to Claude, let it reconcile them freely, return the report.

No rigid schema, no forced format. Claude acts as the CA expert and decides
everything about how to reconcile and how to present the result.
"""

from __future__ import annotations
import base64
import csv as _csv
import io
import os
from dataclasses import dataclass, field

import pandas as pd
from anthropic import Anthropic

from prompt import SYSTEM_PROMPT, TASK_INSTRUCTION
from writer import report_to_excel

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
MODELS = {
    "Haiku (fastest, lowest cost)":      {"id": "claude-haiku-4-5-20251001", "in_per_m": 1.0,  "out_per_m": 5.0},
    "Sonnet (balanced — recommended)":   {"id": "claude-sonnet-4-6",         "in_per_m": 3.0,  "out_per_m": 15.0},
    "Opus (most capable, highest cost)": {"id": "claude-opus-4-8",           "in_per_m": 15.0, "out_per_m": 75.0},
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
# Turn an uploaded file into Claude message content block(s).
#   • Excel / CSV  → converted to plain text (Claude reads the structure)
#   • PDF          → sent natively as a document block (Claude reads it directly)
# ---------------------------------------------------------------------------
def ledger_blocks(label: str, file_bytes: bytes, filename: str) -> list[dict]:
    header = {"type": "text", "text": f"===== {label}: {filename} ====="}
    footer = {"type": "text", "text": f"===== END {label} ====="}

    if filename.lower().endswith(".pdf"):
        body = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": base64.standard_b64encode(file_bytes).decode("ascii"),
            },
        }
    else:
        body = {"type": "text", "text": file_to_text(file_bytes, filename)}

    return [header, body, footer]

def file_to_text(file_bytes: bytes, filename: str) -> str:
    if filename.lower().endswith((".csv", ".tsv", ".txt")):
        return _csv_to_text(file_bytes)
    return _excel_to_text(file_bytes)

def _decode(b):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    return b.decode("utf-8", errors="replace")

def _csv_to_text(b: bytes) -> str:
    text = _decode(b)
    delim = ","
    try:
        delim = _csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
    except Exception:
        pass
    rows = [r for r in _csv.reader(io.StringIO(text), delimiter=delim)
            if any(c.strip() for c in r)]
    buf = io.StringIO()
    _csv.writer(buf).writerows(rows)
    return buf.getvalue()

def _excel_to_text(b: bytes) -> str:
    last = None
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
# Result
# ---------------------------------------------------------------------------
@dataclass
class RecoResult:
    report: str            # Claude's reconciliation, verbatim (markdown)
    excel_bytes: bytes     # the same report, rendered to a neat Excel
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

    content = (
        [{"type": "text", "text": TASK_INSTRUCTION}]
        + ledger_blocks("LEDGER 1", file_a_bytes, file_a_name)
        + ledger_blocks("LEDGER 2", file_b_bytes, file_b_name)
    )

    msg = client.messages.create(
        model=model["id"],
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}],
    )
    report = "".join(b.text for b in msg.content if b.type == "text").strip()

    return RecoResult(
        report=report,
        excel_bytes=report_to_excel(report),
        model_label=model_label,
        input_tokens=msg.usage.input_tokens,
        output_tokens=msg.usage.output_tokens,
        usd_inr=usd_inr,
        pricing={"in_per_m": model["in_per_m"], "out_per_m": model["out_per_m"]},
    )
