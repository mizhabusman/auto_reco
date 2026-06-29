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
from typing import Any

import pandas as pd
from anthropic import Anthropic
from prompt import build_prompt

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

def _parse_sheets(raw: str) -> list[dict]:
    txt = _extract(raw, "sheets")
    if not txt:
        return []
    # strip any accidental code fences
    txt = re.sub(r"^```[a-z]*\n?", "", txt.strip())
    txt = re.sub(r"\n?```$", "", txt.strip())
    return json.loads(txt)

# ---------------------------------------------------------------------------
# Write Excel — pandas ExcelWriter, one sheet per item, no styling
# ---------------------------------------------------------------------------
def sheets_to_excel(sheets: list[dict]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        seen = {}
        for sh in sheets:
            name = str(sh.get("name", "Sheet"))
            # sanitise sheet name
            for ch in r'/\?*[]:\x00': name = name.replace(ch, "-")
            name = name[:31].strip() or "Sheet"
            # deduplicate
            seen[name] = seen.get(name, 0) + 1
            if seen[name] > 1:
                name = name[:28] + f" {seen[name]}"

            rows = sh.get("rows", [])
            if not rows:
                pd.DataFrame().to_excel(writer, sheet_name=name, index=False)
                continue
            df = pd.DataFrame(rows[1:], columns=rows[0])
            df.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()

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

    summary = _extract(raw, "summary")
    sheets  = _parse_sheets(raw)
    excel   = sheets_to_excel(sheets)

    return RecoResult(
        excel_bytes=excel,
        summary=summary,
        sheets=sheets,
        raw_response=raw,
        model_label=model_label,
        input_tokens=msg.usage.input_tokens,
        output_tokens=msg.usage.output_tokens,
        usd_inr=usd_inr,
        pricing={"in_per_m": model["in_per_m"], "out_per_m": model["out_per_m"]},
    )