"""
reco.py — Reads raw files, calls Claude, parses the flexible workbook JSON.

Also owns:
  • the model catalogue (id, label, cost tier, pricing)
  • token usage + cost calculation in INR
"""

from __future__ import annotations
import io
import json
import os
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from anthropic import Anthropic

from prompt import build_prompt

# ---------------------------------------------------------------------------
# Model catalogue
# ---------------------------------------------------------------------------
# Pricing is USD per 1,000,000 tokens (input, output).
# >>> VERIFY against https://docs.claude.com pricing and adjust if needed. <<<
MODELS: dict[str, dict[str, Any]] = {
    "Haiku (fastest, lowest cost)": {
        "id": "claude-haiku-4-5-20251001",
        "tier": "Low cost · Fast",
        "in_per_m": 1.0,
        "out_per_m": 5.0,
    },
    "Sonnet (balanced — recommended)": {
        "id": "claude-sonnet-4-6",
        "tier": "Balanced cost · Accurate",
        "in_per_m": 3.0,
        "out_per_m": 15.0,
    },
    "Opus (most capable, highest cost)": {
        "id": "claude-opus-4-8",
        "tier": "High cost · Most capable",
        "in_per_m": 15.0,
        "out_per_m": 75.0,
    },
}
DEFAULT_MODEL = "Sonnet (balanced — recommended)"

# USD -> INR. Editable in the UI; this is just the default.
DEFAULT_USD_INR = 85.0

MAX_OUTPUT_TOKENS = 16000


# ---------------------------------------------------------------------------
# Raw file -> text dump (no cleaning; Claude handles structure)
# Robust to: ragged CSV rows, odd delimiters, wrong file extensions,
# and .xls files that are actually .xlsx underneath (common in Tally exports).
# ---------------------------------------------------------------------------
import csv as _csv


def file_to_text(file_bytes: bytes, filename: str) -> str:
    name = filename.lower()
    if name.endswith((".csv", ".tsv", ".txt")):
        return _csv_to_text(file_bytes)
    return _excel_to_text(file_bytes)


def _decode(b: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def _csv_to_text(file_bytes: bytes) -> str:
    """Parse with the tolerant `csv` module (handles ragged rows that break
    pandas' C parser), drop fully-empty rows, re-emit clean CSV text."""
    text = _decode(file_bytes)
    # sniff delimiter from a sample; fall back to comma
    delim = ","
    try:
        dialect = _csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        delim = dialect.delimiter
    except Exception:
        pass
    rows = []
    for row in _csv.reader(io.StringIO(text), delimiter=delim):
        if any((c or "").strip() for c in row):
            rows.append(row)
    out = io.StringIO()
    _csv.writer(out).writerows(rows)
    return f"## Sheet: CSV\n{out.getvalue()}"


def _excel_to_text(file_bytes: bytes) -> str:
    """Try engines in order so a mislabeled file (e.g. .xls that is really
    .xlsx, or vice-versa) still reads."""
    last_err = None
    for engine in ("openpyxl", "xlrd", None):
        try:
            buf = io.BytesIO(file_bytes)
            xl = pd.ExcelFile(buf, engine=engine) if engine else pd.ExcelFile(buf)
            parts = []
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, header=None, dtype=str, keep_default_na=False)
                parts.append(_df_to_csv_text(df, sheet))
            return "\n\n".join(parts)
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
    raise RuntimeError(
        f"Could not read this spreadsheet (tried xlsx and xls readers). "
        f"Last error: {last_err}"
    )


def _df_to_csv_text(df: pd.DataFrame, sheet_name: str) -> str:
    df = df.replace("", pd.NA).dropna(how="all").dropna(axis=1, how="all").fillna("")
    csv = df.to_csv(index=False, header=False)
    return f"## Sheet: {sheet_name}\n{csv}"


# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
@dataclass
class RecoResult:
    data: dict[str, Any]
    raw_response: str
    model_label: str
    input_tokens: int
    output_tokens: int
    usd_inr: float = DEFAULT_USD_INR
    pricing: dict[str, float] = field(default_factory=dict)

    @property
    def cost_usd(self) -> float:
        return (self.input_tokens * self.pricing.get("in_per_m", 0)
                + self.output_tokens * self.pricing.get("out_per_m", 0)) / 1_000_000

    @property
    def cost_inr(self) -> float:
        return self.cost_usd * self.usd_inr


# ---------------------------------------------------------------------------
# Claude call
# ---------------------------------------------------------------------------
def run_reconciliation(
    file_a_bytes: bytes, file_a_name: str,
    file_b_bytes: bytes, file_b_name: str,
    model_label: str = DEFAULT_MODEL,
    usd_inr: float = DEFAULT_USD_INR,
    api_key: str | None = None,
) -> RecoResult:
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY not set")

    model = MODELS[model_label]

    file_a_text = file_to_text(file_a_bytes, file_a_name)
    file_b_text = file_to_text(file_b_bytes, file_b_name)
    prompt = build_prompt(file_a_name, file_a_text, file_b_name, file_b_text)

    client = Anthropic(api_key=api_key)
    msg = client.messages.create(
        model=model["id"],
        max_tokens=MAX_OUTPUT_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = "".join(b.text for b in msg.content if b.type == "text")
    data = _parse_json(raw)

    return RecoResult(
        data=data,
        raw_response=raw,
        model_label=model_label,
        input_tokens=msg.usage.input_tokens,
        output_tokens=msg.usage.output_tokens,
        usd_inr=usd_inr,
        pricing={"in_per_m": model["in_per_m"], "out_per_m": model["out_per_m"]},
    )


def _parse_json(text: str) -> dict[str, Any]:
    t = text.strip()
    if t.startswith("```"):
        t = t.split("\n", 1)[1] if "\n" in t else t
        if t.endswith("```"):
            t = t.rsplit("```", 1)[0]
        t = t.strip()
        if t.startswith("json"):
            t = t[4:].lstrip()
    if not t.startswith("{"):
        i, j = t.find("{"), t.rfind("}")
        if i != -1 and j != -1:
            t = t[i:j + 1]
    return json.loads(t)