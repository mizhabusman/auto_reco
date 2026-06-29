"""
reco.py — Read files, call Claude, parse response, exec the Excel code.

Three things come out:
  - excel_bytes  : the .xlsx file Claude produced
  - summary      : dict for the web UI (party names, insights, etc.)
  - token usage  : for cost display
"""

from __future__ import annotations
import csv as _csv
import io
import json
import os
import re
import textwrap
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from anthropic import Anthropic

from prompt import build_prompt

# ---------------------------------------------------------------------------
# Model catalogue  (edit pricing here if rates change)
# ---------------------------------------------------------------------------
MODELS: dict[str, dict[str, Any]] = {
    "Haiku (fastest, lowest cost)": {
        "id": "claude-haiku-4-5-20251001",
        "tier": "Low cost · Fast",
        "in_per_m": 1.0,
        "out_per_m": 5.0,
    },
    "Sonnet (balanced — recommended)": {
        "id": "claude-sonnet-4-6",
        "tier": "Balanced · Accurate",
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
DEFAULT_USD_INR = 85.0
MAX_OUTPUT_TOKENS = 16000

MODEL_NAMES    = ["Haiku", "Sonnet", "Opus"]
MODEL_CAPTIONS = ["Low cost · Fast", "Balanced · Recommended", "High cost · Most capable"]
NAME_TO_LABEL  = {
    "Haiku":  "Haiku (fastest, lowest cost)",
    "Sonnet": "Sonnet (balanced — recommended)",
    "Opus":   "Opus (most capable, highest cost)",
}
DEFAULT_INDEX = 1


# ---------------------------------------------------------------------------
# File → raw text (Claude reads the structure; we just dump bytes as text)
# ---------------------------------------------------------------------------
def file_to_text(file_bytes: bytes, filename: str) -> str:
    name = filename.lower()
    if name.endswith((".csv", ".tsv", ".txt")):
        return _csv_bytes_to_text(file_bytes)
    return _excel_bytes_to_text(file_bytes)


def _decode(b: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def _csv_bytes_to_text(file_bytes: bytes) -> str:
    text = _decode(file_bytes)
    delim = ","
    try:
        dialect = _csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        delim = dialect.delimiter
    except Exception:
        pass
    rows = [r for r in _csv.reader(io.StringIO(text), delimiter=delim)
            if any((c or "").strip() for c in r)]
    buf = io.StringIO()
    _csv.writer(buf).writerows(rows)
    return f"## Sheet: CSV\n{buf.getvalue()}"


def _excel_bytes_to_text(file_bytes: bytes) -> str:
    last_err = None
    for engine in ("openpyxl", "xlrd", None):
        try:
            buf = io.BytesIO(file_bytes)
            xl = pd.ExcelFile(buf, engine=engine) if engine else pd.ExcelFile(buf)
            parts = []
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, header=None, dtype=str, keep_default_na=False)
                df = df.replace("", pd.NA).dropna(how="all").dropna(axis=1, how="all").fillna("")
                parts.append(f"## Sheet: {sheet}\n{df.to_csv(index=False, header=False)}")
            return "\n\n".join(parts)
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Could not read spreadsheet — {last_err}")


# ---------------------------------------------------------------------------
# Parse Claude's response  (<summary>…</summary>  <excel_code>…</excel_code>)
# ---------------------------------------------------------------------------
def _extract_tag(text: str, tag: str) -> str:
    m = re.search(rf"<{tag}>(.*?)</{tag}>", text, re.DOTALL)
    return m.group(1).strip() if m else ""


def _parse_summary(raw: str) -> dict:
    txt = _extract_tag(raw, "summary")
    if not txt:
        return {"party_a": "Party A", "party_b": "Party B",
                "period": "", "closing_diff": 0, "sheet_count": 1, "insights": []}
    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        # strip any stray backticks
        txt = re.sub(r"```.*?```", "", txt, flags=re.DOTALL).strip()
        return json.loads(txt)


def _exec_excel_code(code: str) -> bytes:
    """Run Claude's openpyxl code in a clean namespace; return excel_bytes."""
    ns: dict[str, Any] = {}
    exec(textwrap.dedent(code), ns)          # noqa: S102
    excel_bytes = ns.get("excel_bytes")
    if not isinstance(excel_bytes, bytes):
        raise RuntimeError(
            "Claude's code ran but did not produce `excel_bytes`. "
            "Raw code:\n" + code[:400]
        )
    return excel_bytes


# ---------------------------------------------------------------------------
# Result dataclass
# ---------------------------------------------------------------------------
@dataclass
class RecoResult:
    excel_bytes: bytes
    summary: dict[str, Any]
    raw_response: str
    model_label: str
    input_tokens: int
    output_tokens: int
    usd_inr: float = DEFAULT_USD_INR
    pricing: dict[str, float] = field(default_factory=dict)

    @property
    def cost_usd(self) -> float:
        return (self.input_tokens  * self.pricing.get("in_per_m",  0)
              + self.output_tokens * self.pricing.get("out_per_m", 0)) / 1_000_000

    @property
    def cost_inr(self) -> float:
        return self.cost_usd * self.usd_inr


# ---------------------------------------------------------------------------
# Main entry point
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
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    model = MODELS[model_label]
    prompt = build_prompt(
        file_a_name, file_to_text(file_a_bytes, file_a_name),
        file_b_name, file_to_text(file_b_bytes, file_b_name),
    )

    client = Anthropic(api_key=api_key)
    msg = client.messages.create(
        model=model["id"],
        max_tokens=MAX_OUTPUT_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = "".join(b.text for b in msg.content if b.type == "text")

    summary     = _parse_summary(raw)
    excel_code  = _extract_tag(raw, "excel_code")
    excel_bytes = _exec_excel_code(excel_code)

    return RecoResult(
        excel_bytes=excel_bytes,
        summary=summary,
        raw_response=raw,
        model_label=model_label,
        input_tokens=msg.usage.input_tokens,
        output_tokens=msg.usage.output_tokens,
        usd_inr=usd_inr,
        pricing={"in_per_m": model["in_per_m"], "out_per_m": model["out_per_m"]},
    )