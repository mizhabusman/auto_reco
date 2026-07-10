"""
engine.py — The whole backend.

Reads the two uploaded ledgers into plain text, hands them to Claude
("you are a CA — reconcile these"), streams the report back live, then makes
a second call asking Claude to lay the same reconciliation out as an Excel
workbook of its own design.

No reconciliation logic lives in Python. Claude decides everything.
"""
from __future__ import annotations

import csv as _csv
import io
import json
import re

import anthropic
import pandas as pd
from anthropic import Anthropic

from excel_writer import workbook_to_excel
from prompts import reconcile_prompt, workbook_prompt

USD_INR = 85.0
MAX_TOKENS = 32000

# thinking=True → adaptive thinking (supported on Sonnet 4.6 / Opus 4.8;
# Haiku 4.5 does not support adaptive thinking, so it runs without).
MODELS = {
    "Haiku":  {"id": "claude-haiku-4-5",  "caption": "Fastest · lowest cost",       "in_per_m": 1.0, "out_per_m": 5.0,  "thinking": False},
    "Sonnet": {"id": "claude-sonnet-4-6", "caption": "Balanced · recommended",       "in_per_m": 3.0, "out_per_m": 15.0, "thinking": True},
    "Opus":   {"id": "claude-opus-4-8",   "caption": "Most capable · highest cost",  "in_per_m": 5.0, "out_per_m": 25.0, "thinking": True},
}
MODEL_NAMES = list(MODELS)
MODEL_CAPTIONS = [MODELS[k]["caption"] for k in MODELS]
DEFAULT_INDEX = 1  # Sonnet


# ---------------------------------------------------------------------------
# Friendly errors — the user only ever sees plain words, never a traceback
# ---------------------------------------------------------------------------
class RecoError(Exception):
    """A problem worth showing the user in plain English (message + hint)."""

    def __init__(self, message: str, hint: str = ""):
        super().__init__(message)
        self.message = message
        self.hint = hint


def describe_api_error(e: Exception) -> tuple[str, str]:
    """Turn an Anthropic SDK error into a plain (message, hint) pair."""
    body = str(getattr(e, "message", "") or e).lower()

    if isinstance(e, anthropic.AuthenticationError):
        return ("Your API key is not valid.",
                "Check ANTHROPIC_API_KEY in your .env file — it may be wrong, "
                "expired, or revoked.")
    if isinstance(e, anthropic.PermissionDeniedError):
        return ("Your API key doesn't have access to this.",
                "The key is valid but isn't allowed to use this model. "
                "Try another model, or use a different key.")
    if isinstance(e, anthropic.RateLimitError):
        if any(w in body for w in ("credit", "quota", "billing")):
            return ("Your account has run out of credits.",
                    "Add credits or check your plan at console.anthropic.com, "
                    "then try again.")
        return ("Too many requests right now (rate limit).",
                "Wait a few seconds and run it again.")
    if isinstance(e, anthropic.RequestTooLargeError):
        return ("The ledgers are too large to process in one request.",
                "Try smaller files, or split them and reconcile in parts.")
    if isinstance(e, anthropic.BadRequestError):
        if any(w in body for w in ("credit", "billing")):
            return ("Your account has run out of credits.",
                    "Add credits or check your plan at console.anthropic.com, "
                    "then try again.")
        return ("The reconciliation couldn't be processed.",
                "The files may be unreadable or too large. Try different files.")
    if isinstance(e, anthropic.NotFoundError):
        return ("The selected model isn't available on your account.",
                "Pick a different model and try again.")
    if isinstance(e, anthropic.OverloadedError):
        return ("The service is very busy right now.",
                "This is temporary — wait a moment and run it again.")
    if isinstance(e, anthropic.InternalServerError):
        return ("The service hit a temporary problem.",
                "Please run it again in a moment.")
    if isinstance(e, anthropic.APITimeoutError):
        return ("The request took too long and timed out.",
                "Try again, or pick the faster Haiku model for large files.")
    if isinstance(e, anthropic.APIConnectionError):
        return ("Couldn't reach the reconciliation service.",
                "Check your internet connection and try again.")
    return ("Something went wrong during reconciliation.",
            "Please try again in a moment.")


# ---------------------------------------------------------------------------
# Uploaded file → plain text (Claude reads the structure itself)
#
# Goal: extract EVERYTHING. No row, cell, or line is ever dropped during
# reading. Each reader captures the complete content of the file; the router
# picks the right reader by the file's real signature (not just its name),
# and never sends binary garbage to Claude. A scanned/image PDF (no text to
# read) is reported clearly instead of being silently mis-read.
# ---------------------------------------------------------------------------
class _NoTextPDF(Exception):
    """A PDF that opened fine but contains no extractable text (a scan/image)."""


def _sniff(b: bytes) -> str:
    """Identify the real format from the file's magic bytes."""
    head = b[:8]
    if head[:4] == b"%PDF":
        return "pdf"
    if head[:2] == b"PK":                      # zip container: xlsx / xlsm / docx
        return "zip"
    if head[:4] == b"\xD0\xCF\x11\xE0":        # OLE2 container: legacy xls / doc
        return "ole"
    return "text"


def _looks_like_text(b: bytes) -> bool:
    """True if the bytes are (mostly) decodable text — used to gate the raw
    fallback so we never hand binary junk to Claude as if it were a ledger."""
    sample = b[:8192]
    if not sample or b"\x00" in sample:
        return False
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            s = sample.decode(enc)
            break
        except UnicodeDecodeError:
            s = None
    if not s:
        return False
    printable = sum(ch.isprintable() or ch in "\r\n\t" for ch in s)
    return printable / len(s) > 0.85


def ledger_to_text(file_bytes: bytes, filename: str) -> str:
    name = (filename or "").lower()
    ext = name.rsplit(".", 1)[-1] if "." in name else ""
    kind = _sniff(file_bytes)

    by_ext = {
        "csv": _csv_to_text, "tsv": _csv_to_text, "txt": _csv_to_text,
        "xls": _excel_to_text, "xlsx": _excel_to_text,
        "xlsm": _excel_to_text, "xltx": _excel_to_text,
        "pdf": _pdf_to_text,
        "docx": _docx_to_text,
    }

    # Choose the reader by the file's true signature first (most reliable),
    # then by its extension, then the other structured parsers as a fallback
    # for mislabelled files. Every reader reads the whole file, so nothing is
    # lost. The plain-text reader is treated specially: it "succeeds" on ANY
    # bytes (it just decodes them), so it is only ever tried when the file
    # genuinely looks like text — otherwise binary junk would slip through.
    text_like = _looks_like_text(file_bytes)
    order: list = []
    if kind == "pdf":
        order = [_pdf_to_text]
    elif kind == "zip":                 # xlsx / xlsm / docx
        order = [_excel_to_text, _docx_to_text]
    elif kind == "ole":                 # legacy xls / doc
        order = [_excel_to_text]
    elif text_like:
        order = [_csv_to_text]
    # extension hint (structured formats only; the text reader is gated below)
    hint = by_ext.get(ext)
    if hint and hint not in order and hint is not _csv_to_text:
        order.append(hint)
    # the other structured parsers safely raise on the wrong format
    for reader in (_excel_to_text, _pdf_to_text, _docx_to_text):
        if reader not in order:
            order.append(reader)
    # the raw-text reader last, and only for genuinely text-like bytes
    if text_like and _csv_to_text not in order:
        order.append(_csv_to_text)

    scanned_pdf = False
    for reader in order:
        try:
            text = reader(file_bytes)
            if text and text.strip():
                return text
        except _NoTextPDF:
            scanned_pdf = True
        except Exception:
            continue  # not this format — try the next reader

    if scanned_pdf or kind == "pdf":
        raise RecoError(
            f"“{filename}” looks like a scanned / image-only PDF.",
            "There is no digital text inside it to read. Please re-export the "
            "ledger as a text PDF, Excel, or CSV (a scan can't be read reliably).")

    if text_like:
        raw = _decode(file_bytes)
        if raw.strip():
            return raw

    raise RecoError(
        f"Couldn't read “{filename}”.",
        "The file may be empty, password-protected, or corrupted. "
        "Re-export it as Excel, CSV, or PDF and upload again.")


def _decode(b: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except (UnicodeDecodeError, LookupError):
            pass
    return b.decode("utf-8", errors="replace")


def _best_delimiter(text: str) -> str:
    """Detect the delimiter, falling back to whichever splits rows most
    consistently — so columns are never accidentally glued together."""
    sample = text[:8192]
    try:
        return _csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except _csv.Error:
        best, best_score = ",", -1
        lines = [ln for ln in sample.splitlines() if ln.strip()][:50]
        for d in (",", ";", "\t", "|"):
            counts = [ln.count(d) for ln in lines]
            score = min(counts) if counts else 0   # every row splits the same way
            if score > best_score:
                best, best_score = d, score
        return best


def _csv_to_text(b: bytes) -> str:
    text = _decode(b)
    delim = _best_delimiter(text)
    # csv.reader handles quoted fields and newlines inside quotes correctly,
    # so no row is split or lost. Only fully-blank rows are skipped.
    rows = [r for r in _csv.reader(io.StringIO(text), delimiter=delim)
            if any(c.strip() for c in r)]
    buf = io.StringIO()
    _csv.writer(buf).writerows(rows)
    return buf.getvalue()


def _excel_to_text(b: bytes) -> str:
    last = None
    # openpyxl (xlsx/xlsm), xlrd (legacy xls), then pandas' auto-pick.
    for engine in ("openpyxl", "xlrd", None):
        try:
            xl = pd.ExcelFile(io.BytesIO(b), engine=engine) if engine else pd.ExcelFile(io.BytesIO(b))
            parts = []
            for sheet in xl.sheet_names:                     # every sheet, incl. hidden
                df = xl.parse(sheet, header=None, dtype=str, keep_default_na=False)
                # Trim only rows/cols that are entirely empty (they carry no data).
                df = df.replace("", pd.NA).dropna(how="all").dropna(axis=1, how="all").fillna("")
                if df.empty:
                    continue
                parts.append(f"## Sheet: {sheet}\n{df.to_csv(index=False, header=False)}")
            if parts:
                return "\n\n".join(parts)
            raise RuntimeError("workbook had no data rows")
        except Exception as e:  # try the next engine
            last = e
    raise RuntimeError(f"Could not read the file: {last}")


def _pdf_to_text(b: bytes) -> str:
    import pdfplumber  # lazy: a missing lib just means PDF falls back, no crash

    parts = []
    with pdfplumber.open(io.BytesIO(b)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            # extract_text() returns EVERY text object on the page, so nothing
            # on the page is skipped (unlike table-only extraction, which drops
            # any text the table detector misses). layout=True keeps columns
            # aligned by position so rows stay readable as a table.
            txt = ""
            try:
                txt = page.extract_text(layout=True) or ""
            except Exception:
                txt = ""
            if not txt.strip():
                try:
                    txt = page.extract_text() or ""
                except Exception:
                    txt = ""
            if txt.strip():
                parts.append(f"## Page {i}\n{txt}")
    text = "\n\n".join(parts).strip()
    if not text:
        raise _NoTextPDF()  # opened fine but no readable text → scan/image
    return text


def _docx_to_text(b: bytes) -> str:
    import docx  # python-docx; lazy for the same reason as above
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(io.BytesIO(b))
    out = []

    def emit_table(tbl):
        buf = io.StringIO()
        w = _csv.writer(buf)
        for row in tbl.rows:
            seen, vals = set(), []
            for cell in row.cells:
                tc = id(cell._tc)          # merged cells repeat — count each once
                if tc in seen:
                    continue
                seen.add(tc)
                vals.append(cell.text.replace("\n", " ").strip())
            w.writerow(vals)
        out.append(buf.getvalue())

    # Walk the document body in order so paragraphs and tables stay in sequence.
    for child in d.element.body.iterchildren():
        if child.tag == qn("w:p"):
            t = Paragraph(child, d).text
            if t.strip():
                out.append(t)
        elif child.tag == qn("w:tbl"):
            emit_table(Table(child, d))

    # Headers & footers can hold running totals — include them too.
    for sec in d.sections:
        for hf in (sec.header, sec.footer):
            for p in hf.paragraphs:
                if p.text.strip():
                    out.append(p.text)

    text = "\n".join(out).strip()
    if not text:
        raise RuntimeError("no text in .docx")
    return text


# ---------------------------------------------------------------------------
# Workbook JSON parsing (robust to fences / stray prose)
# ---------------------------------------------------------------------------
def parse_workbook_json(text: str) -> dict:
    txt = text.strip()
    txt = re.sub(r"^```[a-zA-Z]*\n?", "", txt)
    txt = re.sub(r"\n?```$", "", txt.strip())
    try:
        data = json.loads(txt)
    except json.JSONDecodeError:
        data = None
        start, end = txt.find("{"), txt.rfind("}")
        if start != -1 and end > start:
            try:
                data = json.loads(txt[start:end + 1])
            except json.JSONDecodeError:
                pass
    if not isinstance(data, dict) or not isinstance(data.get("sheets"), list):
        return {"sheets": []}
    return data


# ---------------------------------------------------------------------------
# The session: call 1 (streamed report) + call 2 (workbook → Excel)
# ---------------------------------------------------------------------------
class RecoSession:
    def __init__(self, api_key, model_key, a_bytes, a_name, b_bytes, b_name):
        if not api_key:
            raise RuntimeError("ANTHROPIC_API_KEY is not set.")
        self.client = Anthropic(api_key=api_key)
        self.model_key = model_key
        self.model = MODELS[model_key]
        self.a_name, self.b_name = a_name, b_name
        self.a_text = ledger_to_text(a_bytes, a_name)
        self.b_text = ledger_to_text(b_bytes, b_name)
        self.report = ""
        self.workbook = {"sheets": []}
        self.input_tokens = 0
        self.output_tokens = 0

    def _stream(self, prompt, thinking):
        kwargs = dict(
            model=self.model["id"],
            max_tokens=MAX_TOKENS,
            messages=[{"role": "user", "content": prompt}],
        )
        if thinking:
            kwargs["thinking"] = {"type": "adaptive"}
        return self.client.messages.stream(**kwargs)

    def stream_report(self):
        """Yield report text as Claude writes it — feed to st.write_stream."""
        prompt = reconcile_prompt(self.a_name, self.a_text, self.b_name, self.b_text)
        parts = []
        try:
            with self._stream(prompt, thinking=self.model["thinking"]) as s:
                for text in s.text_stream:
                    parts.append(text)
                    yield text
                final = s.get_final_message()
        except anthropic.AnthropicError as e:
            raise RecoError(*describe_api_error(e)) from e
        except Exception as e:
            raise RecoError(
                "The reconciliation couldn't be completed.",
                "Please try again. If it keeps happening, try a different "
                "file or another model.") from e
        self.report = "".join(parts)
        self.input_tokens += final.usage.input_tokens
        self.output_tokens += final.usage.output_tokens

    def build_workbook(self) -> bytes:
        """Second call: reconciliation → {'sheets': [...]} → .xlsx bytes.

        If Claude's Excel layout call fails for any reason, we fall back to
        dumping the (already successful) report into a single sheet, so the
        download is never empty and the user never sees an error here.
        """
        prompt = workbook_prompt(self.a_name, self.a_text, self.b_name, self.b_text, self.report)
        parts = []
        try:
            with self._stream(prompt, thinking=False) as s:
                for text in s.text_stream:
                    parts.append(text)
                final = s.get_final_message()
            self.input_tokens += final.usage.input_tokens
            self.output_tokens += final.usage.output_tokens
            self.workbook = parse_workbook_json("".join(parts))
        except Exception:
            self.workbook = {"sheets": []}  # excel_writer falls back to the report
        return workbook_to_excel(self.workbook, self.report)

    @property
    def cost_usd(self) -> float:
        return (self.input_tokens * self.model["in_per_m"]
                + self.output_tokens * self.model["out_per_m"]) / 1_000_000

    @property
    def cost_inr(self) -> float:
        return self.cost_usd * USD_INR
