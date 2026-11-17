"""PDF invoice extraction: pdfplumber -> raw dict -> validated ``InvoiceRecord``.

Every PDF yields exactly one outcome: an ``InvoiceRecord`` or an
``ExtractionFailure`` with a reason code. Nothing raises past
``extract_invoice`` / ``extract_directory``.

Usage:
    python extractor.py [./sample_invoices]
"""

from __future__ import annotations

import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pdfplumber
from pydantic import ValidationError

from schemas import ExtractionFailure, InvoiceRecord, ReasonCode

log = logging.getLogger(__name__)

MATH_ERROR_TYPES = {"invoice_math"}
DATE_FORMATS = ("%Y-%m-%d", "%b %d, %Y", "%B %d, %Y", "%m/%d/%Y", "%d %b %Y", "%d %B %Y")
MONEY_RE = r"\$?\s*(-?[\d,]+\.\d{2})"

FIELD_PATTERNS: dict[str, re.Pattern[str]] = {
    "invoice_id": re.compile(r"Invoice\s*(?:Number|No\.?|#)\s*:?\s*([A-Za-z0-9][A-Za-z0-9\-/]+)", re.I),
    "invoice_date": re.compile(r"Invoice\s*Date\s*:?\s*(.+)", re.I),
    "due_date": re.compile(r"Due\s*Date\s*:?\s*(.+)", re.I),
    "currency": re.compile(r"Currency\s*:?\s*([A-Z]{3})", re.I),
    "subtotal": re.compile(r"^\s*Sub[\s-]?total\s*:?\s*" + MONEY_RE, re.I | re.M),
    "tax": re.compile(r"^\s*(?:Sales\s+)?Tax\b[^:\n]*:?\s*" + MONEY_RE, re.I | re.M),
    "total": re.compile(r"^\s*(?:Total\s*Due|Amount\s*Due|Total)\s*:?\s*" + MONEY_RE, re.I | re.M),
}
# Fallback for line items when no ruled table is detected.
LINE_ITEM_RE = re.compile(r"^(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+\$?(?P<price>[\d,]+\.\d{2})\s+\$?(?P<amount>[\d,]+\.\d{2})\s*$")


class ExtractionError(Exception):
    """Raised internally when a PDF cannot be read or parsed."""


def parse_money(raw: str | None) -> str | None:
    if raw is None:
        return None
    cleaned = raw.replace("$", "").replace(",", "").strip()
    return cleaned or None


def parse_date(raw: str | None) -> date | str | None:
    """Return a ``date`` when a known format matches, else the raw string so
    pydantic reports a clear schema error instead of us guessing."""
    if raw is None:
        return None
    raw = raw.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return raw


def _first_line(text: str) -> str | None:
    for line in text.splitlines():
        if line.strip():
            return line.strip()
    return None


def _items_from_tables(tables: list[list[list[str | None]]]) -> list[dict[str, Any]]:
    for table in tables:
        if not table:
            continue
        header = [(c or "").strip().lower() for c in table[0]]
        if "description" not in header or "amount" not in header:
            continue
        idx = {name: header.index(name) for name in header if name}
        qty_i = idx.get("qty", idx.get("quantity"))
        price_i = idx.get("unit price", idx.get("rate", idx.get("price")))
        items = []
        for row in table[1:]:
            cells = [(c or "").strip() for c in row]
            if not any(cells):
                continue
            items.append({
                "description": cells[idx["description"]],
                "quantity": cells[qty_i] if qty_i is not None else "1",
                "unit_price": parse_money(cells[price_i]) if price_i is not None else parse_money(cells[idx["amount"]]),
                "amount": parse_money(cells[idx["amount"]]),
            })
        if items:
            return items
    return []


def _items_from_text(text: str) -> list[dict[str, Any]]:
    items, in_table = [], False
    for line in text.splitlines():
        if re.match(r"^\s*Description\b", line, re.I):
            in_table = True
            continue
        if in_table and FIELD_PATTERNS["subtotal"].match(line):
            break
        if in_table and (m := LINE_ITEM_RE.match(line.strip())):
            items.append({
                "description": m["desc"],
                "quantity": m["qty"],
                "unit_price": parse_money(m["price"]),
                "amount": parse_money(m["amount"]),
            })
    return items


def extract_raw(pdf_path: Path) -> dict[str, Any]:
    """Parse a PDF into an unvalidated dictionary of invoice fields."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            tables = [t for page in pdf.pages for t in page.extract_tables()]
    except Exception as exc:  # pdfminer raises a zoo of exception types
        raise ExtractionError(f"Unreadable PDF: {exc}") from exc
    if not text.strip():
        raise ExtractionError("PDF contains no extractable text (scanned image?)")

    raw: dict[str, Any] = {"source_file": pdf_path.name, "vendor_name": _first_line(text)}
    for key, pattern in FIELD_PATTERNS.items():
        m = pattern.search(text)
        raw[key] = m.group(1).strip() if m else None

    raw["invoice_date"] = parse_date(raw["invoice_date"])
    raw["due_date"] = parse_date(raw["due_date"])
    for key in ("subtotal", "tax", "total"):
        raw[key] = parse_money(raw[key])
    raw["currency"] = (raw["currency"] or "USD").upper()
    raw["line_items"] = _items_from_tables(tables) or _items_from_text(text)
    return {k: v for k, v in raw.items() if v is not None}


def _format_errors(exc: ValidationError) -> tuple[ReasonCode, list[str]]:
    errors = exc.errors(include_url=False)
    is_math = any(e["type"] in MATH_ERROR_TYPES for e in errors)
    messages = []
    for e in errors:
        loc = ".".join(str(p) for p in e["loc"]) or "invoice"
        messages.append(f"{loc}: {e['msg']}")
    return (ReasonCode.MATH_ERROR if is_math else ReasonCode.SCHEMA_ERROR), messages


def extract_invoice(pdf_path: Path) -> InvoiceRecord | ExtractionFailure:
    pdf_path = Path(pdf_path)
    try:
        raw = extract_raw(pdf_path)
    except ExtractionError as exc:
        log.warning("%s: %s", pdf_path.name, exc)
        return ExtractionFailure(source_file=pdf_path.name, reason=ReasonCode.PARSE_ERROR, errors=[str(exc)])

    try:
        return InvoiceRecord.model_validate(raw)
    except ValidationError as exc:
        reason, messages = _format_errors(exc)
        log.warning("%s rejected (%s): %s", pdf_path.name, reason.value, "; ".join(messages))
        return ExtractionFailure(
            source_file=pdf_path.name,
            reason=reason,
            invoice_id=raw.get("invoice_id"),
            vendor_name=raw.get("vendor_name"),
            total=raw.get("total"),
            errors=messages,
        )


@dataclass
class ExtractionBatch:
    records: list[InvoiceRecord] = field(default_factory=list)
    failures: list[ExtractionFailure] = field(default_factory=list)


def flag_duplicates(records: list[InvoiceRecord]) -> tuple[list[InvoiceRecord], list[ExtractionFailure]]:
    """Keep the first copy of each invoice ID (by invoice date, then file name);
    later copies become ``DUPLICATE_INVOICE_ID`` failures so they can never be paid twice."""
    kept: dict[str, InvoiceRecord] = {}
    dupes: list[ExtractionFailure] = []
    for rec in sorted(records, key=lambda r: (r.invoice_date, r.source_file or "")):
        original = kept.get(rec.invoice_id)
        if original is None:
            kept[rec.invoice_id] = rec
            continue
        same = (original.vendor_name, original.total) == (rec.vendor_name, rec.total)
        dupes.append(ExtractionFailure(
            source_file=rec.source_file or "?",
            reason=ReasonCode.DUPLICATE_INVOICE_ID,
            invoice_id=rec.invoice_id,
            vendor_name=rec.vendor_name,
            total=rec.total,
            errors=[
                f"Invoice ID already seen in {original.source_file}"
                + ("" if same else f" with DIFFERENT content (vendor/total {original.vendor_name}/{original.total})")
            ],
        ))
    return list(kept.values()), dupes


def extract_directory(directory: Path) -> ExtractionBatch:
    batch = ExtractionBatch()
    valid: list[InvoiceRecord] = []
    for pdf_path in sorted(Path(directory).glob("*.pdf")):
        outcome = extract_invoice(pdf_path)
        (valid if isinstance(outcome, InvoiceRecord) else batch.failures).append(outcome)
    batch.records, dupes = flag_duplicates(valid)
    batch.failures.extend(dupes)
    return batch


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    directory = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample_invoices")
    batch = extract_directory(directory)
    print(f"\n{len(batch.records)} valid invoices, {len(batch.failures)} rejected")
    for r in batch.records:
        print(f"  OK   {r.invoice_id}  {r.vendor_name:<36} {r.total:>10}")
    for f in batch.failures:
        print(f"  FAIL {f.source_file:<34} {f.reason.value}: {'; '.join(f.errors)}")


if __name__ == "__main__":
    main()
