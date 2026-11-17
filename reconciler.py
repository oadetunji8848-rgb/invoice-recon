"""Match validated invoices to bank ledger entries and route the results.

Scoring (0-100), computed only for candidates that pass the hard gates
(amount equal to the cent AND posted within +/-5 days of the invoice date):

    amount match      50 pts  (always awarded; it is a hard gate)
    date proximity    20 pts  (full within 2 days, then down to 10 at day 5)
    vendor similarity 30 pts  (rapidfuzz token_sort_ratio on normalized names)

Routing:
    approved      score >= 90 AND vendor similarity >= 80 AND the match is unambiguous
    needs_review  everything else, including no match, math errors, duplicates

Usage:
    python reconciler.py [--invoices ./sample_invoices] [--ledger ./bank_ledger.csv] [--out ./output]
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import ValidationError
from rapidfuzz import fuzz

from extractor import extract_directory
from schemas import (
    ExtractionFailure,
    InvoiceRecord,
    LedgerEntry,
    ReasonCode,
    ReconciliationResult,
)

log = logging.getLogger(__name__)

# ----------------------------------------------------------------- thresholds
DATE_WINDOW_DAYS = 5
DATE_FULL_CREDIT_DAYS = 2
AMOUNT_POINTS = 50.0
DATE_POINTS = 20.0
DATE_POINTS_AT_EDGE = 10.0
VENDOR_POINTS = 30.0

APPROVE_THRESHOLD = 90.0
MIN_VENDOR_SIM_FOR_APPROVAL = 80.0
FUZZY_FLAG_BELOW = 95.0
AMBIGUITY_MARGIN = 5.0  # runner-up within this many points => ambiguous

# Tokens that carry no vendor identity: legal suffixes and bank memo noise.
NOISE_TOKENS = {
    "inc", "incorporated", "llc", "llp", "ltd", "limited", "co", "corp", "corporation",
    "company", "plc", "the", "and",
    "ach", "pmt", "payment", "wire", "transfer", "check", "chk", "pos", "debit", "eft", "ref",
}


def normalize_vendor(name: str) -> str:
    tokens = re.sub(r"[^a-z0-9 ]+", " ", name.lower().replace("&", " and ")).split()
    return " ".join(t for t in tokens if t not in NOISE_TOKENS and not t.isdigit())


def vendor_similarity(invoice_vendor: str, ledger_description: str) -> float:
    a, b = normalize_vendor(invoice_vendor), normalize_vendor(ledger_description)
    if not a or not b:
        return 0.0
    return round(float(fuzz.token_sort_ratio(a, b)), 1)


def date_score(delta_days: int) -> float:
    d = abs(delta_days)
    if d > DATE_WINDOW_DAYS:
        return 0.0
    if d <= DATE_FULL_CREDIT_DAYS:
        return DATE_POINTS
    span = DATE_WINDOW_DAYS - DATE_FULL_CREDIT_DAYS
    return DATE_POINTS - (d - DATE_FULL_CREDIT_DAYS) * (DATE_POINTS - DATE_POINTS_AT_EDGE) / span


@dataclass(frozen=True)
class Candidate:
    invoice: InvoiceRecord
    entry: LedgerEntry
    score: float
    similarity: float
    delta_days: int


def score_pair(invoice: InvoiceRecord, entry: LedgerEntry) -> Candidate | None:
    """Return a scored candidate, or None if the pair fails a hard gate."""
    if invoice.total != entry.amount:
        return None
    delta = (entry.posted_date - invoice.invoice_date).days
    if abs(delta) > DATE_WINDOW_DAYS:
        return None
    sim = vendor_similarity(invoice.vendor_name, entry.description)
    score = AMOUNT_POINTS + date_score(delta) + VENDOR_POINTS * sim / 100.0
    return Candidate(invoice, entry, round(min(score, 100.0), 1), sim, delta)


# --------------------------------------------------------------------- ledger
@dataclass
class LedgerLoad:
    entries: list[LedgerEntry] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "txn_id": ("txn_id", "transaction id", "transaction_id", "txn id", "reference", "ref", "reference number",
               "ref no", "id", "fitid"),
    "posted_date": ("posted_date", "posting date", "posted date", "date", "transaction date", "trans date",
                    "value date", "booking date"),
    "description": ("description", "payee", "narration", "narrative", "memo", "details", "name",
                    "transaction details", "remarks"),
    "amount": ("amount", "debit", "debit amount", "withdrawal", "withdrawals", "money out", "paid out"),
}
CREDIT_ALIASES = ("credit", "credit amount", "deposit", "deposits", "money in", "paid in")
LEDGER_DATE_FORMATS = {
    "iso": ("%Y-%m-%d", "%Y/%m/%d"),
    "day-first": ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%b-%Y", "%d %b %Y", "%d-%b-%y"),
    "month-first": ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%b %d, %Y", "%b %d %Y"),
}


def _canon(header: str) -> str:
    return re.sub(r"[\s_]+", " ", header.strip().lower())


def _map_columns(fieldnames: list[str]) -> tuple[dict[str, str], str | None]:
    by_canon = {_canon(f): f for f in fieldnames}
    mapping = {}
    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if _canon(alias) in by_canon:
                mapping[key] = by_canon[_canon(alias)]
                break
    credit = next((by_canon[_canon(a)] for a in CREDIT_ALIASES if _canon(a) in by_canon), None)
    missing = {"posted_date", "description", "amount"} - mapping.keys()
    if missing:
        raise ValueError(
            f"Ledger is missing column(s) {sorted(missing)}. Found columns: {fieldnames}. "
            f"Accepted names: { {k: COLUMN_ALIASES[k] for k in sorted(missing)} }"
        )
    return mapping, credit


def _pick_date_parser(values: list[str], date_order: str) -> tuple[str, ...]:
    """Choose one consistent format family for the whole column so that
    03/04/2026 is never read as March in one row and April in another."""
    if date_order != "auto":
        return LEDGER_DATE_FORMATS["iso"] + LEDGER_DATE_FORMATS[date_order]

    def parses_all(fmts: tuple[str, ...]) -> bool:
        return all(_parse_with(v, fmts) is not None for v in values if v.strip())

    for family in ("iso", "day-first", "month-first"):
        fmts = LEDGER_DATE_FORMATS["iso"] + LEDGER_DATE_FORMATS[family]
        if parses_all(fmts):
            return fmts
    return LEDGER_DATE_FORMATS["iso"] + LEDGER_DATE_FORMATS["day-first"]


def _parse_with(value: str, fmts: tuple[str, ...]):
    value = value.strip()
    for fmt in fmts:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _parse_amount(raw: str | None) -> Decimal | None:
    if raw is None or not raw.strip():
        return None
    text = raw.strip()
    negative = text.startswith("(") and text.endswith(")")
    cleaned = re.sub(r"[^\d.\-]", "", text)
    if cleaned in ("", "-", "."):
        raise InvalidOperation(raw)
    value = Decimal(cleaned)
    return -value if negative else value


@dataclass
class LedgerLoad:
    entries: list[LedgerEntry] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    skipped_credits: int = 0


def load_ledger(path: Path, date_order: str = "auto") -> LedgerLoad:
    """Load a bank export. Column names are matched against common aliases;
    rows that are deposits (credit column only) are skipped; outgoing
    payments are stored as positive amounts."""
    if date_order not in ("auto", "day-first", "month-first"):
        raise ValueError("date_order must be auto, day-first or month-first")
    out = LedgerLoad()
    with Path(path).open(newline="", encoding="utf-8-sig") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.DictReader(fh, dialect=dialect))
        fieldnames = list(rows[0].keys()) if rows else []
    if not rows:
        raise ValueError(f"Ledger {path} has no data rows")
    cols, credit_col = _map_columns(fieldnames)
    date_fmts = _pick_date_parser([r.get(cols["posted_date"]) or "" for r in rows], date_order)

    seen: set[str] = set()
    for lineno, row in enumerate(rows, start=2):
        try:
            amount = _parse_amount(row.get(cols["amount"]))
            if amount is None:
                if credit_col and _parse_amount(row.get(credit_col)) is not None:
                    out.skipped_credits += 1
                    continue
                raise InvalidOperation("empty amount")
            posted = _parse_with(row.get(cols["posted_date"]) or "", date_fmts)
            if posted is None:
                raise ValueError(f"unrecognised date {row.get(cols['posted_date'])!r}")
            txn_id = (row.get(cols["txn_id"]) or "").strip() if "txn_id" in cols else ""
            entry = LedgerEntry(
                txn_id=txn_id or f"ROW-{lineno}",
                posted_date=posted,
                description=row.get(cols["description"]) or "",
                amount=abs(amount),
            )
        except (InvalidOperation, ValidationError, ValueError, AttributeError) as exc:
            out.errors.append(f"line {lineno}: {exc}".splitlines()[0])
            continue
        if entry.txn_id in seen:
            out.errors.append(f"line {lineno}: duplicate txn_id {entry.txn_id} skipped")
            continue
        seen.add(entry.txn_id)
        out.entries.append(entry)
    for err in out.errors:
        log.warning("ledger %s", err)
    return out


# ------------------------------------------------------------------ reconcile
@dataclass
class ReconciliationReport:
    approved: list[ReconciliationResult]
    review: list[ReconciliationResult]
    unmatched_ledger: list[LedgerEntry]
    ledger_errors: list[str] = field(default_factory=list)
    generated_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat(timespec="seconds"))

    @property
    def all_results(self) -> list[ReconciliationResult]:
        return self.approved + self.review

    def by_invoice_id(self, invoice_id: str) -> list[ReconciliationResult]:
        return [r for r in self.all_results if r.invoice_id == invoice_id]


def _result_from_failure(f: ExtractionFailure) -> ReconciliationResult:
    return ReconciliationResult(
        status="needs_review", confidence=0.0, reasons=[f.reason], detail="; ".join(f.errors),
        source_file=f.source_file, invoice_id=f.invoice_id, vendor_name=f.vendor_name, invoice_total=f.total,
    )


def _result_from_match(c: Candidate, ambiguous: bool) -> ReconciliationResult:
    reasons: list[ReasonCode] = []
    if c.similarity < MIN_VENDOR_SIM_FOR_APPROVAL:
        reasons.append(ReasonCode.LOW_VENDOR_SIMILARITY)
    elif c.similarity < FUZZY_FLAG_BELOW:
        reasons.append(ReasonCode.FUZZY_VENDOR)
    if abs(c.delta_days) > DATE_FULL_CREDIT_DAYS:
        reasons.append(ReasonCode.DATE_DRIFT)
    if ambiguous:
        reasons.append(ReasonCode.MULTIPLE_CANDIDATES)
    if c.score < APPROVE_THRESHOLD:
        reasons.append(ReasonCode.LOW_CONFIDENCE)

    approved = (
        c.score >= APPROVE_THRESHOLD
        and c.similarity >= MIN_VENDOR_SIM_FOR_APPROVAL
        and not ambiguous
    )
    if not reasons:
        reasons.append(ReasonCode.EXACT_MATCH)
    inv, e = c.invoice, c.entry
    return ReconciliationResult(
        status="approved" if approved else "needs_review",
        confidence=c.score,
        reasons=reasons,
        detail=f"amount exact; posted {c.delta_days:+d}d; vendor similarity {c.similarity:.1f}",
        source_file=inv.source_file or "?", invoice_id=inv.invoice_id, vendor_name=inv.vendor_name,
        invoice_date=inv.invoice_date, invoice_total=inv.total,
        ledger_txn_id=e.txn_id, ledger_description=e.description, ledger_date=e.posted_date, ledger_amount=e.amount,
        vendor_similarity=c.similarity, date_delta_days=c.delta_days,
    )


def reconcile(
    invoices: list[InvoiceRecord],
    ledger: list[LedgerEntry],
    failures: list[ExtractionFailure] | None = None,
) -> ReconciliationReport:
    """One-to-one assignment: best-scoring pairs are claimed first, so a ledger
    entry can settle at most one invoice and vice versa."""
    candidates_by_invoice: dict[str, list[Candidate]] = {}
    all_candidates: list[Candidate] = []
    for inv in invoices:
        cands = [c for e in ledger if (c := score_pair(inv, e))]
        candidates_by_invoice[inv.invoice_id] = sorted(cands, key=lambda c: -c.score)
        all_candidates.extend(cands)

    # Deterministic tie-break: score, then smaller date drift, then txn id.
    all_candidates.sort(key=lambda c: (-c.score, abs(c.delta_days), c.entry.txn_id, c.invoice.invoice_id))
    claimed_inv: set[str] = set()
    claimed_txn: set[str] = set()
    results: list[ReconciliationResult] = []
    for c in all_candidates:
        if c.invoice.invoice_id in claimed_inv or c.entry.txn_id in claimed_txn:
            continue
        claimed_inv.add(c.invoice.invoice_id)
        claimed_txn.add(c.entry.txn_id)
        rivals = [o for o in candidates_by_invoice[c.invoice.invoice_id] if o.entry.txn_id != c.entry.txn_id]
        ambiguous = any(c.score - o.score <= AMBIGUITY_MARGIN for o in rivals)
        results.append(_result_from_match(c, ambiguous))

    for inv in invoices:
        if inv.invoice_id in claimed_inv:
            continue
        had = candidates_by_invoice[inv.invoice_id]
        results.append(ReconciliationResult(
            status="needs_review", confidence=0.0, reasons=[ReasonCode.NO_LEDGER_MATCH],
            detail=(
                "candidate ledger entries were claimed by better-scoring invoices"
                if had else f"no ledger entry for {inv.total} within +/-{DATE_WINDOW_DAYS} days"
            ),
            source_file=inv.source_file or "?", invoice_id=inv.invoice_id, vendor_name=inv.vendor_name,
            invoice_date=inv.invoice_date, invoice_total=inv.total,
        ))

    results.extend(_result_from_failure(f) for f in failures or [])
    results.sort(key=lambda r: (r.invoice_id or "~", r.source_file))
    return ReconciliationReport(
        approved=[r for r in results if r.status == "approved"],
        review=sorted((r for r in results if r.status == "needs_review"), key=lambda r: (r.confidence, r.source_file)),
        unmatched_ledger=[e for e in ledger if e.txn_id not in claimed_txn],
    )


# -------------------------------------------------------------------- outputs
def _total(results: list[ReconciliationResult]) -> Decimal:
    return sum((r.invoice_total or Decimal("0") for r in results), Decimal("0"))


def write_json(results: list[ReconciliationResult], path: Path, generated_at: str) -> None:
    payload = {
        "generated_at": generated_at,
        "count": len(results),
        "total_amount": str(_total(results)),
        "items": [r.model_dump(mode="json") for r in results],
    }
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color="1F3A5F")
THIN = Side(style="thin", color="D0D7E2")
BORDER = Border(bottom=THIN)
MONEY_FMT = '"$"#,##0.00'
GREEN_FILL = PatternFill("solid", fgColor="E3F4E8")
AMBER_FILL = PatternFill("solid", fgColor="FFF3D6")
RED_FILL = PatternFill("solid", fgColor="FBE3E1")


def _write_table(ws, headers: list[str], rows: list[list], money_cols: set[int], widths: list[int]) -> None:
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append(row)
    for r in ws.iter_rows(min_row=2):
        for cell in r:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.column in money_cols:
                cell.number_format = MONEY_FMT
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def _result_rows(results: list[ReconciliationResult]) -> list[list]:
    return [[
        r.invoice_id or "", r.vendor_name or "", r.invoice_date, float(r.invoice_total) if r.invoice_total else None,
        r.ledger_txn_id or "", r.ledger_description or "", r.ledger_date,
        r.confidence, r.vendor_similarity, r.date_delta_days,
        ", ".join(x.value for x in r.reasons), r.detail, r.source_file,
    ] for r in results]


RESULT_HEADERS = [
    "Invoice ID", "Vendor", "Invoice Date", "Invoice Total", "Ledger Txn", "Ledger Description", "Posted Date",
    "Confidence", "Vendor Sim.", "Date Delta (d)", "Reasons", "Detail", "Source File",
]
RESULT_WIDTHS = [16, 32, 13, 14, 12, 40, 13, 12, 11, 12, 34, 60, 36]


def write_excel(report: ReconciliationReport, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Invoice-to-Ledger Reconciliation"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {report.generated_at} (UTC)"
    ws["A2"].font = Font(italic=True, color="666666")

    n_total = len(report.all_results)
    approved_amt, review_amt = _total(report.approved), _total(report.review)
    by_reason: dict[str, int] = {}
    for r in report.review:
        for reason in r.reasons:
            by_reason[reason.value] = by_reason.get(reason.value, 0) + 1

    kpis = [
        ("Documents processed", n_total, None),
        ("Auto-approved", len(report.approved), None),
        ("Needs review", len(report.review), None),
        ("Auto-approval rate", (len(report.approved) / n_total) if n_total else 0, "0.0%"),
        ("Approved value", float(approved_amt), MONEY_FMT),
        ("Value held for review", float(review_amt), MONEY_FMT),
        ("Unmatched ledger entries", len(report.unmatched_ledger), None),
        ("Unmatched ledger value", float(sum((e.amount for e in report.unmatched_ledger), Decimal("0"))), MONEY_FMT),
        ("Ledger rows rejected on load", len(report.ledger_errors), None),
    ]
    ws["A4"], ws["B4"] = "Key metric", "Value"
    for c in (ws["A4"], ws["B4"]):
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for i, (label, value, fmt) in enumerate(kpis, start=5):
        ws.cell(row=i, column=1, value=label).border = BORDER
        cell = ws.cell(row=i, column=2, value=value)
        cell.border = BORDER
        cell.font = Font(bold=True)
        if fmt:
            cell.number_format = fmt

    row = 5 + len(kpis) + 1
    ws.cell(row=row, column=1, value="Review queue by reason").font = Font(bold=True, color="1F3A5F")
    row += 1
    for reason, count in sorted(by_reason.items(), key=lambda kv: -kv[1]):
        ws.cell(row=row, column=1, value=reason).border = BORDER
        ws.cell(row=row, column=2, value=count).border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Scoring policy").font = Font(bold=True, color="1F3A5F")
    for text in (
        f"Hard gates: amount equal to the cent; posted within +/-{DATE_WINDOW_DAYS} days",
        f"Score = {AMOUNT_POINTS:.0f} amount + up to {DATE_POINTS:.0f} date + up to {VENDOR_POINTS:.0f} vendor similarity",
        f"Approve when score >= {APPROVE_THRESHOLD:.0f}, vendor similarity >= {MIN_VENDOR_SIM_FOR_APPROVAL:.0f}, "
        "and no rival candidate within "
        f"{AMBIGUITY_MARGIN:.0f} pts",
        "Math errors, duplicates, parse failures and unmatched invoices always go to review",
    ):
        row += 1
        ws.cell(row=row, column=1, value=text)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18

    for title, results in (("Approved", report.approved), ("Review Queue", report.review)):
        sheet = wb.create_sheet(title)
        _write_table(sheet, RESULT_HEADERS, _result_rows(results), {4}, RESULT_WIDTHS)
        for r in range(2, sheet.max_row + 1):
            for c in (3, 7):
                sheet.cell(row=r, column=c).number_format = "yyyy-mm-dd"
            sheet.cell(row=r, column=8).number_format = "0.0"
        if results:
            rng = f"H2:H{sheet.max_row}"
            sheet.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=[str(APPROVE_THRESHOLD)], fill=GREEN_FILL))
            sheet.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1", str(APPROVE_THRESHOLD - 0.01)], fill=AMBER_FILL))
            sheet.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1"], fill=RED_FILL))

    ul = wb.create_sheet("Unmatched Ledger")
    _write_table(
        ul, ["Txn ID", "Posted Date", "Description", "Amount"],
        [[e.txn_id, e.posted_date, e.description, float(e.amount)] for e in report.unmatched_ledger],
        {4}, [14, 13, 48, 14],
    )
    for r in range(2, ul.max_row + 1):
        ul.cell(row=r, column=2).number_format = "yyyy-mm-dd"

    wb.save(path)


def write_outputs(report: ReconciliationReport, out_dir: Path) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "approved": out_dir / "approved_matches.json",
        "review": out_dir / "review_queue.json",
        "report": out_dir / "reconciliation_report.xlsx",
    }
    write_json(report.approved, paths["approved"], report.generated_at)
    write_json(report.review, paths["review"], report.generated_at)
    write_excel(report, paths["report"])
    return paths


def _md_money(v: Decimal | None) -> str:
    return f"{v:,.2f}" if v is not None else ""


def _md_cell(text: object) -> str:
    return str(text if text is not None else "").replace("|", "\\|").replace("\n", " ")


def write_markdown_summary(report: ReconciliationReport, path: Path) -> None:
    """A readable summary for GitHub's job-summary page (or any Markdown viewer)."""
    lines = [
        "## Reconciliation results",
        "",
        f"**{len(report.approved)}** auto-approved ({_md_money(_total(report.approved))}) · "
        f"**{len(report.review)}** need review ({_md_money(_total(report.review))}) · "
        f"**{len(report.unmatched_ledger)}** ledger payments with no invoice",
        "",
    ]
    if report.ledger_errors:
        lines += [f"> ⚠️ {len(report.ledger_errors)} ledger row(s) could not be read: "
                  + "; ".join(report.ledger_errors[:5]) + ("…" if len(report.ledger_errors) > 5 else ""), ""]
    if report.review:
        lines += ["### Needs review", "", "| Invoice | Vendor | Total | Confidence | Why | Detail |",
                  "|---|---|--:|--:|---|---|"]
        for r in report.review:
            lines.append("| " + " | ".join(_md_cell(x) for x in (
                r.invoice_id or r.source_file, r.vendor_name, _md_money(r.invoice_total), f"{r.confidence:.0f}",
                ", ".join(x.value for x in r.reasons), r.detail)) + " |")
        lines.append("")
    if report.approved:
        lines += ["### Approved", "", "| Invoice | Vendor | Total | Ledger txn | Posted | Confidence |",
                  "|---|---|--:|---|---|--:|"]
        for r in report.approved:
            lines.append("| " + " | ".join(_md_cell(x) for x in (
                r.invoice_id, r.vendor_name, _md_money(r.invoice_total), r.ledger_txn_id, r.ledger_date,
                f"{r.confidence:.0f}")) + " |")
        lines.append("")
    if report.unmatched_ledger:
        lines += ["### Ledger payments with no matching invoice", "", "| Txn | Posted | Description | Amount |",
                  "|---|---|---|--:|"]
        for e in report.unmatched_ledger:
            lines.append(f"| {_md_cell(e.txn_id)} | {e.posted_date} | {_md_cell(e.description)} | {_md_money(e.amount)} |")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


class PipelineInputError(Exception):
    """Bad input paths or files; reported to the user without a traceback."""


def run_pipeline(
    invoice_dir: Path,
    ledger_path: Path,
    out_dir: Path | None = None,
    date_order: str = "auto",
) -> ReconciliationReport:
    invoice_dir, ledger_path = Path(invoice_dir), Path(ledger_path)
    if not invoice_dir.is_dir():
        raise PipelineInputError(f"Invoice folder not found: {invoice_dir}")
    if not any(invoice_dir.glob("*.pdf")) and not any(invoice_dir.glob("*.PDF")):
        raise PipelineInputError(f"No PDF files in {invoice_dir}")
    if not ledger_path.is_file():
        raise PipelineInputError(f"Ledger CSV not found: {ledger_path}")
    batch = extract_directory(invoice_dir)
    try:
        ledger = load_ledger(ledger_path, date_order=date_order)
    except (ValueError, UnicodeDecodeError) as exc:
        raise PipelineInputError(f"Could not read ledger {ledger_path}: {exc}") from exc
    report = reconcile(batch.records, ledger.entries, batch.failures)
    report.ledger_errors = ledger.errors
    if out_dir is not None:
        write_outputs(report, Path(out_dir))
        write_markdown_summary(report, Path(out_dir) / "summary.md")
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--invoices", type=Path, default=Path("sample_invoices"), help="folder of PDF invoices")
    ap.add_argument("--ledger", type=Path, default=Path("bank_ledger.csv"), help="bank export CSV")
    ap.add_argument("--out", type=Path, default=Path("output"))
    ap.add_argument("--date-order", choices=["auto", "day-first", "month-first"], default="auto",
                    help="how to read ambiguous ledger dates like 03/04/2026 (auto prefers day-first)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO if args.verbose else logging.WARNING, format="%(levelname)s %(message)s")

    try:
        report = run_pipeline(args.invoices, args.ledger, args.out, args.date_order)
    except PipelineInputError as exc:
        print(f"ERROR: {exc}")
        return 2
    print(f"Approved: {len(report.approved)}  Needs review: {len(report.review)}  "
          f"Unmatched ledger: {len(report.unmatched_ledger)}")
    for r in report.approved:
        print(f"  APPROVED {r.invoice_id}  {r.confidence:5.1f}  {r.vendor_name}")
    for r in report.review:
        print(f"  REVIEW   {r.invoice_id or '-':<14} {r.confidence:5.1f}  {','.join(x.value for x in r.reasons)}")
    print(f"Outputs written to {args.out}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
