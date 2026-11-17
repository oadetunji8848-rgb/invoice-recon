"""pytest suite for the invoice-to-ledger reconciliation pipeline."""

from __future__ import annotations

import json
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook
from pydantic import ValidationError

import seed_invoices
from extractor import extract_directory, extract_invoice, flag_duplicates
from reconciler import (
    APPROVE_THRESHOLD,
    load_ledger,
    normalize_vendor,
    reconcile,
    run_pipeline,
    score_pair,
    vendor_similarity,
)
from schemas import (
    InvoiceRecord,
    LedgerEntry,
    LineItem,
    ReasonCode,
    ReconciliationResult,
)

D = Decimal
BASE_DATE = date(2026, 10, 10)


# ------------------------------------------------------------------ fixtures
@pytest.fixture(scope="session")
def seeded(tmp_path_factory) -> tuple[Path, Path]:
    root = tmp_path_factory.mktemp("seed")
    inv_dir, ledger = root / "sample_invoices", root / "bank_ledger.csv"
    seed_invoices.generate(inv_dir, ledger)
    return inv_dir, ledger


@pytest.fixture(scope="session")
def pipeline(seeded, tmp_path_factory):
    inv_dir, ledger = seeded
    out = tmp_path_factory.mktemp("out")
    return run_pipeline(inv_dir, ledger, out), out


def make_invoice(
    invoice_id: str = "INV-T-1",
    vendor: str = "Northwind Logistics LLC",
    total: str = "1080.00",
    on: date = BASE_DATE,
    source: str = "t.pdf",
) -> InvoiceRecord:
    amount = D(total)
    return InvoiceRecord(
        invoice_id=invoice_id,
        vendor_name=vendor,
        invoice_date=on,
        line_items=[LineItem(description="Service", quantity=D(1), unit_price=amount, amount=amount)],
        subtotal=amount,
        tax=D("0"),
        total=amount,
        source_file=source,
    )


def make_entry(txn: str = "TXN-1", desc: str = "ACH PMT NORTHWIND LOGISTICS", amount: str = "1080.00",
               days: int = 0) -> LedgerEntry:
    return LedgerEntry(txn_id=txn, posted_date=BASE_DATE + timedelta(days=days), description=desc, amount=D(amount))


def raw_invoice(**overrides) -> dict:
    data = {
        "invoice_id": "INV-1",
        "vendor_name": "Acme",
        "invoice_date": "2026-10-01",
        "line_items": [
            {"description": "A", "quantity": "2", "unit_price": "10.00", "amount": "20.00"},
            {"description": "B", "quantity": "1", "unit_price": "5.50", "amount": "5.50"},
        ],
        "subtotal": "25.50",
        "tax": "2.10",
        "total": "27.60",
    }
    data.update(overrides)
    return data


def _error_types(exc: pytest.ExceptionInfo[ValidationError]) -> set[str]:
    return {e["type"] for e in exc.value.errors()}


# ---------------------------------------------------------- math validation
class TestMathValidation:
    def test_valid_invoice_passes(self):
        rec = InvoiceRecord.model_validate(raw_invoice())
        assert rec.total == D("27.60")

    def test_total_mismatch_rejected(self):
        with pytest.raises(ValidationError) as exc:
            InvoiceRecord.model_validate(raw_invoice(total="42.60"))
        assert "invoice_math" in _error_types(exc)

    def test_line_items_not_summing_to_subtotal_rejected(self):
        with pytest.raises(ValidationError) as exc:
            InvoiceRecord.model_validate(raw_invoice(subtotal="30.00", total="32.10"))
        assert "invoice_math" in _error_types(exc)

    def test_line_item_extension_error_rejected(self):
        with pytest.raises(ValidationError) as exc:
            LineItem(description="X", quantity=D(3), unit_price=D("10.00"), amount=D("29.00"))
        assert "invoice_math" in _error_types(exc)

    @pytest.mark.parametrize("total,ok", [("27.61", True), ("27.59", True), ("27.62", False), ("27.58", False)])
    def test_one_cent_tolerance(self, total, ok):
        if ok:
            InvoiceRecord.model_validate(raw_invoice(total=total))
        else:
            with pytest.raises(ValidationError):
                InvoiceRecord.model_validate(raw_invoice(total=total))

    def test_seeded_math_error_pdf_is_rejected_cleanly(self, seeded):
        inv_dir, _ = seeded
        bad_id = seed_invoices.EDGE_CASES["math_error"]
        failure = extract_invoice(next(inv_dir.glob(f"*{bad_id}.pdf")))
        assert not isinstance(failure, InvoiceRecord)
        assert failure.reason is ReasonCode.MATH_ERROR
        assert failure.invoice_id == bad_id
        assert "1787.35" in failure.errors[0]

    def test_unreadable_pdf_becomes_parse_error(self, tmp_path):
        bogus = tmp_path / "corrupt.pdf"
        bogus.write_bytes(b"%PDF-1.4 not really a pdf")
        result = extract_invoice(bogus)
        assert result.reason is ReasonCode.PARSE_ERROR


# ------------------------------------------------------------- extraction
class TestExtraction:
    def test_clean_invoice_fields(self, seeded):
        inv_dir, _ = seeded
        rec = extract_invoice(inv_dir / "invoice_01_INV-2026-1001.pdf")
        assert isinstance(rec, InvoiceRecord)
        assert rec.vendor_name == "Acme Industrial Supply Inc."
        assert rec.invoice_date == date(2026, 10, 1)
        assert rec.due_date == date(2026, 10, 31)
        assert len(rec.line_items) == 3
        assert rec.total == D("1209.15")

    def test_all_twelve_pdfs_accounted_for(self, seeded):
        inv_dir, _ = seeded
        batch = extract_directory(inv_dir)
        assert len(batch.records) + len(batch.failures) == 12
        assert len(batch.records) == 10


# --------------------------------------------------------------- duplicates
class TestDuplicates:
    def test_seeded_duplicate_detected(self, seeded):
        inv_dir, _ = seeded
        batch = extract_directory(inv_dir)
        dupes = [f for f in batch.failures if f.reason is ReasonCode.DUPLICATE_INVOICE_ID]
        assert [d.invoice_id for d in dupes] == [seed_invoices.EDGE_CASES["duplicate"]]
        assert "resend" in dupes[0].source_file
        assert sum(r.invoice_id == "INV-2026-1003" for r in batch.records) == 1

    def test_conflicting_duplicate_is_called_out(self):
        a = make_invoice("INV-9", total="100.00", source="a.pdf")
        b = make_invoice("INV-9", total="250.00", on=BASE_DATE + timedelta(days=1), source="b.pdf")
        kept, dupes = flag_duplicates([b, a])
        assert [k.source_file for k in kept] == ["a.pdf"]
        assert "DIFFERENT content" in dupes[0].errors[0]

    def test_duplicate_is_never_approved(self, pipeline):
        report, _ = pipeline
        approved_ids = [r.invoice_id for r in report.approved]
        assert len(approved_ids) == len(set(approved_ids))


# ------------------------------------------------------------ fuzzy vendor
class TestFuzzyVendor:
    @pytest.mark.parametrize("raw,expected", [
        ("Acme Industrial Supply Inc.", "acme industrial supply"),
        ("ACH PAYMENT ACME INDUSTRIAL SUPPLY INC", "acme industrial supply"),
        ("CHECK 10482 BRIGHTLINE ELECTRICAL", "brightline electrical"),
        ("Pinnacle Print & Design", "pinnacle print design"),
    ])
    def test_normalization(self, raw, expected):
        assert normalize_vendor(raw) == expected

    def test_abbreviation_scores_high(self):
        assert vendor_similarity("Northwind Logistics LLC", "ACH PMT NORTHWIND LOGIST") >= 90

    def test_ampersand_vs_and_is_exact(self):
        assert vendor_similarity("Pinnacle Print & Design", "ACH PAYMENT PINNACLE PRINT AND DESIGN") == 100

    def test_heavier_abbreviation_scores_lower(self):
        sim = vendor_similarity("Meridian Engineering Group Inc.", "ACH PMT MERIDIAN ENG GRP")
        assert 60 <= sim < 80

    def test_unrelated_vendor_scores_low(self):
        assert vendor_similarity("Cobalt Cloud Services", "ACH PMT NORTHWIND LOGIST") < 50

    def test_seeded_abbreviations_still_match(self, pipeline):
        report, _ = pipeline
        approved = report.by_invoice_id(seed_invoices.EDGE_CASES["abbreviated_approved"])[0]
        review = report.by_invoice_id(seed_invoices.EDGE_CASES["abbreviated_review"])[0]
        assert approved.status == "approved" and approved.ledger_description == "ACH PMT NORTHWIND LOGIST"
        assert ReasonCode.FUZZY_VENDOR in approved.reasons
        assert review.status == "needs_review" and review.ledger_txn_id is not None
        assert ReasonCode.LOW_VENDOR_SIMILARITY in review.reasons


# ---------------------------------------------------------- scoring/routing
class TestConfidenceRouting:
    def test_perfect_match_scores_100_and_approves(self):
        report = reconcile([make_invoice()], [make_entry()])
        (r,) = report.approved
        assert r.confidence == 100.0 and r.reasons == [ReasonCode.EXACT_MATCH]

    @pytest.mark.parametrize("days,expected", [(0, 100.0), (2, 100.0), (-2, 100.0), (3, 96.7), (5, 90.0)])
    def test_date_component(self, days, expected):
        c = score_pair(make_invoice(), make_entry(days=days))
        assert c is not None and c.score == pytest.approx(expected, abs=0.05)

    def test_outside_date_window_is_not_a_candidate(self):
        assert score_pair(make_invoice(), make_entry(days=6)) is None
        report = reconcile([make_invoice()], [make_entry(days=-6)])
        assert report.review[0].reasons == [ReasonCode.NO_LEDGER_MATCH]
        assert report.review[0].confidence == 0

    def test_amount_must_match_to_the_cent(self):
        assert score_pair(make_invoice(total="1080.00"), make_entry(amount="1080.01")) is None

    def test_right_amount_wrong_vendor_goes_to_review(self):
        report = reconcile([make_invoice()], [make_entry(desc="ACH PMT TOTALLY DIFFERENT HOLDINGS")])
        (r,) = report.review
        assert r.confidence < APPROVE_THRESHOLD
        assert ReasonCode.LOW_VENDOR_SIMILARITY in r.reasons

    def test_high_score_but_weak_vendor_is_not_approved(self):
        # 50 + 20 + 30*0.70 = 91 >= threshold, but similarity < 80 => review.
        inv = make_invoice(vendor="Meridian Engineering Group")
        c = score_pair(inv, make_entry(desc="MERIDIAN ENG GRP"))
        assert c.score >= APPROVE_THRESHOLD and c.similarity < 80
        assert reconcile([inv], [make_entry(desc="MERIDIAN ENG GRP")]).approved == []

    def test_ambiguous_candidates_go_to_review(self):
        inv = make_invoice()
        ledger = [make_entry("TXN-A", days=0), make_entry("TXN-B", days=1)]
        (r,) = reconcile([inv], ledger).review
        assert ReasonCode.MULTIPLE_CANDIDATES in r.reasons
        assert r.ledger_txn_id == "TXN-A"

    def test_ledger_entry_settles_only_one_invoice(self):
        a = make_invoice("INV-A", on=BASE_DATE)
        b = make_invoice("INV-B", on=BASE_DATE + timedelta(days=4))
        report = reconcile([a, b], [make_entry("TXN-1", days=0)])
        assert [r.invoice_id for r in report.approved] == ["INV-A"]
        assert report.by_invoice_id("INV-B")[0].reasons == [ReasonCode.NO_LEDGER_MATCH]

    def test_approved_result_requires_ledger_reference(self):
        with pytest.raises(ValidationError):
            ReconciliationResult(status="approved", confidence=99, source_file="x.pdf")

    def test_confidence_bounded(self):
        with pytest.raises(ValidationError):
            ReconciliationResult(status="needs_review", confidence=101, source_file="x.pdf")


# --------------------------------------------------------------- end to end
class TestEndToEnd:
    def test_routing_counts(self, pipeline):
        report, _ = pipeline
        assert len(report.approved) == 8
        assert len(report.review) == 4
        reasons = {r.invoice_id: r.reasons for r in report.review}
        assert ReasonCode.MATH_ERROR in reasons[seed_invoices.EDGE_CASES["math_error"]]
        assert ReasonCode.DUPLICATE_INVOICE_ID in reasons[seed_invoices.EDGE_CASES["duplicate"]]
        assert reasons[seed_invoices.EDGE_CASES["unmatched"]] == [ReasonCode.NO_LEDGER_MATCH]

    def test_rejected_invoice_payment_surfaces_as_unmatched_ledger(self, pipeline):
        report, _ = pipeline
        assert [e.description for e in report.unmatched_ledger] == ["ACH PAYMENT PINNACLE PRINT AND DESIGN"]

    def test_output_files(self, pipeline):
        report, out = pipeline
        approved = json.loads((out / "approved_matches.json").read_text())
        review = json.loads((out / "review_queue.json").read_text())
        assert approved["count"] == 8 and len(approved["items"]) == 8
        assert review["count"] == 4
        assert all(i["status"] == "approved" for i in approved["items"])

        wb = load_workbook(out / "reconciliation_report.xlsx")
        assert wb.sheetnames == ["Summary", "Approved", "Review Queue", "Unmatched Ledger"]
        assert wb["Approved"].max_row == 9
        assert wb["Review Queue"].max_row == 5

    def test_ledger_loader_skips_bad_rows(self, tmp_path):
        p = tmp_path / "ledger.csv"
        p.write_text(
            "txn_id,posted_date,description,amount\n"
            "T1,2026-10-01,ACH ACME,\"$1,200.00\"\n"
            "T2,not-a-date,ACH ACME,10.00\n"
            "T3,2026-10-02,ACH ACME,abc\n"
            "T1,2026-10-03,ACH ACME,5.00\n"
            "T4,2026-10-04,ACH ACME,-75.25\n"
        )
        result = load_ledger(p)
        assert [(e.txn_id, e.amount) for e in result.entries] == [("T1", D("1200.00")), ("T4", D("75.25"))]
        assert len(result.errors) == 3
