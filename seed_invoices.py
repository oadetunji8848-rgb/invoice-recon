"""Generate 12 realistic sample PDF invoices plus a matching bank ledger CSV.

Deliberate edge cases (see ``EDGE_CASES``):
  * INV-2026-1009  tax/total math error (total overstated by $15.00)
  * INV-2026-1003  duplicate: the vendor re-sent the same invoice
  * INV-2026-1002  ledger payee abbreviated ("NORTHWIND LOGIST")       -> still approved
  * INV-2026-1004  ledger payee abbreviated ("MERIDIAN ENG GRP") + 4 day drift -> review
  * INV-2026-1011  no ledger payment at all                             -> unmatched

Usage:
    python seed_invoices.py [--out ./sample_invoices] [--ledger ./bank_ledger.csv]
"""

from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from fpdf import FPDF
from fpdf.fonts import FontFace

CENT = Decimal("0.01")
BILL_TO = ("Harborview Analytics Ltd.", "400 Quayside Avenue, Suite 12", "Portland, OR 97209")


def money(v: Decimal) -> str:
    return f"${v:,.2f}"


@dataclass
class SeedInvoice:
    file_stem: str
    invoice_id: str
    vendor: str
    vendor_address: str
    invoice_date: date
    items: list[tuple[str, int, str]]  # (description, qty, unit price)
    tax_rate: Decimal
    ledger_payee: str | None  # None => no payment in the ledger
    ledger_day_offset: int = 0
    total_error: Decimal = Decimal("0")  # added to printed total to create a math error
    in_ledger_as_separate_payment: bool = True
    terms_days: int = 30
    _computed: dict = field(default_factory=dict, repr=False)

    def amounts(self) -> tuple[list[tuple[str, int, Decimal, Decimal]], Decimal, Decimal, Decimal]:
        rows = []
        for desc, qty, price in self.items:
            p = Decimal(price)
            rows.append((desc, qty, p, (p * qty).quantize(CENT)))
        subtotal = sum((r[3] for r in rows), Decimal("0"))
        tax = (subtotal * self.tax_rate).quantize(CENT, rounding=ROUND_HALF_UP)
        total = subtotal + tax + self.total_error
        return rows, subtotal, tax, total


INVOICES: list[SeedInvoice] = [
    SeedInvoice(
        "invoice_01", "INV-2026-1001", "Acme Industrial Supply Inc.", "18 Foundry Road, Tacoma, WA 98402",
        date(2026, 10, 1),
        [("Nitrile gloves, case of 1000", 6, "84.50"), ("Safety glasses, anti-fog", 40, "7.25"),
         ("Hi-vis vest, class 2", 25, "12.80")],
        Decimal("0.0825"), "ACH PAYMENT ACME INDUSTRIAL SUPPLY INC", 2,
    ),
    SeedInvoice(
        "invoice_02", "INV-2026-1002", "Northwind Logistics LLC", "77 Harbor Freightway, Oakland, CA 94607",
        date(2026, 10, 3),
        [("LTL freight, Oakland to Portland", 3, "640.00"), ("Liftgate service", 3, "75.00"),
         ("Fuel surcharge", 1, "118.40")],
        Decimal("0.00"), "ACH PMT NORTHWIND LOGIST", 1,
    ),
    SeedInvoice(
        "invoice_03", "INV-2026-1003", "Cobalt Cloud Services", "1200 Pine Street, Seattle, WA 98101",
        date(2026, 10, 5),
        [("Compute instances, October", 1, "2340.00"), ("Object storage (TB-month)", 18, "21.50"),
         ("Premium support plan", 1, "450.00")],
        Decimal("0.065"), "WIRE TRANSFER COBALT CLOUD SERVICES", 0,
    ),
    SeedInvoice(
        "invoice_04", "INV-2026-1004", "Meridian Engineering Group Inc.", "9 Ridgeline Plaza, Boise, ID 83702",
        date(2026, 10, 8),
        [("Structural assessment, Bldg C", 1, "3800.00"), ("Site visit, senior engineer (hrs)", 12, "185.00")],
        Decimal("0.06"), "ACH PMT MERIDIAN ENG GRP", 4,
    ),
    SeedInvoice(
        "invoice_05", "INV-2026-1005", "Vertex Data Labs Ltd.", "55 Circuit Lane, Austin, TX 78701",
        date(2026, 10, 10),
        [("Annotation services (hrs)", 120, "38.00"), ("QA review pass", 1, "640.00")],
        Decimal("0.0825"), "ACH PAYMENT VERTEX DATA LABS LTD", 3,
    ),
    SeedInvoice(
        "invoice_06", "INV-2026-1006", "Greenleaf Facilities Management", "310 Elm Court, Portland, OR 97205",
        date(2026, 10, 12),
        [("Janitorial service, monthly", 1, "1950.00"), ("Window cleaning", 2, "275.00"),
         ("Restroom supplies restock", 1, "186.35")],
        Decimal("0.00"), "ACH PAYMENT GREENLEAF FACILITIES MANAGEMENT", 1,
    ),
    SeedInvoice(
        "invoice_07", "INV-2026-1007", "Brightline Electrical Contractors", "2 Voltage Way, Salem, OR 97301",
        date(2026, 10, 15),
        [("Panel upgrade, 200A", 1, "2650.00"), ("LED troffer retrofit", 30, "96.00"),
         ("Permit and inspection fee", 1, "210.00")],
        Decimal("0.00"), "CHECK 10482 BRIGHTLINE ELECTRICAL CONTRACTORS", 4,
    ),
    SeedInvoice(
        "invoice_08", "INV-2026-1008", "Oakridge Office Furniture Co.", "640 Timber Blvd, Eugene, OR 97401",
        date(2026, 10, 18),
        [("Ergonomic task chair", 14, "329.00"), ("Sit-stand desk, 60in", 6, "689.00"),
         ("Delivery and assembly", 1, "350.00")],
        Decimal("0.00"), "ACH PAYMENT OAKRIDGE OFFICE FURNITURE CO", 2,
    ),
    SeedInvoice(  # EDGE CASE: math error, total overstated by $15.00
        "invoice_09", "INV-2026-1009", "Pinnacle Print & Design", "88 Press Street, Spokane, WA 99201",
        date(2026, 10, 20),
        [("Tri-fold brochures, 2,000 qty", 2, "315.00"), ("Trade show banner 8ft", 3, "142.50"),
         ("Design hours", 6, "95.00")],
        Decimal("0.089"), "ACH PAYMENT PINNACLE PRINT AND DESIGN", 2, total_error=Decimal("15.00"),
    ),
    SeedInvoice(  # EDGE CASE: duplicate of invoice_03 (vendor re-sent it)
        "invoice_10_resend", "INV-2026-1003", "Cobalt Cloud Services", "1200 Pine Street, Seattle, WA 98101",
        date(2026, 10, 5),
        [("Compute instances, October", 1, "2340.00"), ("Object storage (TB-month)", 18, "21.50"),
         ("Premium support plan", 1, "450.00")],
        Decimal("0.065"), None,  # the ledger only has one payment for INV-2026-1003
    ),
    SeedInvoice(  # EDGE CASE: unmatched - never paid
        "invoice_11", "INV-2026-1011", "Bluebird Catering Co.", "14 Orchard Row, Portland, OR 97214",
        date(2026, 10, 24),
        [("Working lunch, 35 guests", 35, "22.50"), ("Coffee service", 2, "95.00"),
         ("Staffing (hrs)", 4, "45.00")],
        Decimal("0.00"), None,
    ),
    SeedInvoice(
        "invoice_12", "INV-2026-1012", "Summit Legal Advisors LLP", "1 Courthouse Square, Portland, OR 97204",
        date(2026, 10, 27),
        [("Contract review (hrs)", 9, "420.00"), ("Filing fees (pass-through)", 1, "185.00")],
        Decimal("0.00"), "WIRE TRANSFER SUMMIT LEGAL ADVISORS LLP", 3,
    ),
]

EDGE_CASES = {
    "math_error": "INV-2026-1009",
    "duplicate": "INV-2026-1003",
    "abbreviated_approved": "INV-2026-1002",
    "abbreviated_review": "INV-2026-1004",
    "unmatched": "INV-2026-1011",
}


class InvoicePDF(FPDF):
    def footer(self) -> None:
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def render_invoice(inv: SeedInvoice, path: Path) -> None:
    rows, subtotal, tax, total = inv.amounts()
    pdf = InvoicePDF(format="Letter")
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_creator("seed_invoices.py")
    pdf.set_title(f"Invoice {inv.invoice_id}")
    pdf.add_page()

    # Letterhead: vendor name is always the first line of the document.
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 9, inv.vendor, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(90)
    pdf.cell(0, 5, inv.vendor_address, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    pdf.set_text_color(0)
    pdf.set_font("Helvetica", "B", 22)
    pdf.cell(0, 10, "INVOICE", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    due = inv.invoice_date + timedelta(days=inv.terms_days)
    pdf.set_font("Helvetica", "", 10)
    meta = [
        ("Invoice Number:", inv.invoice_id),
        ("Invoice Date:", inv.invoice_date.strftime("%b %d, %Y")),
        ("Due Date:", due.strftime("%b %d, %Y")),
        ("Terms:", f"Net {inv.terms_days}"),
        ("Currency:", "USD"),
    ]
    for label, value in meta:
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(35, 6, label)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, value, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Bill To:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    for line in BILL_TO:
        pdf.cell(0, 5, line, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    pdf.set_font("Helvetica", "", 10)
    header_style = FontFace(emphasis="BOLD", color=255, fill_color=(20, 40, 80))
    with pdf.table(
        col_widths=(95, 20, 35, 35),
        text_align=("LEFT", "RIGHT", "RIGHT", "RIGHT"),
        headings_style=header_style,
        line_height=7,
    ) as table:
        head = table.row()
        for h in ("Description", "Qty", "Unit Price", "Amount"):
            head.cell(h)
        for desc, qty, price, amount in rows:
            r = table.row()
            r.cell(desc)
            r.cell(str(qty))
            r.cell(money(price))
            r.cell(money(amount))
    pdf.ln(4)

    tax_label = f"Tax ({inv.tax_rate * 100:.2f}%):" if inv.tax_rate else "Tax (exempt):"
    for label, value, bold in (
        ("Subtotal:", subtotal, False),
        (tax_label, tax, False),
        ("Total Due:", total, True),
    ):
        pdf.set_font("Helvetica", "B" if bold else "", 11 if bold else 10)
        pdf.cell(150, 7, label, align="R")
        pdf.cell(35, 7, money(value), align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(10)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(90)
    pdf.multi_cell(
        0, 5,
        f"Please remit payment by ACH or wire, quoting invoice {inv.invoice_id}. "
        "Thank you for your business.",
    )
    pdf.output(str(path))


def build_ledger(invoices: list[SeedInvoice]) -> list[dict[str, str]]:
    entries = []
    for inv in invoices:
        if inv.ledger_payee is None:
            continue
        _, _, _, total = inv.amounts()
        entries.append({
            "posted_date": (inv.invoice_date + timedelta(days=inv.ledger_day_offset)).isoformat(),
            "description": inv.ledger_payee,
            "amount": f"{total:.2f}",
        })
    entries.sort(key=lambda e: (e["posted_date"], e["description"]))
    for i, e in enumerate(entries, start=1):
        e["txn_id"] = f"TXN-{88000 + i}"
    return [{k: e[k] for k in ("txn_id", "posted_date", "description", "amount")} for e in entries]


def generate(out_dir: Path, ledger_path: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.pdf"):
        stale.unlink()
    paths = []
    for inv in INVOICES:
        p = out_dir / f"{inv.file_stem}_{inv.invoice_id}.pdf"
        render_invoice(inv, p)
        paths.append(p)

    rows = build_ledger(INVOICES)
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    with ledger_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["txn_id", "posted_date", "description", "amount"])
        writer.writeheader()
        writer.writerows(rows)
    return paths


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", type=Path, default=Path("sample_invoices"))
    ap.add_argument("--ledger", type=Path, default=Path("bank_ledger.csv"))
    args = ap.parse_args()
    paths = generate(args.out, args.ledger)
    print(f"Wrote {len(paths)} invoices to {args.out}/ and ledger to {args.ledger}")
    for case, inv_id in EDGE_CASES.items():
        print(f"  edge case {case:<22} {inv_id}")


if __name__ == "__main__":
    main()
