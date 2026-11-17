# Invoice-to-Ledger Reconciliation

Matches PDF invoices against a bank export and tells you which payments are confirmed and which need a human look: wrong maths, duplicate invoices, unpaid invoices, and payments whose vendor name or date doesn't quite line up.

## Use it on GitHub (no installation)

> ⚠️ **Make the repository private first** if you'll upload real invoices or bank statements: *Settings → General → Danger Zone → Change visibility*.

1. **Upload your invoices.** Open the `data/invoices` folder → **Add file → Upload files** → drop in your PDFs → **Commit changes**.
2. **Upload your bank export.** Open the `data` folder → **Add file → Upload files** → upload your CSV, **named `bank_ledger.csv`** → **Commit changes**.
3. **See the results.** Open the **Actions** tab → the latest **Reconcile invoices** run. The page lists what was approved, what needs review and why, and bank payments with no invoice. Download **reconciliation-report** (bottom of the page) for the Excel workbook.

Uploading new files reruns it automatically. To rerun by hand, or to point it at a different folder or file: **Actions → Reconcile invoices → Run workflow**.

**Try the demo first:** Actions → Reconcile invoices → Run workflow, set *invoices* to `sample_invoices` and *ledger* to `bank_ledger.csv`.

### What your files need to look like

**Invoices:** text-based PDFs (not scans or photos). Each needs an invoice number, an invoice date, line items, a subtotal, tax (or VAT) and a total. The vendor is read from a `From:` / `Vendor:` / `Supplier:` label, or else the first line of the invoice. `$`, `₦`, `£`, `€` and codes like `NGN` are recognised.

**Bank export:** a CSV with a date column, a description column and an amount column. Common bank headings work as they are: for example *Transaction Date / Value Date*, *Narration / Description / Payee*, *Debit / Amount / Withdrawal*, *Reference*. Deposits in a separate *Credit* column are ignored. Dates like `03/04/2026` are read day-first unless every date only makes sense month-first; you can force either with the *date_order* option when you run it by hand.

## How matching works

A bank payment is only considered for an invoice if the **amount matches to the cent** and it was **posted within 5 days** of the invoice date. Each candidate gets a 0–100 confidence score:

| Part | Points |
|---|---|
| Exact amount | 50 |
| Date closeness | 20 within 2 days, falling to 10 at 5 days |
| Vendor name similarity | up to 30 (ignores Ltd/Inc/LLC and bank wording like "ACH PMT", "TRF TO") |

**Auto-approved** if the score is at least 90, the vendor names are at least 80% similar, and no other payment scores within 5 points. Everything else goes to the **review queue** with a reason: `MATH_ERROR`, `DUPLICATE_INVOICE_ID`, `NO_LEDGER_MATCH`, `LOW_VENDOR_SIMILARITY`, `DATE_DRIFT`, `MULTIPLE_CANDIDATES`, `SCHEMA_ERROR` (a required field couldn't be found) or `PARSE_ERROR` (unreadable PDF).

## Run it on your own computer

```bash
pip install -r requirements.txt
python reconciler.py --invoices path/to/pdfs --ledger path/to/bank.csv --out output
```

Results land in `output/`: `approved_matches.json`, `review_queue.json`, `reconciliation_report.xlsx` and `summary.md`. `python seed_invoices.py` regenerates the sample data; `pytest` runs the tests.

| Module | Role |
|---|---|
| `schemas.py` | Pydantic v2 models; invoice maths checked to $0.01 |
| `extractor.py` | PDF → validated invoice, or a failure with a reason; duplicate detection |
| `reconciler.py` | Bank CSV loading, scoring, matching, JSON/Excel/Markdown output |
| `seed_invoices.py` | Sample invoices and ledger with built-in edge cases |
