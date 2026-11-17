# Invoice-to-Ledger Reconciliation

Local-first pipeline: PDF invoices → validated records → matched against a bank ledger → routed to approval or review.

```bash
pip install -r requirements.txt
python seed_invoices.py          # 12 PDFs in ./sample_invoices + bank_ledger.csv
python reconciler.py -v          # writes ./output/{approved_matches,review_queue}.json + reconciliation_report.xlsx
pytest
```

| Module | Role |
|---|---|
| `schemas.py` | Pydantic v2 models; math checks with $0.01 tolerance |
| `extractor.py` | pdfplumber → raw dict → `InvoiceRecord` or `ExtractionFailure`; duplicate detection |
| `reconciler.py` | Scoring, one-to-one matching, routing, JSON + Excel output |
| `seed_invoices.py` | Sample data with the edge cases listed in `EDGE_CASES` |

**Scoring:** amount must match to the cent and the payment must post within ±5 days. Score = 50 (amount) + up to 20 (date: full credit within 2 days, 10 at day 5) + up to 30 (vendor similarity). Approval needs score ≥ 90, vendor similarity ≥ 80, and no competing ledger row within 5 points. Everything else goes to the review queue.

## GitHub Actions

`.github/workflows/pipeline.yml` runs the tests on Python 3.11–3.13 for every push to `main` and every pull request. It then runs the reconciliation and uploads `output/` (the JSON queues and the Excel report) as the `reconciliation-output` artifact. To run it by hand, open **Actions → Reconciliation pipeline → Run workflow**.
