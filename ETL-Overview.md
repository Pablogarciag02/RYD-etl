# RYD ETL Pipeline - Overview

## What is this?

A Python-based ETL (Extract, Transform, Load) pipeline that takes the RYD Excel data files and loads them into the Supabase database, fully normalized and ready for AI querying.

Replaces the manual process of copying data from Excel into the database.

---

## How it works

```
Excel File (from providers)
        │
        ▼
┌──────────────────────┐
│  1. VALIDATE         │  Per-sheet contract check: headers in order,
│     (Python/openpyxl) │  key column completeness. Logged to
│                      │  ingest.import_validation_results.
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  2. IDEMPOTENCY      │  If same checksum + sheet was already loaded
│     (lookup)         │  successfully, skip and log a `skipped` sheet_run.
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  3. EXTRACT          │  Reads the configured sheet, maps columns
│     (Python/openpyxl) │  to the database schema.
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  4. LOAD RAW         │  Inserts rows into raw.* with savepoint per chunk.
│     (Postgres)       │  Reconciles loaded count vs. db_count_for_batch.
│                      │  Tracks each import in ingest.import_batches
│                      │  and ingest.import_sheet_runs.
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  5. TRANSFORM        │  Populates dimension tables (brands, products, etc.)
│     (SQL)            │  Populates fact tables with foreign-key lookups.
│                      │  core.dim_dealer is governed manually — never written.
└──────────────────────┘
```

### Layer details

| Layer | Schema | Purpose |
|-------|--------|---------|
| Ingestion Control | `ingest` | Tracks every import: file name, sheet, timestamp, row counts, success/failure |
| Raw | `raw` | Exact copy of the Excel data stored as text. No data loss. Full traceability back to source. |
| Core | `core` | Normalized star schema. Dimension tables (dealers, brands, products, etc.) + fact tables (leads, sales, finance apps, etc.) with proper relationships. This is what gets queried. |

---

## Data sources currently supported

| Excel Sheet | Raw Table | What it contains | Frequency |
|-------------|-----------|------------------|-----------|
| DBSolicitudes_Cetelem | `raw.raw_finance_applications` | Finance applications via Cetelem/Inbursa | Daily (Mon-Fri) |
| DBPreciosMexico_ConMG | `raw.raw_market_prices` | Market prices by model/version | Monthly |
| DBVentas_ConMG | `raw.raw_sales` | Sales by dealer (retail vs fleet) | Monthly |
| DBSiniestros_Marsh | `raw.raw_claims` | Insurance claims (Marsh) | Monthly |
| BaseINEGIAutosLigerosMexico | `raw.raw_inegi_sales` | INEGI light-vehicle market benchmark | Monthly |

### Excluded by design

- **Leads / Lead Activities** (`BaseLeads1`, `BaseLeads2`): no longer ingested through this pipeline. Historical `raw.raw_leads`, `raw.raw_lead_activities`, `core.fact_leads`, `core.fact_lead_activities` remain queryable but receive no new rows.
- **Dealer catalog** (`CatalogoTiendas ConMG`) and dealer alias mappings: managed manually in the database. The ETL never writes to `core.dim_dealer` or `core.dealer_alias_map`.
- **`core.dim_source_channel` and `core.dim_campaign`** were only ever fed from leads data. They retain their historical values; new ingestion does not extend them.

---

## How to run it

```bash
# 1. Activate the Python environment
source .venv/bin/activate

# 2. Run with default file (BasesEjemplo.xlsx)
python -m etl.run

# Or specify a different Excel file
python -m etl.run /path/to/new_file.xlsx
```

That's it. The pipeline handles everything: reading the Excel, creating import batch records, loading raw data, and transforming into the normalized core tables.

---

## Current load stats (from BasesEjemplo.xlsx)

**Raw tables:** 109,514 total rows loaded
- raw_leads: 65,784
- raw_lead_activities: 19,598
- raw_finance_applications: 10,023
- raw_market_prices: 1,502
- raw_sales: 12,607

**Core tables:** 12 dimensions + 5 fact tables populated
- 278 dealers, 68 brands, 71 groups, 3,670 advisors
- 65,784 lead facts, 57,037 activity facts, 10,023 finance facts, 36,497 sales facts

---

## Repo structure

```
agenticRYD/
├── .env                  # Supabase connection credentials
├── requirements.txt      # Python dependencies
├── app.py                # Streamlit upload UI (uses etl/* under the hood)
├── etl/
│   ├── db.py             # Database connection (CLI usage)
│   ├── extract.py        # Excel readers (one per sheet) + SHEET_MAP
│   ├── validate.py       # SHEET_CONTRACTS + per-sheet validation checks
│   ├── audit.py          # ingest.import_sheet_runs / errors / validation_results helpers
│   ├── load_raw.py       # Bulk insert + idempotency + reconciliation
│   ├── transform.py      # Dimension + fact table population
│   └── run.py            # CLI orchestrator
├── DefinicionBasesConMG.xlsx   # Sample data file
└── RYD-Schema.rtf        # Database schema DDL
```

---

## Next steps

- **Incremental loads**: Only process new/changed rows instead of full reload
- **Dealer catalog reconciliation**: Surface raw `distribuidor` strings that don't match any `core.dim_dealer.commercial_name` so the steward can resolve them
- **Scheduling**: Automate runs on a schedule matching provider delivery frequency
