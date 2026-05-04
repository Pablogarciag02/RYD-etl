# RYD — Data Upload

Streamlit app that ingests per-table Excel uploads into a Supabase Postgres DB. Used as a temporary upload interface while a longer-term solution is built.

## What it does

Accepts one Excel file per upload, validates the sheet against a per-table contract (expected headers in order, key-column completeness threshold), and runs:

1. **Validate** — `etl/validate.py` runs structural checks against the workbook and records every check (pass/warn/fail) into `ingest.import_validation_results`.
2. **Idempotency** — if the same file checksum + sheet was already loaded successfully, the upload is skipped.
3. **Raw load** — bulk insert into `raw.*` (with savepoints so a single bad chunk doesn't poison the transaction). Each load creates an `ingest.import_batches` row plus an `ingest.import_sheet_runs` audit record.
4. **Reconcile** — compare reported `rows_loaded` against the actual count of rows tagged with the new `batch_id` in the target table.
5. **Dim refresh** — idempotent full-scan population of `core.dim_*` via `ON CONFLICT DO NOTHING`. `core.dim_dealer` is *not* touched — it is treated as manually-governed master data.
6. **Fact refresh** — per-table. Tables with unique keys use `ON CONFLICT`; tables without (claims, INEGI) use `TRUNCATE + INSERT` from full raw history.

Supported tables: Finance Applications, Market Prices, Sales, Claims, INEGI Sales.

> Leads / Lead Activities and the dealer catalog (CatalogoTiendas) are intentionally excluded. Leads no longer flow through this pipeline; the dealer catalog and any dealer alias mappings are governed manually inside the database.

## Local development

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# Fill in DB_* and APP_PASSWORD in .env

streamlit run app.py
```

Opens at `http://localhost:8501`. Enter the app password to access the upload form.

## Deploying to Streamlit Cloud

1. Push this repo to GitHub (no real data — all `.xlsx` files are gitignored).
2. Go to [share.streamlit.io](https://share.streamlit.io), connect the repo, set the main file to `app.py`.
3. In the app's **Settings → Secrets** panel, paste a TOML block with the same keys as `.streamlit/secrets.toml.example`, using real values.
4. Deploy. Share the URL + password with the client via a secure channel.

## Security notes

- **The password is a shared secret, not real auth.** Rotate it before sharing with a new party; revoke by changing the secret value.
- **Never commit `.env`, `.streamlit/secrets.toml`, or any `.xlsx` file.** The gitignore already blocks these.
- The app only writes to the DB; it never reads or displays customer data.

## Project layout

```
.
├── app.py                           # Streamlit UI shell (uses etl/* under the hood)
├── etl/
│   ├── db.py                        # Postgres connection (CLI usage)
│   ├── extract.py                   # Per-sheet Excel readers + SHEET_MAP
│   ├── validate.py                  # SHEET_CONTRACTS + validate_sheet_contract
│   ├── audit.py                     # Writes to ingest.import_sheet_runs / errors / validation_results
│   ├── load_raw.py                  # Bulk insert + idempotency + reconciliation
│   ├── transform.py                 # Dimension and fact populators
│   └── run.py                       # CLI orchestrator (python -m etl.run path/to.xlsx)
├── requirements.txt
├── .env.example
├── .streamlit/
│   └── secrets.toml.example
└── RYD-Schema.rtf                   # reference DDL for the Supabase schema
```

Both the Streamlit app and `python -m etl.run` go through the same validate / load / transform / audit primitives in `etl/`.
