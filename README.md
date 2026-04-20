# RYD — Data Upload

Streamlit app that ingests per-table Excel uploads into a Supabase Postgres DB. Used as a temporary upload interface while a longer-term solution is built.

## What it does

Accepts one Excel file per upload, validates headers against the selected table template, and runs:

1. **Raw load** — bulk insert into `raw.*` with an `ingest.import_batches` record for traceability.
2. **Dim refresh** — idempotent full-scan population of `core.dim_*` via `ON CONFLICT DO NOTHING`.
3. **Fact refresh** — per-table. Tables with unique keys use `ON CONFLICT`; tables without (activities, claims, INEGI) use `TRUNCATE + INSERT` from full raw history.

Supported tables: Leads, Lead Activities, Finance Applications, Market Prices, Sales, Claims, INEGI Sales.

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
├── app.py                           # single-file Streamlit app (UI + ETL)
├── requirements.txt
├── .env.example
├── .streamlit/
│   └── secrets.toml.example
├── RYD-Schema.rtf                   # reference DDL for the Supabase schema
└── etl/                             # legacy CLI pipeline; kept for reference, not used by the app
```
