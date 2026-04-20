"""Main ETL orchestrator: Extract → Load Raw → Transform Core."""

import sys
import os
import time

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from etl.db import get_connection
from etl.extract import SHEET_MAP
from etl.load_raw import create_import_batch, bulk_insert, update_batch_status
from etl.transform import run_transforms


def run_etl(filepath, sheets=None):
    """
    Run the full ETL pipeline for the given Excel file.

    Args:
        filepath: Path to the Excel file.
        sheets:   List of sheet names to process. None = all available.
    """
    if not os.path.exists(filepath):
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    conn = get_connection()
    conn.autocommit = False

    try:
        # ── Phase 1: Extract & Load Raw ─────────────────────────
        print("=" * 60)
        print("PHASE 1: EXTRACT & LOAD RAW")
        print("=" * 60)

        available = SHEET_MAP.keys()
        to_process = sheets if sheets else list(available)

        for sheet_name in to_process:
            if sheet_name not in SHEET_MAP:
                print(f"\n  SKIP: '{sheet_name}' not in SHEET_MAP")
                continue

            extractor, table_name = SHEET_MAP[sheet_name]
            print(f"\n  Processing: {sheet_name} → {table_name}")

            cur = conn.cursor()
            batch_id = create_import_batch(cur, filepath, sheet_name)
            print(f"    Batch ID: {batch_id}")

            t0 = time.time()
            rows = extractor(filepath)
            loaded, failed = bulk_insert(cur, table_name, rows, batch_id)
            elapsed = time.time() - t0

            update_batch_status(cur, batch_id, loaded + failed, loaded, failed)
            conn.commit()
            print(f"    Loaded: {loaded} | Failed: {failed} | Time: {elapsed:.1f}s")

        # ── Phase 2: Transform Raw → Core ───────────────────────
        print("\n" + "=" * 60)
        print("PHASE 2: TRANSFORM RAW → CORE")
        print("=" * 60)

        cur = conn.cursor()
        run_transforms(cur)
        conn.commit()

        # ── Summary ─────────────────────────────────────────────
        print("\n" + "=" * 60)
        print("SUMMARY")
        print("=" * 60)

        cur = conn.cursor()
        for schema in ["ingest", "raw", "core"]:
            cur.execute("""
                SELECT table_name FROM information_schema.tables
                WHERE table_schema = %s ORDER BY table_name
            """, (schema,))
            tables = cur.fetchall()
            print(f"\n  {schema.upper()}:")
            for (tbl,) in tables:
                cur.execute(f"SELECT COUNT(*) FROM {schema}.{tbl}")
                count = cur.fetchone()[0]
                if count > 0:
                    print(f"    {tbl}: {count:,} rows")

        cur.close()

    except Exception as e:
        conn.rollback()
        print(f"\nERROR: {e}")
        raise
    finally:
        conn.close()

    print("\nETL complete.")


if __name__ == "__main__":
    # Default path to Excel file
    default_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "DefinicionBasesConMG.xlsx",
    )
    filepath = sys.argv[1] if len(sys.argv) > 1 else default_path
    run_etl(filepath)
