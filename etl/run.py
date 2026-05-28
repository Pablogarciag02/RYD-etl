"""Main ETL orchestrator: Validate → Extract → Load Raw → Transform Core."""

import sys
import os
import time
from datetime import datetime

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from etl.db import get_connection
from etl.extract import SHEET_MAP
from etl.load_raw import (
    create_import_batch,
    find_completed_batch,
    get_batch_counts,
    bulk_insert,
    count_rows_for_batch,
    update_batch_status,
    fail_import_batch,   # add this helper in load_raw.py
)
from etl.transform import run_transforms
from etl.validate import validate_sheet_contract
from etl.audit import (
    create_sheet_run,
    update_sheet_run_status,
    record_import_error,
    record_validation_result,
)


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
        # ── Phase 1: Validate, Extract & Load Raw ───────────────────
        print("=" * 60)
        print("PHASE 1: VALIDATE, EXTRACT & LOAD RAW")
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
            started_at = datetime.utcnow()

            # 0) Idempotency: skip if this exact file+sheet was already processed successfully.
            existing_batch_id = find_completed_batch(cur, filepath, sheet_name)
            if existing_batch_id is not None:
                counts = get_batch_counts(cur, existing_batch_id) or {}
                db_count = count_rows_for_batch(cur, table_name, existing_batch_id)
                recon_ok = (db_count == counts.get("rows_loaded", 0))
                sheet_run_id = create_sheet_run(
                    cur=cur,
                    import_batch_id=existing_batch_id,
                    source_file_name=os.path.basename(filepath),
                    source_sheet_name=sheet_name,
                    target_table=table_name,
                )
                update_sheet_run_status(
                    cur=cur,
                    sheet_run_id=sheet_run_id,
                    status="skipped",
                    rows_detected=counts.get("rows_detected", 0),
                    rows_loaded=counts.get("rows_loaded", 0),
                    rows_failed=counts.get("rows_failed", 0),
                    notes=(
                        "Skipped: same file checksum + sheet was already processed."
                        if recon_ok
                        else (
                            "Skipped: previously processed, but reconciliation mismatch: "
                            f"loaded={counts.get('rows_loaded', 0)} db_count={db_count}"
                        )
                    ),
                    started_at=started_at,
                    finished_at=datetime.utcnow(),
                    metadata={
                        "existing_batch_id": str(existing_batch_id),
                        "reconciliation": {
                            "loaded": counts.get("rows_loaded", 0),
                            "db_count_for_batch": db_count,
                            "ok": recon_ok,
                        },
                    },
                )
                conn.commit()
                print(f"    SKIPPED (already processed) | Existing Batch ID: {existing_batch_id}")
                cur.close()
                continue

            # 1) Create import batch
            batch_id = create_import_batch(cur, filepath, sheet_name)
            print(f"    Batch ID: {batch_id}")

            # 2) Create sheet run audit record
            sheet_run_id = create_sheet_run(
                cur=cur,
                import_batch_id=batch_id,
                source_file_name=os.path.basename(filepath),
                source_sheet_name=sheet_name,
                target_table=table_name,
            )
            conn.commit()

            try:
                # 3) Mark as validating
                update_sheet_run_status(
                    cur=cur,
                    sheet_run_id=sheet_run_id,
                    status="validating",
                )
                conn.commit()

                # 4) Validate sheet contract
                validation = validate_sheet_contract(filepath, sheet_name)

                for check in validation.get("checks", []):
                    record_validation_result(
                        cur=cur,
                        import_batch_id=batch_id,
                        sheet_run_id=sheet_run_id,
                        source_sheet_name=sheet_name,
                        validation_stage=check["validation_stage"],
                        validation_name=check["validation_name"],
                        status=check["status"],
                        severity=check.get("severity", "error"),
                        expected_value=str(check.get("expected_value")) if check.get("expected_value") is not None else None,
                        actual_value=str(check.get("actual_value")) if check.get("actual_value") is not None else None,
                        details=check.get("details", {}),
                    )

                # 5) Stop if validation failed
                if not validation["ok"]:
                    print(f"    VALIDATION FAILED: {validation['message']}")

                    record_import_error(
                        cur=cur,
                        import_batch_id=batch_id,
                        sheet_run_id=sheet_run_id,
                        source_file_name=os.path.basename(filepath),
                        source_sheet_name=sheet_name,
                        target_table=table_name,
                        error_stage="pre_ingestion_validation",
                        error_type=validation["error_type"],
                        error_message=validation["message"],
                        error_details=validation.get("details", {}),
                        source_row_number=None,
                    )

                    fail_import_batch(cur, batch_id, validation["message"])

                    update_sheet_run_status(
                        cur=cur,
                        sheet_run_id=sheet_run_id,
                        status="failed",
                        rows_detected=0,
                        rows_loaded=0,
                        rows_failed=0,
                        notes=validation["message"],
                        started_at=started_at,
                        finished_at=datetime.utcnow(),
                        metadata=validation.get("details", {}),
                    )

                    conn.commit()
                    cur.close()
                    continue

                # 6) Mark as processing
                update_sheet_run_status(
                    cur=cur,
                    sheet_run_id=sheet_run_id,
                    status="processing",
                )
                conn.commit()

                # 7) Extract + load raw
                t0 = time.time()
                rows = extractor(filepath)
                detected, loaded, failed = bulk_insert(cur, conn, table_name, rows, batch_id)
                elapsed = time.time() - t0

                # 8) Update batch + sheet run
                update_batch_status(cur, batch_id, detected, loaded, failed)

                # 9) Reconcile with actual DB row count for this batch
                db_count = count_rows_for_batch(cur, table_name, batch_id)
                recon_ok = (db_count == loaded)
                recon_note = None
                if not recon_ok:
                    recon_note = (
                        "Row count mismatch: detected=%s loaded=%s failed=%s db_count=%s"
                        % (detected, loaded, failed, db_count)
                    )

                final_status = "success" if failed == 0 else "partial_success"
                update_sheet_run_status(
                    cur=cur,
                    sheet_run_id=sheet_run_id,
                    status=final_status,
                    rows_detected=detected,
                    rows_loaded=loaded,
                    rows_failed=failed,
                    notes=(
                        recon_note
                        if recon_note is not None
                        else (None if failed == 0 else "Some rows failed during raw load.")
                    ),
                    started_at=started_at,
                    finished_at=datetime.utcnow(),
                    metadata={
                        "elapsed_seconds": round(elapsed, 2),
                        "reconciliation": {
                            "detected": detected,
                            "loaded": loaded,
                            "failed": failed,
                            "db_count_for_batch": db_count,
                            "ok": recon_ok,
                        },
                    },
                )

                conn.commit()
                print(f"    Loaded: {loaded} | Failed: {failed} | Time: {elapsed:.1f}s")

            except Exception as sheet_error:
                conn.rollback()

                # reopen cursor after rollback
                cur = conn.cursor()

                record_import_error(
                    cur=cur,
                    import_batch_id=batch_id,
                    sheet_run_id=sheet_run_id,
                    source_file_name=os.path.basename(filepath),
                    source_sheet_name=sheet_name,
                    target_table=table_name,
                    error_stage="sheet_processing",
                    error_type="runtime_exception",
                    error_message=str(sheet_error),
                    error_details={"exception_class": sheet_error.__class__.__name__},
                    source_row_number=None,
                )

                fail_import_batch(cur, batch_id, str(sheet_error))

                update_sheet_run_status(
                    cur=cur,
                    sheet_run_id=sheet_run_id,
                    status="failed",
                    rows_detected=0,
                    rows_loaded=0,
                    rows_failed=0,
                    notes=str(sheet_error),
                    started_at=started_at,
                    finished_at=datetime.utcnow(),
                    metadata={"exception_class": sheet_error.__class__.__name__},
                )

                conn.commit()
                print(f"    ERROR: {sheet_error}")

            finally:
                cur.close()

        # ── Phase 2: Transform Raw → Core ───────────────────────────
        print("\n" + "=" * 60)
        print("PHASE 2: TRANSFORM RAW → CORE")
        print("=" * 60)

        cur = conn.cursor()
        try:
            run_transforms(cur)
            conn.commit()
        except Exception as transform_error:
            conn.rollback()

            cur = conn.cursor()
            record_import_error(
                cur=cur,
                import_batch_id=None,
                sheet_run_id=None,
                source_file_name=os.path.basename(filepath),
                source_sheet_name=None,
                target_table="core.*",
                error_stage="transform",
                error_type="transform_runtime_exception",
                error_message=str(transform_error),
                error_details={"exception_class": transform_error.__class__.__name__},
                source_row_number=None,
            )
            conn.commit()
            cur.close()
            raise

        # ── Summary ─────────────────────────────────────────────────
        print("\n" + "=" * 60)
        print("SUMMARY")
        print("=" * 60)

        cur = conn.cursor()
        for schema in ["ingest", "raw", "core"]:
            cur.execute(
                """
                SELECT table_name
                FROM information_schema.tables
                WHERE table_schema = %s
                ORDER BY table_name
                """,
                (schema,),
            )
            tables = cur.fetchall()
            print(f"\n  {schema.upper()}:")
            for (tbl,) in tables:
                cur.execute(f"SELECT COUNT(*) FROM {schema}.{tbl}")
                row = cur.fetchone()
                count = row[0] if row else 0
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