# etl/audit.py
"""Audit helpers for ingestion and validation logging."""

import json
from datetime import datetime
from typing import Any, Dict, Optional


def create_sheet_run(cur, import_batch_id, source_file_name, source_sheet_name, target_table):
    """
    Create a sheet-level execution record and return its id.
    """
    cur.execute(
        """
        INSERT INTO ingest.import_sheet_runs
            (import_batch_id, source_file_name, source_sheet_name, target_table, status)
        VALUES (%s, %s, %s, %s, %s)
        RETURNING id
        """,
        (
            import_batch_id,
            source_file_name,
            source_sheet_name,
            target_table,
            "pending",
        ),
    )
    return cur.fetchone()[0]


def update_sheet_run_status(
    cur,
    sheet_run_id,
    status,
    rows_detected: Optional[int] = None,
    rows_loaded: Optional[int] = None,
    rows_failed: Optional[int] = None,
    notes: Optional[str] = None,
    started_at: Optional[datetime] = None,
    finished_at: Optional[datetime] = None,
    metadata: Optional[Dict[str, Any]] = None,
):
    """
    Update a sheet run record.
    """
    duration_ms = None
    if started_at and finished_at:
        duration_ms = int((finished_at - started_at).total_seconds() * 1000)

    cur.execute(
        """
        UPDATE ingest.import_sheet_runs
        SET
            status = %s,
            rows_detected = COALESCE(%s, rows_detected),
            rows_loaded = COALESCE(%s, rows_loaded),
            rows_failed = COALESCE(%s, rows_failed),
            notes = COALESCE(%s, notes),
            finished_at = COALESCE(%s, finished_at),
            duration_ms = COALESCE(%s, duration_ms),
            metadata = COALESCE(%s::jsonb, metadata),
            updated_at = now()
        WHERE id = %s
        """,
        (
            status,
            rows_detected,
            rows_loaded,
            rows_failed,
            notes,
            finished_at,
            duration_ms,
            json.dumps(metadata) if metadata is not None else None,
            sheet_run_id,
        ),
    )


def record_import_error(
    cur,
    import_batch_id,
    sheet_run_id,
    source_file_name,
    source_sheet_name,
    target_table,
    error_stage,
    error_type,
    error_message,
    error_details: Optional[Dict[str, Any]] = None,
    source_row_number: Optional[int] = None,
):
    """
    Persist an import/runtime/validation error.
    """
    cur.execute(
        """
        INSERT INTO ingest.import_errors
            (import_batch_id, sheet_run_id, source_file_name, source_sheet_name,
             target_table, source_row_number, error_stage, error_type,
             error_message, error_details)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb)
        """,
        (
            import_batch_id,
            sheet_run_id,
            source_file_name,
            source_sheet_name,
            target_table,
            source_row_number,
            error_stage,
            error_type,
            error_message,
            json.dumps(error_details or {}),
        ),
    )


def record_validation_result(
    cur,
    import_batch_id,
    sheet_run_id,
    source_sheet_name,
    validation_stage,
    validation_name,
    status,
    expected_value=None,
    actual_value=None,
    details: Optional[Dict[str, Any]] = None,
    severity="error",
):
    """
    Persist a validation check result, whether passed, warning, or failed.
    """
    cur.execute(
        """
        INSERT INTO ingest.import_validation_results
            (import_batch_id, sheet_run_id, source_sheet_name,
             validation_stage, validation_name, status, severity,
             expected_value, actual_value, details)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb)
        """,
        (
            import_batch_id,
            sheet_run_id,
            source_sheet_name,
            validation_stage,
            validation_name,
            status,
            severity,
            str(expected_value) if expected_value is not None else None,
            str(actual_value) if actual_value is not None else None,
            json.dumps(details or {}),
        ),
    )