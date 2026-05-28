# etl/load_raw.py
"""Load Raw layer: bulk insert extracted rows into raw.* tables with batch tracking."""

import hashlib
import os

from psycopg2.extras import execute_values


def _file_checksum(filepath):
    h = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def file_checksum(filepath: str) -> str:
    """
    Public checksum helper used for idempotency checks.
    """
    return _file_checksum(filepath)


def find_completed_batch(cur, filepath: str, sheet_name: str):
    """
    Return the most recent completed batch id (success/partial_success) for the same
    file checksum + sheet. Returns None if not found.
    """
    checksum = _file_checksum(filepath)
    cur.execute(
        """
        SELECT id
        FROM ingest.import_batches
        WHERE checksum = %s
          AND source_sheet_name = %s
          AND status IN ('success', 'partial_success')
        ORDER BY updated_at DESC NULLS LAST
        LIMIT 1
        """,
        (checksum, sheet_name),
    )
    row = cur.fetchone()
    return row[0] if row else None


def get_batch_counts(cur, batch_id):
    """
    Fetch (rows_detected, rows_loaded, rows_failed, status) for an import batch.
    """
    cur.execute(
        """
        SELECT rows_detected, rows_loaded, rows_failed, status
        FROM ingest.import_batches
        WHERE id = %s
        """,
        (batch_id,),
    )
    row = cur.fetchone()
    if not row:
        return None
    return {
        "rows_detected": row[0] or 0,
        "rows_loaded": row[1] or 0,
        "rows_failed": row[2] or 0,
        "status": row[3],
    }


def create_import_batch(cur, filepath, sheet_name):
    """Insert a record into ingest.import_batches and return its UUID."""
    cur.execute(
        """
        INSERT INTO ingest.import_batches
            (source_file_name, source_sheet_name, source_system, template_type,
             uploaded_by, status, checksum)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        RETURNING id
        """,
        (
            os.path.basename(filepath),
            sheet_name,
            "excel_manual",
            sheet_name,
            "etl_pipeline",
            "processing",
            _file_checksum(filepath),
        ),
    )
    return cur.fetchone()[0]


def update_batch_status(cur, batch_id, rows_detected, rows_loaded, rows_failed=0):
    status = "success" if rows_failed == 0 else "partial_success"
    if rows_loaded == 0:
        status = "failed"

    cur.execute(
        """
        UPDATE ingest.import_batches
        SET rows_detected = %s,
            rows_loaded = %s,
            rows_failed = %s,
            status = %s,
            updated_at = now()
        WHERE id = %s
        """,
        (rows_detected, rows_loaded, rows_failed, status, batch_id),
    )


def fail_import_batch(cur, batch_id, error_message=None):
    """
    Mark an import batch as failed.
    """
    cur.execute(
        """
        UPDATE ingest.import_batches
        SET status = 'failed',
            rows_detected = COALESCE(rows_detected, 0),
            rows_loaded = COALESCE(rows_loaded, 0),
            rows_failed = COALESCE(rows_failed, 0),
            notes = %s,
            updated_at = now()
        WHERE id = %s
        """,
        (error_message, batch_id),
    )


def bulk_insert(cur, conn, table_name, rows, batch_id, batch_size=5000):
    """
    Insert rows into the given raw table with **chunked commits**, so a
    connection drop loses at most one chunk's work instead of the entire
    upload. Returns (detected, loaded, failed).

    Why this matters: the previous design wrapped the full upload (raw
    load + dimensions + facts) in one giant Postgres transaction. On
    Streamlit Cloud + Supabase, the underlying connection has lifetime
    limits we don't control; for large files (200k+ rows over ~2 hours)
    the connection would drop mid-tx and Postgres would automatically
    roll back everything. This refactor commits each chunk independently
    so already-inserted rows are durable even if a later chunk fails.

    Implementation: temporarily set conn.autocommit = True for the
    duration of the load — every execute_values call then auto-commits
    as its own transaction. The previous autocommit value is restored on
    exit so the caller's transactional contract is unaffected.
    """
    buffered = []
    columns = None
    total_detected = 0
    total_loaded = 0
    total_failed = 0

    prior_autocommit = conn.autocommit
    conn.autocommit = True
    try:
        for row in rows:
            total_detected += 1
            row["import_batch_id"] = str(batch_id)
            if columns is None:
                columns = list(row.keys())

            values = tuple(row[c] for c in columns)
            buffered.append(values)

            if len(buffered) >= batch_size:
                loaded, failed = _flush(cur, table_name, columns, buffered)
                total_loaded += loaded
                total_failed += failed
                buffered = []

        if buffered:
            loaded, failed = _flush(cur, table_name, columns, buffered)
            total_loaded += loaded
            total_failed += failed
    finally:
        conn.autocommit = prior_autocommit

    return total_detected, total_loaded, total_failed


def count_rows_for_batch(cur, table_name: str, batch_id) -> int:
    """
    Count rows physically present in a raw table for a given import_batch_id.
    """
    cur.execute(
        f"SELECT COUNT(*) FROM {table_name} WHERE import_batch_id = %s",
        (str(batch_id),),
    )
    return int(cur.fetchone()[0])


def _flush(cur, table_name, columns, rows):
    """Execute one batch INSERT. Returns (loaded, failed).

    Called with the connection in autocommit mode (bulk_insert sets that),
    so each successful execute_values is its own committed transaction and
    a single bad batch can't poison subsequent ones — no savepoint needed.
    """
    cols_sql = ", ".join(columns)
    template = "(" + ", ".join(["%s"] * len(columns)) + ")"
    try:
        execute_values(
            cur,
            f"INSERT INTO {table_name} ({cols_sql}) VALUES %s",
            rows,
            template=template,
            page_size=len(rows),
        )
        return len(rows), 0
    except Exception as e:
        print(f"    ERROR inserting batch into {table_name}: {e}")
        return 0, len(rows)