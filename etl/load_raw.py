"""Load Raw layer: bulk insert extracted rows into raw.* tables with batch tracking."""

import hashlib
import os
from datetime import datetime

from psycopg2.extras import execute_values


def _file_checksum(filepath):
    h = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


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
        SET rows_detected = %s, rows_loaded = %s, rows_failed = %s,
            status = %s, updated_at = now()
        WHERE id = %s
        """,
        (rows_detected, rows_loaded, rows_failed, status, batch_id),
    )


def bulk_insert(cur, table_name, rows, batch_id, batch_size=1000):
    """Insert rows into the given raw table. Returns (loaded, failed) counts."""
    buffered = []
    columns = None
    total_loaded = 0
    total_failed = 0

    for row in rows:
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

    return total_loaded, total_failed


def _flush(cur, table_name, columns, rows):
    """Execute a batch INSERT. Returns (loaded, failed)."""
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
