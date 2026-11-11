"""
Persistent SQLite-based batch ledger for idempotency, deduplication, and historical audit tracking.
"""

import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from loblaws_automation.models import AuditMetrics, ReconciliationBatch

DEFAULT_LEDGER_PATH = Path("output") / "reconciliation_ledger.db"


def compute_file_hash(file_bytes: bytes) -> str:
    """Computes SHA-256 checksum of raw file bytes."""
    return hashlib.sha256(file_bytes).hexdigest()


class ReconciliationLedger:
    """Manages the persistent audit ledger of processed batches."""

    def __init__(self, db_path: Path | None = None):
        self.db_path = Path(db_path or DEFAULT_LEDGER_PATH)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        """Initializes ledger tables if they do not exist."""
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS processed_batches (
                    cheque_number TEXT PRIMARY KEY,
                    batch_id TEXT NOT NULL,
                    cheque_date TEXT NOT NULL,
                    cheque_date_source TEXT NOT NULL,
                    file_hash TEXT NOT NULL,
                    source_filename TEXT NOT NULL,
                    total_lines INTEGER NOT NULL,
                    total_source_rows INTEGER NOT NULL,
                    skipped_rows_count INTEGER NOT NULL,
                    account_mismatches_count INTEGER NOT NULL DEFAULT 0,
                    net_variance REAL NOT NULL,
                    total_debits REAL NOT NULL,
                    total_credits REAL NOT NULL,
                    is_balanced INTEGER NOT NULL,
                    fbdi_zip_path TEXT,
                    audit_report_path TEXT,
                    execution_time_ms REAL NOT NULL,
                    processed_at TEXT NOT NULL
                )
            """)
            conn.commit()

    def is_batch_processed(self, cheque_number: str) -> dict[str, Any] | None:
        """
        Checks if a cheque number has already been processed.
        Returns the existing record dict if found, or None.
        """
        with self._get_connection() as conn:
            cur = conn.execute(
                "SELECT * FROM processed_batches WHERE cheque_number = ?",
                (cheque_number,),
            )
            row = cur.fetchone()
            return dict(row) if row else None

    def record_batch(
        self,
        batch: ReconciliationBatch,
        metrics: AuditMetrics,
        file_hash: str,
        source_filename: str,
        fbdi_zip_path: str | None = None,
        audit_report_path: str | None = None,
    ) -> None:
        """
        Records or updates a processed batch in the ledger.
        """
        now_iso = datetime.now().isoformat()
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO processed_batches (
                    cheque_number,
                    batch_id,
                    cheque_date,
                    cheque_date_source,
                    file_hash,
                    source_filename,
                    total_lines,
                    total_source_rows,
                    skipped_rows_count,
                    account_mismatches_count,
                    net_variance,
                    total_debits,
                    total_credits,
                    is_balanced,
                    fbdi_zip_path,
                    audit_report_path,
                    execution_time_ms,
                    processed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                batch.cheque_number,
                metrics.batch_id,
                batch.cheque_date,
                metrics.cheque_date_source,
                file_hash,
                source_filename,
                batch.total_count,
                metrics.total_source_rows,
                len(metrics.skipped_rows),
                len(metrics.account_mismatches),
                float(metrics.net_variance),
                float(metrics.total_debits),
                float(metrics.total_credits),
                1 if metrics.is_zero_sum_verified else 0,
                fbdi_zip_path,
                audit_report_path,
                metrics.execution_time_ms,
                now_iso,
            ))
            conn.commit()

    def get_all_batches(self) -> list[dict[str, Any]]:
        """Returns all recorded batch records ordered by processed_at descending."""
        with self._get_connection() as conn:
            cur = conn.execute(
                "SELECT * FROM processed_batches ORDER BY processed_at DESC"
            )
            return [dict(row) for row in cur.fetchall()]
