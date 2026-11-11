"""
Automated regression and unit test suite for Loblaws AR Reconciliation Engine.
"""

import sys
import zipfile
from decimal import Decimal
from pathlib import Path

# Add src to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from loblaws_automation.classifier import classify_invoice_region, identify_claim_category
from loblaws_automation.fbdi_builder import build_fbdi_row, create_fbdi_zip_package
from loblaws_automation.ledger import ReconciliationLedger, compute_file_hash
from loblaws_automation.matcher import LoblawsUpstreamMatcher, MatchingSummary
from loblaws_automation.models import (
    ReconciliationBatch,
    SkippedRow,
    TransactionItem,
)
from loblaws_automation.reconciler import (
    LoblawsReconciler,
    _excel_serial_to_iso,
    _extract_date_from_filename,
)


def test_regional_classification():
    """Verify exact regional routing rules."""
    # West: Length 12, starts with 2
    div, acct = classify_invoice_region("226303821637")
    assert div == "West"
    assert acct == "60463886"

    # Ontario: Length 11
    div, acct = classify_invoice_region("25690727525")
    assert div == "Ontario"
    assert acct == "60477290"

    # Atlantic: Length 12, starts with 1
    div, acct = classify_invoice_region("126127021336")
    assert div == "Atlantic"
    assert acct == "60455049"

    # Quebec: Length 12, starts with 3
    div, acct = classify_invoice_region("326686121507")
    assert div == "Quebec"
    assert acct == "60503134"


def test_claim_category_taxonomy():
    """Verify claim prefix mapping."""
    assert "Operations" in identify_claim_category("OP-25690727525")
    assert "Not Found" in identify_claim_category("NF-6554621612")
    assert "Associate Credit" in identify_claim_category("INASCR-226407321290")
    assert "Associate Debit" in identify_claim_category("INASDR-226530721532")
    assert "Master" in identify_claim_category("NAC2020094821-A")


def test_zero_sum_balancing_invariant():
    """Verify that imbalanced batches are strictly rejected."""
    balanced_batch = ReconciliationBatch(
        cheque_number="TEST001",
        cheque_date="2026-09-08",
        transactions=[
            TransactionItem(account_number="60455049", amount=Decimal("-100.00"), trx_number="NAC-TEST"),
            TransactionItem(account_number="60477290", amount=Decimal("60.00"), trx_number="OP-01"),
            TransactionItem(account_number="60463886", amount=Decimal("40.00"), trx_number="NF-01"),
        ],
    )
    assert balanced_batch.is_balanced is True
    assert balanced_batch.net_sum == Decimal("0.00")

    imbalanced_batch = ReconciliationBatch(
        cheque_number="TEST002",
        cheque_date="2026-09-08",
        transactions=[
            TransactionItem(account_number="60455049", amount=Decimal("-100.00"), trx_number="NAC-TEST"),
            TransactionItem(account_number="60477290", amount=Decimal("50.00"), trx_number="OP-01"),
        ],
    )
    assert imbalanced_batch.is_balanced is False


def test_fbdi_row_builder():
    """Verify exact 19-column mapping for Oracle FBDI."""
    item = TransactionItem(
        account_number="60455049",
        amount=Decimal("-460015.19"),
        trx_number="NAC2020094821-A",
    )
    row = build_fbdi_row(item, "2026-09-08", "2026-09-08")
    assert len(row) == 300
    assert row[1] == "CACBC"                  # BU
    assert row[2] == "CACBC_IMPORTED"         # Source
    assert row[3] == "CACBC DEDUC-MGMT-CR"    # Type
    assert row[4] == "Net 7 Days"             # Terms
    assert row[5] == "2026-09-08"             # Trx Date
    assert row[6] == "2026-09-08"             # GL Date
    assert row[7] == "NAC2020094821-A"        # Trx Number
    assert row[19] == "60455049"              # Customer Acct
    assert row[20] == "60455049"              # Customer Site
    assert row[26] == "LINE"                  # Line Type
    assert row[27] == "NAC2020094821-A"        # Description
    assert row[28] == "CAD"                   # Currency
    assert row[29] == "User"                  # Conversion Type
    assert row[31] == "1"                     # Conversion Rate
    assert row[32] == "-460015.19"            # Amount
    assert row[33] == "1"                     # Quantity
    assert row[37] == "CB_IMPORTED"           # Flexfield Context
    assert row[38] == "NAC2020094821-A"        # Flexfield Segment 1
    assert row[90] == "N"                     # Taxable Flag


def test_live_historical_batch():
    """End-to-end integration test against LOBLAWS_NAC2020094821_090826.xlsx."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    assert zip_path.exists()

    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2020094821_090826.xlsx"
        assert target_name in z.namelist()
        file_bytes = z.read(target_name)

    reconciler = LoblawsReconciler(strict_balancing=True)
    batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name="TestBatch")

    assert batch.cheque_number == "2020094821"
    assert batch.total_count == 464
    assert batch.is_balanced is True
    assert abs(metrics.net_variance) < Decimal("0.01")
    assert metrics.execution_time_ms < 1000.0  # Must execute in under 1 second!

    zip_bytes, zip_name = create_fbdi_zip_package(batch)
    assert len(zip_bytes) > 0
    assert zip_name == "AutoInvoiceImport_2020094821.zip"


def test_all_13_historical_batches():
    """Verify that all 13 historical workbooks balance to $0.00 with zero variance."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    assert zip_path.exists()

    reconciler = LoblawsReconciler(strict_balancing=True)
    with zipfile.ZipFile(zip_path, mode="r") as z:
        excel_files = [n for n in z.namelist() if n.endswith(".xlsx")]
        assert len(excel_files) == 13

        total_lines = 0
        total_debits = Decimal("0.00")
        for fpath in excel_files:
            file_bytes = z.read(fpath)
            batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name=Path(fpath).name)
            assert batch.is_balanced is True
            assert abs(metrics.net_variance) < Decimal("0.005")
            total_lines += batch.total_count
            total_debits += metrics.total_debits

        assert total_lines == 2800
        assert round(total_debits, 2) == Decimal("8352871.39")


# =============================================================================
# NEW TESTS: Bug fix verification
# =============================================================================

def test_excel_serial_date_conversion():
    """BUG FIX #1: Verify Excel serial date → ISO date conversion."""
    # 2026-09-08 = Excel serial 46348 (approx)
    # Known reference: 2024-01-01 = 45292
    result = _excel_serial_to_iso("45292")
    assert result == "2024-01-01"

    # Edge cases
    assert _excel_serial_to_iso("1") == "1899-12-31"
    assert _excel_serial_to_iso("") is None
    assert _excel_serial_to_iso("abc") is None
    assert _excel_serial_to_iso("-5") is None
    assert _excel_serial_to_iso("999999") is None  # Out of range


def test_filename_date_extraction():
    """BUG FIX #1: Verify date extraction from workbook filenames."""
    # Standard format: LOBLAWS_NAC<cheque>_MMDDYY.xlsx
    result = _extract_date_from_filename("LOBLAWS_NAC2020094821_090826.xlsx")
    assert result == "2026-09-08"

    result = _extract_date_from_filename("LOBLAWS_NAC123_010125.xlsx")
    assert result == "2025-01-01"

    # No match
    assert _extract_date_from_filename("random_file.xlsx") is None
    assert _extract_date_from_filename("file.csv") is None


def test_skipped_row_model():
    """BUG FIX #2: Verify SkippedRow model works correctly."""
    sr = SkippedRow(
        row_index=42,
        raw_values={"A": "60455049", "B": "not_a_number", "C": "OP-123"},
        reason="Failed to parse amount 'not_a_number'",
    )
    assert sr.row_index == 42
    assert "not_a_number" in sr.reason
    assert sr.raw_values["A"] == "60455049"


def test_live_batch_has_no_skipped_rows():
    """BUG FIX #2: Verify that the production batch 2020094821 skips zero rows."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    assert zip_path.exists()

    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2020094821_090826.xlsx"
        file_bytes = z.read(target_name)

    reconciler = LoblawsReconciler(strict_balancing=True)
    batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name="LOBLAWS_NAC2020094821_090826.xlsx")

    assert len(metrics.skipped_rows) == 0, (
        f"Expected 0 skipped rows but got {len(metrics.skipped_rows)}: "
        f"{[(sr.row_index, sr.reason) for sr in metrics.skipped_rows]}"
    )
    assert metrics.total_source_rows > 0
    assert batch.total_count == 464


def test_live_batch_cheque_date_is_not_hardcoded():
    """BUG FIX #1: Verify the cheque date is dynamically resolved, not hardcoded."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    assert zip_path.exists()

    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2020094821_090826.xlsx"
        file_bytes = z.read(target_name)

    reconciler = LoblawsReconciler(strict_balancing=True)
    batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name="LOBLAWS_NAC2020094821_090826.xlsx")

    # The date should come from the workbook or filename, not be a hardcoded constant
    assert metrics.cheque_date_source != "unknown", "Cheque date source should be resolved"
    # The filename contains 090826 → 2026-09-08, so at minimum the filename fallback should work
    assert batch.cheque_date != "", "Cheque date must not be empty"


def test_all_13_batches_have_zero_skipped_rows():
    """BUG FIX #2: Verify no rows are silently lost across all 13 historical batches."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    assert zip_path.exists()

    reconciler = LoblawsReconciler(strict_balancing=True)
    with zipfile.ZipFile(zip_path, mode="r") as z:
        excel_files = [n for n in z.namelist() if n.endswith(".xlsx")]

        total_skipped = 0
        for fpath in excel_files:
            file_bytes = z.read(fpath)
            _batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name=Path(fpath).name)
            total_skipped += len(metrics.skipped_rows)
            assert len(metrics.skipped_rows) == 0, (
                f"{fpath}: {len(metrics.skipped_rows)} rows skipped — "
                f"{[(sr.row_index, sr.reason) for sr in metrics.skipped_rows]}"
            )

        assert total_skipped == 0


def test_reconciliation_ledger(tmp_path):
    """Verify persistent SQLite reconciliation ledger for idempotency and history."""
    db_file = tmp_path / "test_ledger.db"
    ledger = ReconciliationLedger(db_file)

    # Initially empty
    assert ledger.is_batch_processed("TEST_CHQ_001") is None
    assert len(ledger.get_all_batches()) == 0

    # Build mock batch and metrics
    batch = ReconciliationBatch(
        cheque_number="TEST_CHQ_001",
        cheque_date="2026-09-08",
        transactions=[
            TransactionItem(account_number="60455049", amount=Decimal("-50.00"), trx_number="NAC-01"),
            TransactionItem(account_number="60455049", amount=Decimal("50.00"), trx_number="OP-01"),
        ],
    )
    from loblaws_automation.models import AuditMetrics
    metrics = AuditMetrics(
        batch_id="BATCH-TEST_CHQ_001",
        cheque_number="TEST_CHQ_001",
        cheque_date="2026-09-08",
        total_lines=2,
        net_variance=Decimal("0.00"),
        is_zero_sum_verified=True,
        total_debits=Decimal("50.00"),
        total_credits=Decimal("50.00"),
        execution_time_ms=12.5,
        total_source_rows=2,
        cheque_date_source="test",
    )

    dummy_hash = compute_file_hash(b"test file content bytes")

    # Record batch
    ledger.record_batch(
        batch=batch,
        metrics=metrics,
        file_hash=dummy_hash,
        source_filename="test_file.xlsx",
        fbdi_zip_path="AutoInvoiceImport_TEST_CHQ_001.zip",
        audit_report_path="Reconciliation_Audit_TEST_CHQ_001.xlsx",
    )

    # Verify idempotency check
    record = ledger.is_batch_processed("TEST_CHQ_001")
    assert record is not None
    assert record["cheque_number"] == "TEST_CHQ_001"
    assert record["file_hash"] == dummy_hash
    assert record["total_lines"] == 2
    assert record["is_balanced"] == 1

    # Verify history listing
    history = ledger.get_all_batches()
    assert len(history) == 1
    assert history[0]["batch_id"] == "BATCH-TEST_CHQ_001"


def test_account_mismatch_cross_validation():
    """Verify that routing discrepancies between invoice region and account are caught."""
    # Batch 022026 is known to have 6 account mismatches in the historical data
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2018845703_022026.xlsx"
        file_bytes = z.read(target_name)

    reconciler = LoblawsReconciler(strict_balancing=True)
    _batch, metrics = reconciler.process_workbook_bytes(file_bytes, source_name="022026.xlsx")

    assert len(metrics.account_mismatches) == 6
    first_mm = metrics.account_mismatches[0]
    assert first_mm.account_number == "60455049"  # Staged as Atlantic
    assert first_mm.inferred_division == "Quebec"  # But starts with 3, length 12
    assert first_mm.inferred_account == "60503134"


def test_missing_txn_sheet_raises():
    """Verify that a workbook missing the required TXN sheet raises an explicit error."""
    import io
    dummy_zip_buffer = io.BytesIO()
    with zipfile.ZipFile(dummy_zip_buffer, mode="w") as z:
        # Minimal empty workbook with only sheet1
        z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="OTHER" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')

    reconciler = LoblawsReconciler()
    import pytest
    with pytest.raises(ValueError, match=r"\[TXN\] not found"):
        reconciler.process_workbook_bytes(dummy_zip_buffer.getvalue(), source_name="bad_file.xlsx")


def test_upstream_matcher_polars_join():
    """Verify that the Polars hash join matches records, categorizes variances, and balances to zero."""
    import polars as pl

    from loblaws_automation.config import CUSTOMER_ACCOUNT_MAP

    matcher = LoblawsUpstreamMatcher(write_off_threshold=11.0)

    # Mock Remittance DataFrame
    # 1. Clean match ($0.00 variance)
    # 2. Small balance write-off (-$5.00 variance -> WO-)
    # 3. Operations claim (-$50.00 variance -> OP-)
    # 4. Pending claim (Not found -> NF-)
    df_remit = pl.DataFrame({
        "invoice_number": ["126127021336", "226303821637", "25690727525", "326686121507"],
        "invoice_date": ["2026-09-01", "2026-09-02", "2026-09-03", "2026-09-04"],
        "gross_payment": [100.0, 95.0, 50.0, 75.0],
        "net_payment": [100.0, 95.0, 50.0, 75.0],
        "store_number": ["001", "002", "003", "004"],
    }).with_columns(
        pl.col("invoice_number").str.slice(-10).alias("match_key")
    )

    # Mock Open AR DataFrame (sheet VIEW)
    df_open_ar = pl.DataFrame({
        "trx_number": ["126127021336", "226303821637", "25690727525"],
        "trx_date": ["2026-08-15", "2026-08-16", "2026-08-17"],
        "due_remaining": [100.0, 100.0, 100.0],
        "bill_to_cust": ["60455049", "60463886", "60477290"],
    }).with_columns(
        pl.col("trx_number").str.slice(-10).alias("match_key")
    )

    meta = {
        "cheque_number": "TEST_UPSTREAM_01",
        "cheque_date": "2026-09-08",
        "total_amount": Decimal("320.00"),
    }
    special_adjs = [
        {"C": "126127021336", "D": "15.00"}
    ]

    batch, _metrics, summary = matcher.match_and_reconcile(
        df_remittance=df_remit,
        df_open_ar=df_open_ar,
        meta=meta,
        special_adjustments=special_adjs,
    )

    assert isinstance(summary, MatchingSummary)
    assert summary.total_remittance_rows == 4
    assert summary.open_ar_ledger_rows == 3
    assert summary.matched_clean_rows == 1
    assert summary.write_off_rows == 1
    assert summary.matched_variance_rows == 1
    assert summary.not_found_rows == 1
    assert summary.adjustment_rows == 1

    # Invariant assertions
    assert batch.is_balanced is True
    assert batch.net_sum == Decimal("0.00")
    assert batch.total_count == 5  # 1 master + 1 WO + 1 OP + 1 NF + 1 adj

    # Master line is first and exactly offsets the sum of all deduction items
    master = batch.transactions[0]
    assert master.trx_number == "NACTEST_UPSTREAM_01-A"
    assert master.amount == -sum((t.amount for t in batch.transactions[1:]), Decimal("0.00"))
    assert master.division == "Atlantic"

    # Verify write-off route to Discounts account
    wo_tx = next(t for t in batch.transactions if t.trx_number.startswith("WO-"))
    assert wo_tx.account_number == CUSTOMER_ACCOUNT_MAP["Discounts"]
    assert wo_tx.amount == Decimal("5.00")

    # Verify operations claim route to West account
    op_tx = next(t for t in batch.transactions if t.trx_number.startswith("OP-"))
    assert op_tx.account_number == "60477290"
    assert op_tx.amount == Decimal("50.00")

    # Verify pending claim route to Quebec account
    nf_tx = next(t for t in batch.transactions if t.trx_number.startswith("NF-"))
    assert nf_tx.account_number == "60503134"
    assert nf_tx.amount == Decimal("-75.00")


def test_upstream_reconcile_raw_workbook():
    """Verify end-to-end upstream matching directly against reference workbook."""
    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2020094821_090826.xlsx"
        file_bytes = z.read(target_name)

    matcher = LoblawsUpstreamMatcher()
    batch, metrics, summary = matcher.reconcile_from_raw_workbook(
        file_bytes, source_name="LOBLAWS_NAC2020094821_090826.xlsx"
    )

    assert batch.cheque_number == "2020094821"
    assert batch.cheque_date == "2026-09-08"
    assert batch.is_balanced is True
    assert batch.net_sum == Decimal("0.00")
    assert metrics.total_debits == metrics.total_credits
    assert metrics.total_debits == Decimal("470795.53")

    # Latency check: vectorized hash join must complete in under 500 ms (typically < 10 ms)
    assert summary.hash_join_duration_ms < 500.0
    assert summary.matched_clean_rows == 1104
    assert summary.open_ar_ledger_rows == 33622
    assert summary.total_remittance_rows == 1617
    assert summary.write_off_rows == 18
    assert summary.matched_variance_rows == 33
    assert summary.not_found_rows == 462


def test_reconciler_modes_and_auto_routing():
    """Verify reconciler mode handling (auto, staged, upstream) and auto-fallback routing."""
    import io

    import pytest

    zip_path = Path(r"C:\Antigravity\Loblaws\Sample Files after Manual steps.zip")
    with zipfile.ZipFile(zip_path, mode="r") as z:
        target_name = "Sample Files after Manual steps/LOBLAWS_NAC2020094821_090826.xlsx"
        file_bytes = z.read(target_name)

    reconciler = LoblawsReconciler(strict_balancing=True)

    # 1. Mode: staged
    batch_staged, _metrics_staged = reconciler.process_workbook_bytes(
        file_bytes, source_name="test_staged.xlsx", mode="staged"
    )
    assert batch_staged.total_count == 464
    assert batch_staged.is_balanced is True

    # 2. Mode: upstream
    batch_up, metrics_up = reconciler.process_workbook_bytes(
        file_bytes, source_name="test_up.xlsx", mode="upstream"
    )
    assert batch_up.total_count == 514
    assert batch_up.is_balanced is True
    assert reconciler.last_matching_summary is not None
    assert metrics_up.matching_summary is not None

    # 3. Mode: auto on file with TXN (defaults to staged TXN)
    batch_auto, _ = reconciler.process_workbook_bytes(
        file_bytes, source_name="test_auto.xlsx", mode="auto"
    )
    assert batch_auto.total_count == 464

    # 4. Mode: auto on raw workbook WITHOUT TXN (auto-routes to upstream)
    # Build a workbook having only N and VIEW (TXN stripped from workbook.xml)
    import re
    raw_buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(file_bytes), mode="r") as z_in:
        with zipfile.ZipFile(raw_buffer, mode="w") as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    # Remove TXN sheet declaration
                    data_str = data.decode("utf-8")
                    data_str = re.sub(r'<sheet [^>]*name="TXN"[^>]*/>', "", data_str)
                    data = data_str.encode("utf-8")
                z_out.writestr(item, data)

    batch_routed, metrics_routed = reconciler.process_workbook_bytes(
        raw_buffer.getvalue(), source_name="raw_only.xlsx", mode="auto"
    )
    assert batch_routed.total_count == 514
    assert batch_routed.is_balanced is True
    assert metrics_routed.matching_summary is not None

    # 5. Invalid mode raises ValueError
    with pytest.raises(ValueError, match="Invalid mode"):
        reconciler.process_workbook_bytes(file_bytes, mode="invalid_mode")


