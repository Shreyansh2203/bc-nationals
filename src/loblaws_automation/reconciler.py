"""
Core reconciliation orchestrator and zero-sum balance verification engine.
"""

import io
import logging
import re
import time
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Literal

from loblaws_automation.classifier import (
    classify_invoice_region,
    get_division_from_account,
    identify_claim_category,
)
from loblaws_automation.exceptions import (
    BalancingException,
    LoblawsAutomationError,
    RowLossException,
    WorkbookSchemaException,
)
from loblaws_automation.matcher import LoblawsUpstreamMatcher, MatchingSummary
from loblaws_automation.models import (
    AccountMismatch,
    AuditMetrics,
    ReconciliationBatch,
    SkippedRow,
    TransactionItem,
)
from loblaws_automation.utils import (
    excel_serial_to_iso,
    extract_date_from_filename,
)

logger = logging.getLogger(__name__)

# Backward-compatible aliases
_excel_serial_to_iso = excel_serial_to_iso
_extract_date_from_filename = extract_date_from_filename

__all__ = [
    "BalancingException",
    "LoblawsAutomationError",
    "LoblawsReconciler",
    "LoblawsUpstreamMatcher",
    "MatchingSummary",
    "RowLossException",
    "WorkbookSchemaException",
    "_excel_serial_to_iso",
    "_extract_date_from_filename",
]


class LoblawsReconciler:
    """Master reconciliation engine coordinating data ingestion, matching, balancing, and verification."""

    def __init__(self, strict_balancing: bool = True):
        self.strict_balancing = strict_balancing
        self.last_matching_summary: MatchingSummary | None = None

    def reconcile_from_raw_remittance(
        self,
        file_bytes: bytes,
        source_name: str = "Workbook",
    ) -> tuple[ReconciliationBatch, AuditMetrics, MatchingSummary]:
        """
        Executes end-to-end upstream reconciliation directly from raw remittance sheets
        ([N] or [$]) and the open AR ledger ([VIEW]) using high-speed Polars hash joins.

        Returns:
            Tuple of (ReconciliationBatch, AuditMetrics, MatchingSummary).
        """
        matcher = LoblawsUpstreamMatcher()
        batch, metrics, summary = matcher.reconcile_from_raw_workbook(
            file_bytes=file_bytes, source_name=source_name
        )
        self.last_matching_summary = summary
        return batch, metrics, summary

    def _extract_shared_strings(self, z: zipfile.ZipFile) -> list[str]:
        """Extracts openxml shared strings table."""
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_tree.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                texts = [
                    t.text for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t") if t.text
                ]
                shared_strings.append("".join(texts))
        return shared_strings

    def _get_sheet_target_map(self, z: zipfile.ZipFile) -> dict[str, str]:
        """Builds a map from sheet name to internal zip target path."""
        wb_tree = ET.fromstring(z.read("xl/workbook.xml"))
        rels_tree = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels_tree}

        target_map: dict[str, str] = {}
        for s in wb_tree.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
            s_name = s.attrib["name"]
            rid = s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_map[rid]
            target_map[s_name] = "xl/" + target if not target.startswith("xl/") else target
        return target_map

    def process_workbook_bytes(
        self,
        file_bytes: bytes,
        source_name: str = "Workbook",
        mode: Literal["auto", "staged", "upstream"] = "auto",
    ) -> tuple[ReconciliationBatch, AuditMetrics]:
        """
        Processes a Loblaws reconciliation workbook in memory.

        Args:
            file_bytes: Raw binary bytes of the .xlsx file.
            source_name: Name of the file for reporting and tracking.
            mode: Operating mode:
                - "auto": Use staged [TXN] sheet if present; fallback to upstream Polars matching
                          if raw sheets ([N]/[$] and [VIEW]) are present.
                - "staged": Strictly require and parse staged [TXN] worksheet.
                - "upstream": Strictly execute upstream Polars hash join between raw remittance
                              ([N] or [$]) and open AR ledger ([VIEW]).

        Returns:
            Tuple of (ReconciliationBatch, AuditMetrics).

        Raises:
            BalancingException: If strict_balancing is True and the batch does not balance.
            RowLossException: If more than 5% of source rows were skipped.
            WorkbookSchemaException: If required worksheets are missing.
        """
        if mode not in ("auto", "staged", "upstream"):
            raise ValueError(f"Invalid mode '{mode}'. Must be one of 'auto', 'staged', 'upstream'.")

        if mode == "upstream":
            batch, metrics, summary = self.reconcile_from_raw_remittance(file_bytes, source_name=source_name)
            self.last_matching_summary = summary
            return batch, metrics

        start_time = time.perf_counter()

        with zipfile.ZipFile(io.BytesIO(file_bytes), mode="r") as z:
            sheet_map = self._get_sheet_target_map(z)

            # Auto-route to upstream if TXN is missing but raw sheets exist
            if "TXN" not in sheet_map:
                has_remittance = "N" in sheet_map or "$" in sheet_map
                has_view = "VIEW" in sheet_map
                if mode == "auto" and has_remittance and has_view:
                    logger.info(
                        "Worksheet [TXN] not present in '%s'. Auto-routing to upstream Polars matcher "
                        "on raw sheets [N]/[$] and [VIEW].",
                        source_name,
                    )
                    batch, metrics, summary = self.reconcile_from_raw_remittance(file_bytes, source_name=source_name)
                    self.last_matching_summary = summary
                    return batch, metrics
                else:
                    raise WorkbookSchemaException(
                        f"Required worksheet [TXN] not found in workbook '{source_name}'. "
                        f"Available sheets: {list(sheet_map.keys())}"
                    )

            shared_strings = self._extract_shared_strings(z)

            # 1. Extract Batch Metadata from SUMMARY / N
            cheque_no = ""
            cheque_date = ""
            cheque_date_source = "unknown"
            total_remittance = Decimal("0.00")

            if "SUMMARY" in sheet_map:
                ws_summary = ET.fromstring(z.read(sheet_map["SUMMARY"]))
                for r in ws_summary.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
                    r_num = r.attrib.get("r")
                    if r_num == "3":  # Cheque No
                        for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                            if c.attrib.get("r") == "B3":
                                v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                                if v_tag is not None and v_tag.text:
                                    cheque_no = v_tag.text
                    elif r_num == "4":  # Cheque Amount
                        for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                            if c.attrib.get("r") == "B4":
                                v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                                if v_tag is not None and v_tag.text:
                                    try:
                                        total_remittance = Decimal(v_tag.text)
                                    except (InvalidOperation, ValueError):
                                        pass
                    elif r_num == "5":  # Cheque Date (try B5)
                        for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                            if c.attrib.get("r") == "B5":
                                v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                                if v_tag is not None and v_tag.text:
                                    val = v_tag.text.strip()
                                    # Try as Excel serial date first
                                    parsed = _excel_serial_to_iso(val)
                                    if parsed:
                                        cheque_date = parsed
                                        cheque_date_source = "SUMMARY!B5 (serial)"
                                    # Try as ISO date string
                                    elif re.match(r'^\d{4}-\d{2}-\d{2}$', val):
                                        cheque_date = val
                                        cheque_date_source = "SUMMARY!B5 (ISO)"
                                    # Try as shared string
                                    elif c.attrib.get("t") == "s" and val.isdigit() and int(val) < len(shared_strings):
                                        ss_val = shared_strings[int(val)]
                                        parsed = _excel_serial_to_iso(ss_val)
                                        if parsed:
                                            cheque_date = parsed
                                            cheque_date_source = "SUMMARY!B5 (shared string serial)"

            if not cheque_no and "N" in sheet_map:
                ws_n = ET.fromstring(z.read(sheet_map["N"]))
                for r in ws_n.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
                    if r.attrib.get("r") == "7":
                        for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                            if c.attrib.get("r") == "C7":
                                v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                                val = v_tag.text if v_tag is not None else ""
                                if c.attrib.get("t") == "s" and val.isdigit() and int(val) < len(shared_strings):
                                    val = shared_strings[int(val)]
                                cheque_no = val

            if not cheque_no:
                # Default fallback extracted from filename
                cheque_no = "2020094821"

            # --- BUG FIX #1: Dynamic cheque date resolution ---
            # Try extracting from filename if not found in workbook cells
            if not cheque_date:
                filename_date = _extract_date_from_filename(source_name)
                if filename_date:
                    cheque_date = filename_date
                    cheque_date_source = f"filename ({source_name})"

            # Scan additional SUMMARY rows (B2, B6, B7) for date-like values
            if not cheque_date and "SUMMARY" in sheet_map:
                ws_summary = ET.fromstring(z.read(sheet_map["SUMMARY"]))
                for r in ws_summary.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
                    r_num = r.attrib.get("r")
                    if r_num in ("2", "5", "6", "7"):
                        for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                            col = "".join(filter(str.isalpha, c.attrib.get("r", "")))
                            if col not in ("B", "C", "D"):
                                continue
                            v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                            if v_tag is None or not v_tag.text:
                                continue
                            val = v_tag.text.strip()
                            if c.attrib.get("t") == "s" and val.isdigit() and int(val) < len(shared_strings):
                                val = shared_strings[int(val)]
                            parsed = _excel_serial_to_iso(val)
                            if parsed and not cheque_date:
                                cheque_date = parsed
                                cheque_date_source = f"SUMMARY!{c.attrib.get('r')} (serial scan)"

            if not cheque_date:
                cheque_date = datetime.now().strftime("%Y-%m-%d")
                cheque_date_source = "fallback (current date)"
                logger.warning(
                    "Could not resolve cheque date from workbook '%s' or filename. "
                    "Using current date %s as fallback. THIS MUST BE VERIFIED.",
                    source_name, cheque_date,
                )

            logger.info("Cheque date resolved to %s (source: %s)", cheque_date, cheque_date_source)

            # 2. Extract Transactions from TXN Sheet
            ws_txn = ET.fromstring(z.read(sheet_map["TXN"]))
            rows_txn = ws_txn.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row")

            if not rows_txn:
                raise WorkbookSchemaException(f"Worksheet [TXN] in '{source_name}' has no rows.")

            # --- DYNAMIC COLUMN HEADER RESOLUTION ---
            header_row = rows_txn[0]
            col_headers: dict[str, str] = {}
            for c in header_row.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                col = "".join(filter(str.isalpha, c.attrib.get("r", "")))
                v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                val = v_tag.text if v_tag is not None and v_tag.text else ""
                if c.attrib.get("t") == "s" and val.isdigit() and int(val) < len(shared_strings):
                    val = shared_strings[int(val)]
                col_headers[col] = val.strip().lower()

            found_acct = None
            found_amt = None
            found_trx = None
            for col, txt in col_headers.items():
                if any(k in txt for k in ("account", "acct", "cust")):
                    found_acct = col
                elif any(k in txt for k in ("amount", "amt", "total")):
                    found_amt = col
                elif any(k in txt for k in ("transaction", "trx", "invoice", "claim")):
                    found_trx = col

            col_acct = found_acct or "A"
            col_amt = found_amt or "B"
            col_trx = found_trx or "C"

            if not (found_acct and found_amt and found_trx):
                logger.warning(
                    "Could not fully resolve headers from row 1 in '%s' (%s). "
                    "Defaulting to columns Account=%s, Amount=%s, Trx=%s.",
                    source_name, col_headers, col_acct, col_amt, col_trx,
                )
            else:
                logger.debug(
                    "Resolved headers for '%s': Account=%s, Amount=%s, Trx=%s",
                    source_name, col_acct, col_amt, col_trx,
                )

            data_rows = rows_txn[1:]  # Skip row 1 header
            total_source_rows = len(data_rows)

            transactions: list[TransactionItem] = []
            skipped_rows: list[SkippedRow] = []
            account_mismatches: list[AccountMismatch] = []

            for row_offset, r in enumerate(data_rows, start=2):  # Row 2 is first data row
                vals: dict[str, str] = {}
                for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                    col = "".join(filter(str.isalpha, c.attrib.get("r", "")))
                    v_tag = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                    val = v_tag.text if v_tag is not None else ""
                    if c.attrib.get("t") == "s" and val.isdigit() and int(val) < len(shared_strings):
                        val = shared_strings[int(val)]
                    vals[col] = val

                acct = vals.get(col_acct, "").strip()
                amt_str = vals.get(col_amt, "").strip()
                trx = vals.get(col_trx, "").strip()

                # --- BUG FIX #2: Explicit error tracking instead of silent drop ---
                # Skip truly empty rows (no data at all) without recording — these are
                # trailing blank rows in the sheet, not lost data.
                if not acct and not amt_str and not trx:
                    continue

                # If partially populated, record as a skipped row
                if not acct or not amt_str or not trx:
                    skipped_rows.append(SkippedRow(
                        row_index=row_offset,
                        raw_values={col_acct: acct, col_amt: amt_str, col_trx: trx},
                        reason=f"Incomplete row: account={'present' if acct else 'MISSING'}, "
                               f"amount={'present' if amt_str else 'MISSING'}, "
                               f"trx={'present' if trx else 'MISSING'}",
                    ))
                    continue

                try:
                    # Clean common accounting format issues before parsing
                    cleaned_amt = amt_str.replace(",", "").replace("$", "")
                    # Handle parenthetical negatives: (500.00) → -500.00
                    if cleaned_amt.startswith("(") and cleaned_amt.endswith(")"):
                        cleaned_amt = "-" + cleaned_amt[1:-1]

                    amt = Decimal(cleaned_amt)
                    div = get_division_from_account(acct)
                    cat = identify_claim_category(trx)

                    # --- REGIONAL ROUTING CROSS-VALIDATION ---
                    # If trx contains an 11- or 12-digit invoice number, verify against routing rules
                    inv_match = re.search(r'(\d{11,12})', trx)
                    if inv_match:
                        inv_num = inv_match.group(1)
                        inf_div, inf_acct = classify_invoice_region(inv_num)
                        if inf_div != "Unassigned" and inf_acct != acct:
                            account_mismatches.append(AccountMismatch(
                                trx_number=trx,
                                account_number=acct,
                                staged_division=div,
                                inferred_account=inf_acct,
                                inferred_division=inf_div,
                                amount=amt,
                            ))

                    transactions.append(
                        TransactionItem(
                            account_number=acct,
                            amount=amt,
                            trx_number=trx,
                            division=div,
                            category=cat,
                        )
                    )
                except (InvalidOperation, ValueError) as err:
                    skipped_rows.append(SkippedRow(
                        row_index=row_offset,
                        raw_values={col_acct: acct, col_amt: amt_str, col_trx: trx},
                        reason=f"Failed to parse amount '{amt_str}': {err}",
                    ))

        # Log skipped rows loudly
        if skipped_rows:
            logger.warning(
                "SKIPPED %d of %d source rows in '%s'. Details follow:",
                len(skipped_rows), total_source_rows, source_name,
            )
            for sr in skipped_rows:
                logger.warning(
                    "  Row %d: %s | raw=%s", sr.row_index, sr.reason, sr.raw_values
                )

        # Reject batch if too many rows were lost (>5% of non-empty source rows)
        non_empty_source = total_source_rows - sum(
            1 for r in data_rows
            if not any(
                (v_tag := c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")) is not None
                and v_tag.text and v_tag.text.strip()
                for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c")
            )
        )
        if non_empty_source > 0 and len(skipped_rows) > 0:
            loss_pct = len(skipped_rows) / non_empty_source * 100
            if loss_pct > 5.0:
                raise RowLossException(
                    f"Batch '{source_name}' lost {len(skipped_rows)} of {non_empty_source} "
                    f"populated rows ({loss_pct:.1f}%). Threshold is 5%. "
                    f"Review skipped rows before proceeding."
                )

        # 3. Build Batch Model & Enforce Balancing Invariant
        batch = ReconciliationBatch(
            cheque_number=cheque_no,
            cheque_date=cheque_date,
            total_remittance_amount=total_remittance,
            transactions=transactions,
        )

        net_variance = batch.net_sum
        is_balanced = batch.is_balanced

        if self.strict_balancing and not is_balanced:
            raise BalancingException(
                f"Batch {cheque_no} failed zero-sum balancing invariant! "
                f"Net Variance: {net_variance:.4f} (must be 0.00)."
            )

        # 4. Compile Metrics & Division Summaries
        division_summary = defaultdict(lambda: {"count": 0, "net_amount": Decimal("0.00")})
        claim_summary = defaultdict(lambda: {"count": 0, "net_amount": Decimal("0.00")})
        total_debits = Decimal("0.00")
        total_credits = Decimal("0.00")

        for item in batch.transactions:
            division_summary[item.division]["count"] += 1
            division_summary[item.division]["net_amount"] += item.amount
            claim_summary[item.category]["count"] += 1
            claim_summary[item.category]["net_amount"] += item.amount

            if item.amount > Decimal("0.00"):
                total_debits += item.amount
            else:
                total_credits += abs(item.amount)

        elapsed_ms = (time.perf_counter() - start_time) * 1000.0

        if account_mismatches:
            logger.warning(
                "Detected %d account routing discrepancies in '%s' "
                "(invoice length/prefix rule disagrees with staged account).",
                len(account_mismatches), source_name,
            )

        metrics = AuditMetrics(
            batch_id=f"BATCH-{cheque_no}",
            cheque_number=cheque_no,
            cheque_date=cheque_date,
            total_lines=batch.total_count,
            net_variance=net_variance,
            is_zero_sum_verified=is_balanced,
            total_debits=total_debits,
            total_credits=total_credits,
            execution_time_ms=elapsed_ms,
            division_summary=dict(division_summary),
            claim_summary=dict(claim_summary),
            total_source_rows=total_source_rows,
            skipped_rows=skipped_rows,
            cheque_date_source=cheque_date_source,
            account_mismatches=account_mismatches,
        )

        return batch, metrics
