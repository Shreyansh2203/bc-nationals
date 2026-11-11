"""
Direct Oracle Fusion Cloud AR AutoInvoice CSV and ZIP package builder.
Eliminates all dependencies on desktop Excel, COM objects, and VBA macros.
"""

import csv
import io
import zipfile
from decimal import Decimal

from loblaws_automation.config import (
    DEFAULT_BATCH_SOURCE,
    DEFAULT_BUSINESS_UNIT,
    DEFAULT_CONVERSION_RATE,
    DEFAULT_CONVERSION_TYPE,
    DEFAULT_CREDIT_TRX_TYPE,
    DEFAULT_CURRENCY_CODE,
    DEFAULT_DEBIT_TRX_TYPE,
    DEFAULT_FLEXFIELD_CONTEXT,
    DEFAULT_LINE_TYPE,
    DEFAULT_PAYMENT_TERMS,
    DEFAULT_QUANTITY,
    DEFAULT_TAXABLE_FLAG,
    ORACLE_FBDI_TOTAL_COLUMNS,
)
from loblaws_automation.models import ReconciliationBatch, TransactionItem


def build_fbdi_row(item: TransactionItem, trx_date: str, gl_date: str) -> list[str]:
    """
    Constructs a single SQL*Loader-compatible row for RA_INTERFACE_LINES_ALL
    with the exact 19 required Oracle interface columns populated.

    Args:
        item: Reconciled transaction line item.
        trx_date: Date string for Transaction Date.
        gl_date: Date string for GL Accounting Date.

    Returns:
        List of 300 string elements (comma-separated positions).
    """
    # 0-indexed array representing up to 300 columns
    row = ["" for _ in range(ORACLE_FBDI_TOTAL_COLUMNS)]

    # Col B (index 1): Business Unit Name
    row[1] = DEFAULT_BUSINESS_UNIT
    # Col C (index 2): Transaction Batch Source Name
    row[2] = DEFAULT_BATCH_SOURCE
    # Col D (index 3): Transaction Type Name (Conditional: CR if amount < 0, else DR)
    row[3] = DEFAULT_CREDIT_TRX_TYPE if item.amount < Decimal("0") else DEFAULT_DEBIT_TRX_TYPE
    # Col E (index 4): Payment Terms
    row[4] = DEFAULT_PAYMENT_TERMS
    # Col F (index 5): Transaction Date
    row[5] = trx_date
    # Col G (index 6): Accounting Date
    row[6] = gl_date
    # Col H (index 7): Transaction Number
    row[7] = item.trx_number
    # Col T (index 19): Bill-to Customer Account Number
    row[19] = item.account_number
    # Col U (index 20): Bill-to Customer Site Number
    row[20] = item.account_number
    # Col AA (index 26): Transaction Line Type
    row[26] = DEFAULT_LINE_TYPE
    # Col AB (index 27): Transaction Line Description
    row[27] = item.trx_number
    # Col AC (index 28): Currency Code
    row[28] = DEFAULT_CURRENCY_CODE
    # Col AD (index 29): Currency Conversion Type
    row[29] = DEFAULT_CONVERSION_TYPE
    # Col AF (index 31): Currency Conversion Rate
    row[31] = DEFAULT_CONVERSION_RATE
    # Col AG (index 32): Transaction Line Amount
    row[32] = f"{item.amount:.2f}"
    # Col AH (index 33): Transaction Line Quantity
    row[33] = DEFAULT_QUANTITY
    # Col AL (index 37): Line Transactions Flexfield Context
    row[37] = DEFAULT_FLEXFIELD_CONTEXT
    # Col AM (index 38): Line Transactions Flexfield Segment 1
    row[38] = item.trx_number
    # Col CM (index 90): Taxable Flag
    row[90] = DEFAULT_TAXABLE_FLAG

    return row


def generate_fbdi_csv(batch: ReconciliationBatch) -> str:
    """
    Generates the complete text content of RaInterfaceLinesAll.csv.

    Args:
        batch: Reconciled payment and deduction batch.

    Returns:
        CSV string content with standard carriage return/newline endings.
    """
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer, lineterminator="\r\n")

    for item in batch.transactions:
        row = build_fbdi_row(item, batch.cheque_date, batch.cheque_date)
        writer.writerow(row)

    return csv_buffer.getvalue()


def create_fbdi_zip_package(batch: ReconciliationBatch) -> tuple[bytes, str]:
    """
    Creates an in-memory ZIP package containing RaInterfaceLinesAll.csv
    ready for upload to Oracle Fusion Cloud UCM.

    Args:
        batch: Reconciled payment and deduction batch.

    Returns:
        Tuple of (Binary ZIP bytes, Suggested ZIP File Name).
    """
    csv_content = generate_fbdi_csv(batch)
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        # Standard Oracle FBDI file name inside ZIP
        z.writestr("RaInterfaceLinesAll.csv", csv_content.encode("utf-8"))

    zip_filename = f"AutoInvoiceImport_{batch.cheque_number}.zip"
    return zip_buffer.getvalue(), zip_filename
