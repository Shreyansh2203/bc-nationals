"""
Enterprise configuration, business rules, and Oracle Fusion AR constants.
"""

from typing import Final

# -----------------------------------------------------------------------------
# Oracle Fusion Cloud ERP AutoInvoice Constants
# -----------------------------------------------------------------------------
DEFAULT_BUSINESS_UNIT: Final[str] = "CACBC"               # Canada Bread Company, Ltd.
DEFAULT_BATCH_SOURCE: Final[str] = "CACBC_IMPORTED"       # External Deduction Staging Source
DEFAULT_CREDIT_TRX_TYPE: Final[str] = "CACBC DEDUC-MGMT-CR" # Deduction Credit Memo
DEFAULT_DEBIT_TRX_TYPE: Final[str] = "CACBC DEDUC-MGMT-DR"  # Deduction Debit Memo
DEFAULT_PAYMENT_TERMS: Final[str] = "Net 7 Days"          # Standard Credit Terms
DEFAULT_CURRENCY_CODE: Final[str] = "CAD"                 # Canadian Dollars
DEFAULT_CONVERSION_TYPE: Final[str] = "User"              # Direct User Conversion
DEFAULT_CONVERSION_RATE: Final[str] = "1"                 # CAD to CAD Base Rate
DEFAULT_LINE_TYPE: Final[str] = "LINE"                    # Line Type Identifier
DEFAULT_QUANTITY: Final[str] = "1"                        # Standard Quantity
DEFAULT_FLEXFIELD_CONTEXT: Final[str] = "CB_IMPORTED"     # Oracle Descriptive Flexfield
DEFAULT_TAXABLE_FLAG: Final[str] = "N"                    # Non-taxable adjustment flag

# Total column width expected by standard Oracle SQL*Loader control file (RaInterfaceLinesAll.ctl)
ORACLE_FBDI_TOTAL_COLUMNS: Final[int] = 300

# -----------------------------------------------------------------------------
# Customer Account Number Crosswalk (Canadian Operating Divisions)
# -----------------------------------------------------------------------------
CUSTOMER_ACCOUNT_MAP: Final[dict[str, str]] = {
    "Atlantic": "60455049",  # Newfoundland, Nova Scotia, New Brunswick, PEI
    "West": "60463886",      # British Columbia, Manitoba, Saskatchewan, Alberta
    "Ontario": "60477290",   # Greater Toronto Area & Ontario Provinces
    "Quebec": "60503134",    # Quebec Stores
    "Discounts": "60503130"  # Corporate Trade Allowances & Small Balance Write-Offs
}

REVERSE_CUSTOMER_ACCOUNT_MAP: Final[dict[str, str]] = {
    v: k for k, v in CUSTOMER_ACCOUNT_MAP.items()
}

# -----------------------------------------------------------------------------
# Claim Code Taxonomy & Operational Business Rules (Sept. 8, 2026 Walkthrough)
# -----------------------------------------------------------------------------
CLAIM_CATEGORY_MAP: Final[dict[str, str]] = {
    "OP-": "Operations (Quantity / Price Variance)",
    "NF-": "Not Found in Open AR (Pending EDI Interface / Holiday Lag)",
    "INASCR-": "Associate Credit (Credit in Bimbo system, paid by Loblaws)",
    "INASDR-": "Associate Debit (Debit in Bimbo system, claimed by Loblaws)",
    "SP-": "Short Payment (Both negative with negative calculated difference)",
    "OV-": "Over Payment (Both positive with remittance exceeding open item)",
    "NAC": "Master Batch Balancing Debit",
}

# Small Balance Write-Off Rule (Mehtab & Linda Policy):
# Absolute differences under $11.00 are treated as automatic write-offs.
# Differences of $11.00 or more are classified as discrepancies / deduction setups.
WRITE_OFF_THRESHOLD: Final[float] = 11.00

# Operational Schedule & Cadence
REMITTANCE_SCHEDULE: Final[str] = "Twice weekly (Mondays and Fridays)"
REMITTANCE_INBOX: Final[str] = "CB Accounts"

# Floating-point tolerance for zero-sum balancing assertion
ZERO_SUM_TOLERANCE: Final[float] = 0.005
