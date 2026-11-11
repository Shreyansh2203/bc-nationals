"""
Deterministic regional routing and claim taxonomy classification engine.
"""


from loblaws_automation.config import (
    CLAIM_CATEGORY_MAP,
    CUSTOMER_ACCOUNT_MAP,
    REVERSE_CUSTOMER_ACCOUNT_MAP,
)


def classify_invoice_region(invoice_number: str) -> tuple[str, str]:
    """
    Applies the mathematical routing rule extracted from production Excel formulas:
    - West: Length 12 and leading digit '2'
    - Ontario: Length 11 (any leading digit)
    - Atlantic: Length 12 and leading digit '1'
    - Quebec: Length 12 and leading digit '3'

    Args:
        invoice_number: The raw or normalized invoice string.

    Returns:
        Tuple of (Division Name, Customer Account Number)
    """
    clean_inv = str(invoice_number).strip()
    length = len(clean_inv)
    leading_digit = clean_inv[0] if length > 0 else ""

    if length == 12 and leading_digit == "2":
        division = "West"
    elif length == 11:
        division = "Ontario"
    elif length == 12 and leading_digit == "1":
        division = "Atlantic"
    elif length == 12 and leading_digit == "3":
        division = "Quebec"
    else:
        division = "Unassigned"

    account = CUSTOMER_ACCOUNT_MAP.get(division, "60503130")
    return division, account


def identify_claim_category(trx_number: str) -> str:
    """
    Identifies the deduction claim category from transaction prefix.

    Args:
        trx_number: The transaction identifier string.

    Returns:
        Human-readable category description.
    """
    trx_clean = str(trx_number).strip()
    for prefix, label in CLAIM_CATEGORY_MAP.items():
        if trx_clean.startswith(prefix):
            return label

    if trx_clean.isdigit():
        return "8-digit Cleared Invoice"

    return "Special / Market Adjustment"


def get_division_from_account(account_number: str) -> str:
    """Returns the regional division name for a given customer account number."""
    return REVERSE_CUSTOMER_ACCOUNT_MAP.get(str(account_number).strip(), "Other / Clearing")
