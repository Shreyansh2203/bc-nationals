"""
Pydantic data models for strong typing, data validation, and enterprise auditability.
"""

from decimal import Decimal

from pydantic import BaseModel, Field, field_validator


class TransactionItem(BaseModel):
    """Represents a single reconciled transaction line (as staged in TXN sheet)."""
    account_number: str = Field(..., description="Customer Bill-To Account Number (e.g. 60455049)")
    amount: Decimal = Field(..., description="Transaction Line Amount (positive or negative)")
    trx_number: str = Field(..., description="Transaction reference string (e.g. NAC2020094821-A or claim)")
    division: str | None = Field(default="Unknown", description="Regional operating division")
    category: str | None = Field(default="Unclassified", description="Claim taxonomy category")

    @field_validator("account_number")
    @classmethod
    def validate_account(cls, v: str) -> str:
        clean = v.strip()
        if not clean:
            raise ValueError("Account number cannot be empty")
        return clean

    @field_validator("trx_number")
    @classmethod
    def validate_trx(cls, v: str) -> str:
        clean = v.strip()
        if not clean:
            raise ValueError("Transaction number cannot be empty")
        return clean


class SkippedRow(BaseModel):
    """Records a source row that could not be parsed, for audit and debugging."""
    row_index: int = Field(..., description="1-based row number in the TXN sheet")
    raw_values: dict = Field(default_factory=dict, description="Raw cell values extracted from the row")
    reason: str = Field(..., description="Human-readable reason the row was skipped")


class AccountMismatch(BaseModel):
    """Tracks discrepancy where an invoice number's regional rule contradicts the customer account."""
    trx_number: str = Field(..., description="Transaction number containing the invoice string")
    account_number: str = Field(..., description="Customer account staged in the workbook")
    staged_division: str = Field(..., description="Division name of the staged account")
    inferred_account: str = Field(..., description="Account number derived from invoice length & prefix")
    inferred_division: str = Field(..., description="Division name derived from invoice length & prefix")
    amount: Decimal = Field(..., description="Transaction amount")


class FBDILine(BaseModel):
    """Represents an Oracle Fusion AutoInvoice interface line (RA_INTERFACE_LINES_ALL)."""
    business_unit: str = "CACBC"
    batch_source: str = "CACBC_IMPORTED"
    trx_type: str
    payment_terms: str = "Net 7 Days"
    trx_date: str
    gl_date: str
    trx_number: str
    bill_to_customer: str
    bill_to_site: str
    line_type: str = "LINE"
    description: str
    currency_code: str = "CAD"
    conversion_type: str = "User"
    conversion_rate: str = "1"
    amount: Decimal
    quantity: str = "1"
    flexfield_context: str = "CB_IMPORTED"
    flexfield_segment1: str
    taxable_flag: str = "N"


class ReconciliationBatch(BaseModel):
    """Represents the complete reconciled payment and deduction batch."""
    cheque_number: str
    cheque_date: str
    total_remittance_amount: Decimal | None = None
    transactions: list[TransactionItem] = Field(default_factory=list)

    @property
    def total_count(self) -> int:
        return len(self.transactions)

    @property
    def net_sum(self) -> Decimal:
        return sum((item.amount for item in self.transactions), Decimal("0.00"))

    @property
    def is_balanced(self) -> bool:
        return abs(self.net_sum) < Decimal("0.01")


class AuditMetrics(BaseModel):
    """Executive metrics and KPI audit summary."""
    batch_id: str
    cheque_number: str
    cheque_date: str
    total_lines: int
    net_variance: Decimal
    is_zero_sum_verified: bool
    total_debits: Decimal
    total_credits: Decimal
    execution_time_ms: float
    division_summary: dict = Field(default_factory=dict)
    claim_summary: dict = Field(default_factory=dict)
    total_source_rows: int = Field(default=0, description="Total data rows found in the TXN sheet before parsing")
    skipped_rows: list[SkippedRow] = Field(default_factory=list, description="Rows that could not be parsed")
    cheque_date_source: str = Field(default="unknown", description="Where the cheque date was resolved from")
    account_mismatches: list[AccountMismatch] = Field(default_factory=list, description="Discrepancies where invoice prefix/length contradicts account division")
    matching_summary: dict | None = Field(default=None, description="Optional upstream matching metrics")
