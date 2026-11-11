"""
Domain exceptions for Loblaws AR Reconciliation & Oracle Cloud AutoInvoice Automation.
"""


class LoblawsAutomationError(Exception):
    """Base exception for all domain errors raised by the Loblaws automation suite."""
    pass


class BalancingException(LoblawsAutomationError):
    """Raised when a reconciliation batch fails the zero-sum invariant check."""
    pass


class RowLossException(LoblawsAutomationError):
    """Raised when too many source rows were skipped during parsing."""
    pass


class WorkbookSchemaException(LoblawsAutomationError, ValueError):
    """Raised when the workbook structure or required column headers are missing/invalid."""
    pass
