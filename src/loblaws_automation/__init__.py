"""
Loblaws AR Remittance Reconciliation & Oracle Fusion AutoInvoice Automation.

A production-grade, enterprise-compliant reconciliation engine for Grupo Bimbo (BU CACBC)
and Loblaw Companies Limited (Loblaws).
"""

from loblaws_automation.ledger import ReconciliationLedger, compute_file_hash
from loblaws_automation.matcher import LoblawsUpstreamMatcher, MatchingSummary
from loblaws_automation.reconciler import (
    BalancingException,
    LoblawsAutomationError,
    LoblawsReconciler,
    RowLossException,
    WorkbookSchemaException,
)

__version__ = "1.2.0"
__author__ = "Principal Solutions Architecture & Financial Engineering"

__all__ = [
    "BalancingException",
    "LoblawsAutomationError",
    "LoblawsReconciler",
    "LoblawsUpstreamMatcher",
    "MatchingSummary",
    "ReconciliationLedger",
    "RowLossException",
    "WorkbookSchemaException",
    "compute_file_hash",
]
