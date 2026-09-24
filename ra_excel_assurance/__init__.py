"""RA Excel Assurance Engine."""
__version__ = "0.1.0"

from .auditor import AuditConfig, AuditResult, Issue, audit_workbook

__all__ = ["AuditConfig", "AuditResult", "Issue", "audit_workbook"]
