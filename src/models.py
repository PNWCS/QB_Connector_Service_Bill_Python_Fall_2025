"""Models for bill comparison between Excel and QuickBooks."""

from __future__ import annotations
from dataclasses import dataclass, asdict, field
from typing import List, Literal, Optional
import json


# ------------------------------------------------------------
# TYPE DEFINITIONS
# ------------------------------------------------------------
SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["missing_in_excel", "missing_in_quickbooks", "data_mismatch"]


# ------------------------------------------------------------
# CORE DATA MODELS
# ------------------------------------------------------------
@dataclass
class BillRecord:
    """Represents a single bill record from Excel or QuickBooks."""

    record_id: str  # Parent ID or Child ID
    supplier: Optional[str] = None
    bank_date: Optional[str] = None
    chart_account: Optional[str] = None  # Tier 2 - Chart of Account or similar
    amount: Optional[float] = None
    memo: Optional[str] = None
    line_memo: Optional[str] = None
    source: SourceLiteral = "excel"

    def __repr__(self):
        return f"<BillRecord {self.record_id} ({self.source})>"


# ------------------------------------------------------------
# CONFLICT MODEL
# ------------------------------------------------------------
@dataclass
class Conflict:
    """Represents a data mismatch or missing record between Excel and QuickBooks."""

    record_id: str
    excel_value: Optional[str]
    qb_value: Optional[str]
    reason: ConflictReason


# ------------------------------------------------------------
# REPORT MODEL
# ------------------------------------------------------------
@dataclass
class ComparisonReport:
    """Contains the full comparison results between Excel and QuickBooks data."""

    excel_only: List[BillRecord] = field(default_factory=list)
    qb_only: List[BillRecord] = field(default_factory=list)
    conflicts: List[Conflict] = field(default_factory=list)
    matched: List[BillRecord] = field(default_factory=list)

    def to_json(self, path: Optional[str] = None) -> str:
        """Convert comparison results to JSON (and optionally write to a file)."""
        data = {
            "excel_only": [asdict(item) for item in self.excel_only],
            "qb_only": [asdict(item) for item in self.qb_only],
            "conflicts": [asdict(conflict) for conflict in self.conflicts],
            "matched": [asdict(item) for item in self.matched],
            "summary": {
                "total_excel_only": len(self.excel_only),
                "total_qb_only": len(self.qb_only),
                "total_conflicts": len(self.conflicts),
                "total_matched": len(self.matched),
            },
        }

        json_str = json.dumps(data, indent=4)

        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(json_str)

        return json_str


@dataclass
class BillRecord:
    record_id: str
    supplier: Optional[str] = None
    bank_date: Optional[str] = None
    chart_account: Optional[str] = None
    amount: Optional[float] = None
    memo: Optional[str] = None
    line_memo: Optional[str] = None
    source: str = "excel"
    added_to_qb: bool = False  # NEW FIELD


@dataclass
class Conflict:
    record_id: str
    excel_name: Optional[str]
    qb_name: Optional[str]
    reason: str


@dataclass
class ComparisonReport:
    excel_only: List[BillRecord]
    qb_only: List[BillRecord]
    conflicts: List[Conflict]


__all__ = [
    "BillRecord",
    "Conflict",
    "ComparisonReport",
    "SourceLiteral",
    "ConflictReason",
]
