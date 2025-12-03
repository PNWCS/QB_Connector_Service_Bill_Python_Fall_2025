"""Models for bill comparison between Excel and QuickBooks."""

from __future__ import annotations
from dataclasses import dataclass, asdict, field
from typing import List, Literal, Optional
import json
from datetime import datetime, date


# TYPE DEFINITIONS
SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["missing_in_excel", "missing_in_quickbooks", "data_mismatch"]


@dataclass
class BillRecord:
    record_id: str
    supplier: Optional[str] = None
    bank_date: Optional[date] = None
    chart_account: Optional[str] = None
    amount: Optional[float] = None
    memo: Optional[str] = None
    line_memo: Optional[str] = None
    source: str = "excel"
    added_to_qb: bool = False

    def __str__(self):
        if isinstance(self.bank_date, (datetime, date)):
            date_str = self.bank_date.strftime("%Y-%m-%d")
        else:
            date_str = str(self.bank_date) if self.bank_date else "N/A"

        return (
            f"BillRecord(record_id={self.record_id}, supplier={self.supplier}, "
            f"bank_date={date_str}, chart_account={self.chart_account}, "
            f"amount={self.amount}, memo={self.memo}, line_memo={self.line_memo}, "
            f"source={self.source}, added_to_qb={self.added_to_qb})"
        )


@dataclass
class Conflict:
    record_id: str
    reason: str

    # full structured bill data
    excel_bill: Optional[BillRecord] = None
    qb_bill: Optional[BillRecord] = None

    # legacy string fields (optional)
    excel_name: Optional[str] = None
    qb_name: Optional[str] = None


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


__all__ = [
    "BillRecord",
    "Conflict",
    "ComparisonReport",
    "SourceLiteral",
    "ConflictReason",
]
