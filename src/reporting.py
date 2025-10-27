import json
from pathlib import Path
from models import ComparisonReport, BillRecord
from datetime import datetime


def save_comparison_report(report: ComparisonReport, path: Path):
    """
    Save comparison report to JSON.
    Includes conflicts, Excel-only bills, QuickBooks-only bills, and added-to-QB bills.
    """

    def bill_to_dict(bill: BillRecord):
        return {
            "record_id": bill.record_id,
            "supplier": bill.supplier,
            "bank_date": str(bill.bank_date) if bill.bank_date else None,
            "chart_account": bill.chart_account,
            "amount": bill.amount,
            "memo": bill.memo,
            "line_memo": bill.line_memo,
            "source": bill.source,
            "added_to_qb": getattr(bill, "added_to_qb", False),
        }

    data = {
        "conflicts": [
            {
                "record_id": c.record_id,
                "excel_name": c.excel_name,
                "qb_name": c.qb_name,
                "reason": c.reason,
            }
            for c in report.conflicts
        ],
        "excel_only": [bill_to_dict(bill) for bill in report.excel_only],
        "qb_only": [bill_to_dict(bill) for bill in report.qb_only],
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print(f"Report saved to {path}")
