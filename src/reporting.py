import json
from pathlib import Path
from datetime import datetime
from src.models import ComparisonReport, BillRecord


def _safe_field(bill: BillRecord | None, field: str):
    """
    Safely extract a field from BillRecord even if bill is None.
    """
    if bill is None:
        return None
    return getattr(bill, field, None)


def save_comparison_report(report: ComparisonReport, path: Path):
    """
    Save comparison report to JSON with fully separated fields:
    supplier, amount, chart_account, line_memo.
    Handles all None checks for mypy safety.
    """

    conflicts_output = []

    # ------------------------
    # DATA MISMATCH CONFLICTS
    # ------------------------
    for c in report.conflicts:
        excel = c.excel_bill
        qb = c.qb_bill

        conflicts_output.append(
            {
                "record_id": c.record_id,
                "reason": c.reason,
                "excel_supplier": _safe_field(excel, "supplier"),
                "qb_supplier": _safe_field(qb, "supplier"),
                "excel_amount": _safe_field(excel, "amount"),
                "qb_amount": _safe_field(qb, "amount"),
                "excel_chart_account": _safe_field(excel, "chart_account"),
                "qb_chart_account": _safe_field(qb, "chart_account"),
                "excel_line_memo": _safe_field(excel, "line_memo"),
                "qb_line_memo": _safe_field(qb, "line_memo"),
            }
        )

    # ------------------------
    # MISSING IN EXCEL (QB ONLY)
    # ------------------------
    for qb_bill in report.qb_only:
        conflicts_output.append(
            {
                "record_id": qb_bill.record_id,
                "reason": "missing_in_excel",
                "excel_supplier": None,
                "qb_supplier": qb_bill.supplier,
                "excel_amount": None,
                "qb_amount": qb_bill.amount,
                "excel_chart_account": None,
                "qb_chart_account": qb_bill.chart_account,
                "excel_line_memo": None,
                "qb_line_memo": qb_bill.line_memo,
            }
        )

    # ------------------------
    # FINAL PAYLOAD
    # ------------------------
    payload = {
        "status": "success",
        "generated_at": datetime.utcnow().isoformat(),
        "added_bills": [
            {
                "record_id": bill.record_id,
                "supplier": bill.supplier,
                "bank_date": bill.bank_date.isoformat() if bill.bank_date else None,
                "amount": bill.amount,
                "chart_account": bill.chart_account,
                "memo": bill.memo,
                "line_memo": bill.line_memo,
                "source": bill.source,
                "added_to_qb": getattr(bill, "added_to_qb", False),
            }
            for bill in report.excel_only
        ],
        "conflicts": conflicts_output,
        "same_bills": len(report.matched),
        "error": None,
    }

    # Ensure folder exists
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=4)

    print(f"Report saved to {path}")
