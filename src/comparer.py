
from typing import List
from models import BillRecord, ComparisonReport, Conflict

def compare_bills(excel_bills: List[BillRecord], qb_bills: List[BillRecord]) -> ComparisonReport:
    """
    Compare Excel and QuickBooks bills.

    Rules:
    1. Excel-only bills → to add to QB
    2. QB-only bills → missing in Excel
    3. Shared record_ids → check for conflicts (amount, CoA, line memo)
    """

    # Index by record_id for O(1) lookup
    excel_by_id = {b.record_id: b for b in excel_bills}
    qb_by_id = {b.record_id: b for b in qb_bills}

    # Excel-only bills
    excel_only = [b for rid, b in excel_by_id.items() if rid not in qb_by_id]

    # QB-only bills
    qb_only = [b for rid, b in qb_by_id.items() if rid not in excel_by_id]

    # Conflicts
    conflicts = []
    for rid in excel_by_id.keys() & qb_by_id.keys():  # Shared record_ids
        excel_bill = excel_by_id[rid]
        qb_bill = qb_by_id[rid]

        # Compare key fields: amount, chart_account, line_memo
        if (excel_bill.amount != qb_bill.amount or
            excel_bill.chart_account != qb_bill.chart_account or
            excel_bill.line_memo != qb_bill.line_memo):
            conflicts.append(
                Conflict(
                    record_id=rid,
                    excel_name=f"Amount: {excel_bill.amount}, CoA: {excel_bill.chart_account}, LineMemo: {excel_bill.line_memo}",
                    qb_name=f"Amount: {qb_bill.amount}, CoA: {qb_bill.chart_account}, LineMemo: {qb_bill.line_memo}",
                    reason="data_mismatch"
                )
            )

    # Return a ComparisonReport
    return ComparisonReport(
        excel_only=excel_only,
        qb_only=qb_only,
        conflicts=conflicts
    )
