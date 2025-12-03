from typing import List
from src.models import BillRecord, ComparisonReport, Conflict


def compare_bills(
    excel_bills: List[BillRecord], qb_bills: List[BillRecord]
) -> ComparisonReport:
    excel_by_id = {b.record_id: b for b in excel_bills}
    qb_by_id = {b.record_id: b for b in qb_bills}

    excel_only = [b for rid, b in excel_by_id.items() if rid not in qb_by_id]
    qb_only = [b for rid, b in qb_by_id.items() if rid not in excel_by_id]

    conflicts = []
    matched = []

    for rid in excel_by_id.keys() & qb_by_id.keys():
        excel_bill = excel_by_id[rid]
        qb_bill = qb_by_id[rid]

        # If any field mismatches → conflict
        if (
            excel_bill.supplier != qb_bill.supplier
            or excel_bill.amount != qb_bill.amount
            or excel_bill.chart_account != qb_bill.chart_account
            or excel_bill.line_memo != qb_bill.line_memo
        ):
            conflicts.append(
                Conflict(
                    record_id=rid,
                    reason="data_mismatch",
                    excel_bill=excel_bill,
                    qb_bill=qb_bill,
                    excel_name=(
                        f"Supplier: {excel_bill.supplier}, "
                        f"Amount: {excel_bill.amount}, "
                        f"CoA: {excel_bill.chart_account}, "
                        f"LineMemo: {excel_bill.line_memo}"
                    ),
                    qb_name=(
                        f"Supplier: {qb_bill.supplier}, "
                        f"Amount: {qb_bill.amount}, "
                        f"CoA: {qb_bill.chart_account}, "
                        f"LineMemo: {qb_bill.line_memo}"
                    ),
                )
            )
        else:
            # MATCHED RECORD
            matched.append(excel_bill)

    return ComparisonReport(
        excel_only=excel_only,
        qb_only=qb_only,
        conflicts=conflicts,
        matched=matched,
    )
