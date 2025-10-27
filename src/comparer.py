from __future__ import annotations
from typing import Iterable
from models import ComparisonReport, Conflict, BillRecord


def compare_bills(
    excel_bills: Iterable[BillRecord], qb_bills: Iterable[BillRecord]
) -> ComparisonReport:
    """
    Compare Excel bills and QuickBooks bills.
    Parent bills: compare record_id and memo
    Child bills: compare record_id and line_memo
    """
    excel_dict = {bill.record_id: bill for bill in excel_bills}
    qb_dict = {bill.record_id: bill for bill in qb_bills}

    excel_only, qb_only, conflicts = [], [], []

    all_ids = set(excel_dict.keys()) | set(qb_dict.keys())

    for record_id in all_ids:
        excel_bill = excel_dict.get(record_id)
        qb_bill = qb_dict.get(record_id)

        # Excel only
        if excel_bill and not qb_bill:
            excel_only.append(excel_bill)

        # QuickBooks only
        elif qb_bill and not excel_bill:
            qb_only.append(qb_bill)

        # Both exist → compare
        elif excel_bill and qb_bill:
            mismatch_fields = []

            # Parent vs memo
            if (
                excel_bill.line_memo == ""
                and qb_bill.line_memo == ""
                and excel_bill.memo != qb_bill.memo
            ):
                mismatch_fields.append("Memo mismatch")

            # Child vs line_memo
            if (
                excel_bill.line_memo
                and qb_bill.line_memo
                and excel_bill.line_memo != qb_bill.line_memo
            ):
                mismatch_fields.append("LineMemo mismatch")

            # Other checks
            if excel_bill.supplier != qb_bill.supplier:
                mismatch_fields.append("Supplier/Vendor mismatch")
            if str(excel_bill.amount) != str(qb_bill.amount):
                mismatch_fields.append("Amount mismatch")
            if excel_bill.chart_account != qb_bill.chart_account:
                mismatch_fields.append("ChartAccount mismatch")

            if mismatch_fields:
                conflicts.append(
                    Conflict(
                        record_id=record_id,
                        excel_name=str(excel_bill) if excel_bill else None,
                        qb_name=str(qb_bill) if qb_bill else None,
                        reason=", ".join(mismatch_fields),
                    )
                )

    return ComparisonReport(excel_only=excel_only, qb_only=qb_only, conflicts=conflicts)
