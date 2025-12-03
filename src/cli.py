import argparse
from pathlib import Path
from src.excel_reader import read_excel_data
from src.comparer import compare_bills
from src.reporting import save_comparison_report
from src.qb_gateway import fetch_bills_from_qb, add_bill_to_qb


def run():
    parser = argparse.ArgumentParser(description="QuickBooks Comparison Tool")
    parser.add_argument("--workbook", required=True, help="Path to Excel workbook")
    parser.add_argument(
        "--report", default="comparison_report.json", help="Output report path"
    )
    args = parser.parse_args()

    excel_file = Path(args.workbook)
    report_path = Path(args.report)

    # Step 1: Read Excel
    excel_rows = read_excel_data(excel_file)
    print(f"Total Excel bills: {len(excel_rows)}")

    # Step 2: Fetch QuickBooks bills
    qb_bills = fetch_bills_from_qb()
    print(f"Total QuickBooks bills: {len(qb_bills)}")

    # Step 3: Compare bills
    comparison_report = compare_bills(excel_rows, qb_bills)
    print("Comparison completed.")

    # Step 4: Batch add Excel-only bills to QB
    if comparison_report.excel_only:
        print(
            f"Adding {len(comparison_report.excel_only)} Excel-only bills to QuickBooks..."
        )
        add_bill_to_qb(comparison_report.excel_only)
        for bill in comparison_report.excel_only:
            bill.added_to_qb = True

    # Step 5: Save report
    save_comparison_report(comparison_report, report_path)
    print(f"Report saved to {report_path}")


if __name__ == "__main__":
    run()
