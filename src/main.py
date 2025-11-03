from pathlib import Path
from excel_reader import read_excel_data
from comparer import compare_bills
from reporting import save_comparison_report
from models import ComparisonReport, BillRecord
from qb_gateway import fetch_bills_from_qb, add_bill_to_qb

excel_file = Path("C:/Users/PavuluriA/Quickbook Services Bill/QB_Connector_Service_Bill_Python_Fall_2025/company_data.xlsx")
report_path = Path("comparison_report.json")

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
    print(f"Adding {len(comparison_report.excel_only)} Excel-only bills to QuickBooks...")
    add_bill_to_qb(comparison_report.excel_only)
    for bill in comparison_report.excel_only:
        bill.added_to_qb = True

# Step 5: Save report
save_comparison_report(comparison_report, report_path)
print("Report saved and processing complete.")
