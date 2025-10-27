import win32com.client
import xml.etree.ElementTree as ET
from models import BillRecord
from datetime import datetime


def _escape_xml(value: str) -> str:
    """Escape XML special characters for QBXML requests."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _parse_response(raw_xml: str) -> ET.Element:
    """Parse QBXML response and validate status."""
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    if status_code not in (0, 1):
        raise RuntimeError(f"QuickBooks error ({status_code}): {status_message}")
    return root


def _send_qbxml(qbxml: str) -> ET.Element:
    """Send QBXML to QuickBooks and return parsed response."""
    APP_NAME = "QB Connector Service Bill Python Fall 2025"
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    session.OpenConnection2("", APP_NAME, 1)
    ticket = session.BeginSession("", 0)
    try:
        raw_response = session.ProcessRequest(ticket, qbxml)
        return _parse_response(raw_response)
    finally:
        session.EndSession(ticket)
        session.CloseConnection()


def fetch_bills_from_qb() -> list[BillRecord]:
    """Fetch all bills from QuickBooks as BillRecord objects."""
    qbxml = """<?xml version="1.0"?>
    <?qbxml version="16.0"?>
    <QBXML>
      <QBXMLMsgsRq onError="stopOnError">
        <BillQueryRq>
          <IncludeLineItems>true</IncludeLineItems>
        </BillQueryRq>
      </QBXMLMsgsRq>
    </QBXML>"""

    root = _send_qbxml(qbxml)
    bills: list[BillRecord] = []

    for bill_ret in root.findall(".//BillRet"):
        parent_id = bill_ret.findtext("TxnID") or ""
        supplier = bill_ret.findtext("VendorRef/FullName") or ""
        txn_date = bill_ret.findtext("TxnDate") or ""
        due_date = bill_ret.findtext("DueDate") or ""
        memo = bill_ret.findtext("Memo") or ""

        # Add parent bill record
        bills.append(
            BillRecord(
                record_id=parent_id,
                supplier=supplier,
                bank_date=txn_date,
                chart_account="",
                amount=float(bill_ret.findtext("AmountDue") or 0),
                memo=memo,
                line_memo="",
                source="quickbooks",
            )
        )

        # Add line items as child records
        for line in bill_ret.findall(".//ExpenseLineRet"):
            line_id = line.findtext("TxnLineID") or ""
            line_amount = float(line.findtext("Amount") or 0)
            line_memo = line.findtext("Memo") or ""
            bills.append(
                BillRecord(
                    record_id=line_id,
                    supplier=supplier,
                    bank_date=txn_date,
                    chart_account=line.findtext("AccountRef/FullName") or "",
                    amount=line_amount,
                    memo=memo,
                    line_memo=line_memo,
                    source="quickbooks",
                )
            )

    return bills


def add_bill_to_qb(bill: BillRecord) -> BillRecord:
    """Add a BillRecord from Excel to QuickBooks with validated QBXML."""
    # --- Validate key fields ---
    if not bill.supplier:
        print(f"Skipping bill {bill.record_id}: missing supplier.")
        return bill
    if not bill.amount or bill.amount <= 0:
        print(f"Skipping bill {bill.record_id}: invalid amount {bill.amount}.")
        return bill

    # --- Format date for QuickBooks ---
    txn_date = ""
    if bill.bank_date:
        try:
            # Handle both datetime and string types
            if isinstance(bill.bank_date, datetime):
                txn_date = bill.bank_date.strftime("%Y-%m-%d")
            else:
                txn_date = str(bill.bank_date).split(" ")[0]
        except Exception:
            txn_date = str(bill.bank_date)

    # --- Build ExpenseLineAdd only if meaningful ---
    expense_line = ""
    if bill.chart_account:
        expense_line = (
            "    <ExpenseLineAdd>\n"
            f"      <AccountRef><FullName>{_escape_xml(bill.chart_account)}</FullName></AccountRef>\n"
            f"      <Amount>{bill.amount:.2f}</Amount>\n"
            f"      <Memo>{_escape_xml(bill.line_memo or '')}</Memo>\n"
            "    </ExpenseLineAdd>\n"
        )

    # --- Build main QBXML body ---
    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillAddRq>\n"
        "      <BillAdd>\n"
        f"        <VendorRef><FullName>{_escape_xml(bill.supplier)}</FullName></VendorRef>\n"
        f"        <TxnDate>{txn_date}</TxnDate>\n"
        f"        <Memo>{_escape_xml(bill.memo or '')}</Memo>\n"
        f"{expense_line}"
        "      </BillAdd>\n"
        "    </BillAddRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        response = _send_qbxml(qbxml)
        bill.added_to_qb = True
        print(f" Successfully added bill to QuickBooks: {bill.record_id}")
    except Exception as e:
        print(f" Failed to add bill {bill.record_id}: {e}")
        print("QBXML sent:\n", qbxml)
        bill.added_to_qb = False

        return bill
