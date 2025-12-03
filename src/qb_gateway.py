import win32com.client
import xml.etree.ElementTree as ET
from datetime import datetime
from src.models import BillRecord


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

    if status_code not in (0, 1):  # 0 = success, 1 = warning
        raise RuntimeError(f"QuickBooks error ({status_code}): {status_message}")

    return root


def _send_qbxml(qbxml: str) -> ET.Element:
    """Send QBXML to QuickBooks and return parsed response."""
    APP_NAME = "Quickbooks Connector"

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
        parent_id = bill_ret.findtext("Memo") or ""
        supplier = bill_ret.findtext("VendorRef/FullName") or ""

        txn_date_str = bill_ret.findtext("TxnDate") or ""
        txn_date = None
        if txn_date_str:
            try:
                txn_date = datetime.strptime(txn_date_str, "%Y-%m-%d").date()
            except ValueError:
                pass

        memo = bill_ret.findtext("Memo") or ""

        # Loop over bill line items
        for line in bill_ret.findall(".//ExpenseLineRet"):
            bills.append(
                BillRecord(
                    record_id=parent_id,
                    supplier=supplier,
                    bank_date=txn_date,
                    memo=memo,
                    chart_account=line.findtext("AccountRef/FullName") or "",
                    amount=float(line.findtext("Amount") or 0),
                    line_memo=line.findtext("Memo") or "",
                    source="quickbooks",
                )
            )

    return bills


def add_bill_to_qb(bills: BillRecord | list[BillRecord]) -> list[BillRecord]:
    """Add one or multiple BillRecord(s) from Excel to QuickBooks."""
    if not isinstance(bills, list):
        bills = [bills]

    qbxml_batch = []

    for bill in bills:
        if not bill.supplier:
            print(f"Skipping bill {bill.record_id}: missing supplier.")
            continue

        if not bill.amount or bill.amount <= 0:
            print(f"Skipping bill {bill.record_id}: invalid amount {bill.amount}.")
            continue

        txn_date = ""
        if bill.bank_date:
            try:
                if isinstance(bill.bank_date, datetime):
                    txn_date = bill.bank_date.strftime("%Y-%m-%d")
                else:
                    txn_date = str(bill.bank_date).split(" ")[0]
            except Exception:
                txn_date = str(bill.bank_date)

        expense_line = ""
        if bill.chart_account:
            expense_line = (
                "        <ExpenseLineAdd>\n"
                f"          <AccountRef><FullName>{_escape_xml(bill.chart_account)}</FullName></AccountRef>\n"
                f"          <Amount>{bill.amount:.2f}</Amount>\n"
                f"          <Memo>{_escape_xml(bill.line_memo or '')}</Memo>\n"
                "        </ExpenseLineAdd>\n"
            )

        qbxml_batch.append(
            "      <BillAddRq>\n"
            "        <BillAdd>\n"
            f"          <VendorRef><FullName>{_escape_xml(bill.supplier)}</FullName></VendorRef>\n"
            f"          <TxnDate>{txn_date}</TxnDate>\n"
            f"          <Memo>{_escape_xml(bill.memo or '')}</Memo>\n"
            f"{expense_line}"
            "        </BillAdd>\n"
            "      </BillAddRq>\n"
        )

    if not qbxml_batch:
        print("No valid bills to add.")
        return bills

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        "  <QBXMLMsgsRq onError='stopOnError'>\n"
        + "".join(qbxml_batch)
        + "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        _send_qbxml(qbxml)
        for bill in bills:
            bill.added_to_qb = True
            print(f"Successfully added bill to QuickBooks: {bill.record_id}")

    except Exception as e:
        print(f"Failed to add bills: {e}")
        print("QBXML sent:\n", qbxml)
        for bill in bills:
            bill.added_to_qb = False

    return bills
