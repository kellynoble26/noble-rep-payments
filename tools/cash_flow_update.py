#!/usr/bin/env python3
"""Append new payments on Kelly's PC and publish a sealed Cash Flow Carnival page.

Reads individual QuickBooks "Payment received" mail from the GC Outlook
mailbox and the WD Gmail mailbox, plus posted bank-deposit alerts. Batched
"money on its way" mail is ignored. When nothing new is found, this exits
without committing. The passcode is read from FIELD_LEDGER_PASSCODE and is
never printed.

This does not install the scheduled task. Run tools/install_cash_flow_task.ps1
on NobleOfficeMgr for that.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from zoneinfo import ZoneInfo

AZ = ZoneInfo("America/Phoenix")
GC_ADDRESS = "mgr@noblegcllc.com"
WD_ADDRESS = "mgr@noblewandd.com"
FIRST_SINCE = datetime(2026, 11, 20, tzinfo=AZ)

WORKBOOK = Path(os.environ.get(
    "CASH_FLOW_WORKBOOK",
    r"C:\Users\kelly\OneDrive\Desktop\Monies_Received_2026.xlsx",
))
REPO = Path(os.environ.get(
    "CASH_FLOW_REPO",
    r"C:\Users\kelly\NobleOfficeMgr\noble-rep-payments",
))
LOG_PATH = Path(os.environ.get(
    "CASH_FLOW_LOG",
    r"C:\Users\kelly\NobleOfficeMgr\logs\cash-flow-carnival.log",
))
STATE_PATH = Path(os.environ.get(
    "CASH_FLOW_STATE",
    r"C:\Users\kelly\NobleOfficeMgr\cash-flow-carnival\state.json",
))
GMAIL_CLIENT = Path(os.environ.get(
    "CASH_FLOW_GMAIL_CLIENT",
    r"C:\Users\kelly\NobleOfficeMgr\cash-flow-carnival\gmail-client.json",
))
GMAIL_TOKEN = Path(os.environ.get(
    "CASH_FLOW_GMAIL_TOKEN",
    r"C:\Users\kelly\NobleOfficeMgr\cash-flow-carnival\gmail-token.json",
))

BATCH_MAIL = re.compile(
    r"money on its way|moneys on its way|money['’]s on its way",
    re.IGNORECASE,
)
PENDING_MAIL = re.compile(
    r"\b(pending|scheduled to arrive|will be deposited|will arrive|not yet posted)\b",
    re.IGNORECASE,
)
PAYMENT_RECEIVED = re.compile(r"payment received", re.IGNORECASE)
MONEY = re.compile(r"\$\s*([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\.(\d{2})")
INVOICE = re.compile(r"Invoice\s*#?\s*[:.]?\s*(\d{3,6})", re.IGNORECASE)
SUBJECT_CUSTOMER = re.compile(r"Invoice\s*#\d+\s*[-–]\s*\((.+)\)\s*$", re.IGNORECASE)
LABELED_LINE = re.compile(r"(?im)^(customer|paid by)\s*[:\-]?\s*(.+?)\s*$")
TXN = re.compile(
    r"(?i)(?:transaction|payment)\s*(?:id|number|#)\s*[:#]?\s*([A-Z0-9]{6,})"
)
RECEIVED_AMOUNT = re.compile(
    r"(?i)(?:you received(?: a payment of)?|payment of|amount received|deposit(?: of| amount)?)\s*\$\s*([0-9,]+\.\d{2})"
)
BANK_HINT = re.compile(
    r"(?i)\b(avidpay|bank of america|bofa|deposit (?:alert|received|posted)|posted deposit)\b"
)


def money_cents(value: str) -> int:
    quantized = Decimal(value.replace(",", "")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(quantized * 100)


def plain_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style)\b.*?>.*?</\1>", " ", value or "")
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = html.unescape(value)
    return re.sub(r"[ \t\r\f\v]+", " ", value)


def norm_name(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip()).casefold()


def log_line(message: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(AZ).strftime("%Y-%m-%d %H:%M:%S %Z")
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{stamp} {message}\n")


def load_state() -> dict:
    if not STATE_PATH.exists():
        return {"qb_ids": [], "pending_publish": False, "pending_summary": []}
    return json.loads(STATE_PATH.read_text(encoding="utf-8"))


def save_state(state: dict) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def since_when(state: dict) -> datetime:
    raw = state.get("last_success")
    if not raw:
        return FIRST_SINCE
    parsed = datetime.fromisoformat(raw)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=AZ)
    return parsed.astimezone(AZ) - timedelta(days=2)


def parse_message(subject: str, body: str, received: datetime, mailbox: str) -> dict | None:
    """Return one payment, or None when the message is not a received payment."""
    subject = subject or ""
    body = plain_text(body or "")
    combined = f"{subject}\n{body}"
    if BATCH_MAIL.search(combined):
        return None
    if PENDING_MAIL.search(combined):
        return None
    if received.tzinfo is None:
        received = received.replace(tzinfo=AZ)
    received = received.astimezone(AZ)
    if PAYMENT_RECEIVED.search(subject) or PAYMENT_RECEIVED.search(body[:500]):
        return _quickbooks_payment(subject, body, received, mailbox)
    if BANK_HINT.search(combined):
        return _bank_deposit(subject, body, received, mailbox)
    return None


def _amount(subject: str, body: str) -> int | None:
    labeled = RECEIVED_AMOUNT.search(f"{subject}\n{body}")
    if labeled:
        return money_cents(labeled.group(1))
    amounts = [money_cents(whole + "." + cents) for whole, cents in MONEY.findall(body)]
    unique = list(dict.fromkeys(amounts))
    if len(unique) == 1:
        return unique[0]
    return None


def _customer(subject: str, body: str) -> str:
    named = SUBJECT_CUSTOMER.search(subject.strip())
    if named:
        return re.sub(r"\s+", " ", named.group(1)).strip()
    labeled = LABELED_LINE.search(body)
    if labeled:
        return re.sub(r"\s+", " ", labeled.group(2)).strip()
    return ""


def _invoice(subject: str, body: str) -> str:
    match = INVOICE.search(subject) or INVOICE.search(body)
    return match.group(1) if match else ""


def _txn(subject: str, body: str) -> str:
    match = TXN.search(f"{subject}\n{body}")
    return match.group(1) if match else ""


def _quickbooks_payment(subject: str, body: str, received: datetime, mailbox: str) -> dict | None:
    amount = _amount(subject, body)
    if amount is None:
        return None
    txn = _txn(subject, body)
    customer = _customer(subject, body)
    invoice = _invoice(subject, body)
    if not customer and not invoice and not txn:
        return None
    method = f"QuickBooks payment (txn {txn})" if txn else "QuickBooks payment"
    return _record(
        mailbox=mailbox,
        received=received,
        customer=customer,
        invoice=invoice,
        amount=amount,
        method=method,
        source="QuickBooks Payments",
        subject=re.sub(r"\s+", " ", subject).strip()[:240],
        qb_id=txn,
        kind="qb",
    )


def _bank_deposit(subject: str, body: str, received: datetime, mailbox: str) -> dict | None:
    if PAYMENT_RECEIVED.search(subject):
        return None
    amounts = [money_cents(whole + "." + cents) for whole, cents in MONEY.findall(body)]
    unique = list(dict.fromkeys(amounts))
    if len(unique) != 1:
        return None
    labeled = BANK_HINT.search(subject) or BANK_HINT.search(body)
    label = re.sub(r"\s+", " ", labeled.group(0)).strip() if labeled else "Bank deposit"
    customer = _customer(subject, body) or label
    return _record(
        mailbox=mailbox,
        received=received,
        customer=customer,
        invoice="",
        amount=unique[0],
        method=label,
        source="Bank deposit alert",
        subject=re.sub(r"\s+", " ", subject).strip()[:240],
        qb_id="",
        kind="bank",
    )


def _record(**fields) -> dict:
    received = fields["received"]
    day = received.strftime("%Y-%m-%d")
    clock = received.strftime("%H:%M")
    received_text = day if clock == "00:00" else f"{day} {clock}"
    return {
        "co": fields["mailbox"],
        "received": received_text,
        "customer": fields["customer"],
        "invoice": fields["invoice"],
        "invoice_amount": None,
        "amount_cents": fields["amount"],
        "deposit": received.strftime("%m/%d/%Y"),
        "method": fields["method"],
        "source": fields["source"],
        "subject": fields["subject"],
        "qb_id": fields["qb_id"],
        "kind": fields["kind"],
    }


def fingerprint(item: dict) -> tuple:
    return (
        item["co"],
        norm_name(item["customer"]),
        item["invoice"],
        item["amount_cents"],
        item["received"][:10],
    )


def bank_key(item: dict) -> tuple:
    return (item["received"][:10], item["amount_cents"], norm_name(item["customer"]))


def workbook_keys(path: Path) -> tuple[set[str], set[tuple], set[tuple]]:
    from openpyxl import load_workbook

    qb_ids: set[str] = set()
    prints: set[tuple] = set()
    banks: set[tuple] = set()
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        for sheet in ("GC", "WD"):
            ws = wb[sheet]
            for values in ws.iter_rows(min_row=2, max_col=10, values_only=True):
                if not values or all(value is None for value in values):
                    continue
                mailbox, received, customer, invoice, _invoice_amount, amount, _deposit, method, _source, subject = values
                blob = f"{method or ''}\n{subject or ''}"
                for txn in TXN.findall(blob):
                    qb_ids.add(txn.upper())
                txn_short = re.findall(r"(?i)txn\s+([A-Z0-9]{6,})", blob)
                qb_ids.update(item.upper() for item in txn_short)
                cents = money_cents(f"{Decimal(str(amount)):.2f}") if amount not in (None, "") else None
                day = str(received or "")[:10].replace("/", "-")
                if re.fullmatch(r"\d{2}-\d{2}-\d{4}", day):
                    month, dom, year = day.split("-")
                    day = f"{year}-{month}-{dom}"
                if cents is not None:
                    prints.add((str(mailbox), norm_name(str(customer or "")), str(invoice or ""), cents, day))
                    if "bank deposit" in str(_source or "").casefold() or "avidpay" in blob.casefold():
                        banks.add((day, cents, norm_name(str(customer or ""))))
    finally:
        wb.close()
    return qb_ids, prints, banks


def is_new(item: dict, qb_ids: set[str], prints: set[tuple], banks: set[tuple]) -> bool:
    if item["qb_id"] and item["qb_id"].upper() in qb_ids:
        return False
    if fingerprint(item) in prints:
        return False
    if item["kind"] == "bank" and bank_key(item) in banks:
        return False
    return True


def append_rows(path: Path, items: list[dict]) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path)
    for item in items:
        sheet = wb[item["co"]]
        amount = Decimal(item["amount_cents"]) / Decimal(100)
        sheet.append([
            item["co"],
            item["received"],
            item["customer"],
            item["invoice"],
            item["invoice_amount"],
            float(amount),
            item["deposit"],
            item["method"],
            item["source"],
            item["subject"],
        ])
    wb.save(path)


def read_outlook(since: datetime) -> list[dict]:
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    store = None
    for account in namespace.Accounts:
        address = str(getattr(account, "SmtpAddress", "") or "").casefold()
        if address == GC_ADDRESS:
            store = account.DeliveryStore
            break
    if store is None:
        raise RuntimeError(f"Outlook is signed in, but {GC_ADDRESS} is not one of the accounts")
    inbox = store.GetDefaultFolder(6)
    found: list[dict] = []
    folders = [inbox, *list(inbox.Folders)]
    for folder in folders:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        count = 0
        for item in items:
            count += 1
            if count > 400:
                break
            received = getattr(item, "ReceivedTime", None)
            if received is None:
                continue
            when = datetime.fromtimestamp(received.timestamp(), AZ) if hasattr(received, "timestamp") else since
            try:
                when = datetime(
                    received.year, received.month, received.day,
                    received.hour, received.minute, tzinfo=AZ,
                )
            except Exception:
                continue
            if when < since:
                break
            subject = str(getattr(item, "Subject", "") or "")
            body = str(getattr(item, "Body", "") or "") or str(getattr(item, "HTMLBody", "") or "")
            parsed = parse_message(subject, body, when, "GC")
            if parsed:
                found.append(parsed)
    return found


def gmail_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/gmail.readonly"]
    creds = Credentials.from_authorized_user_file(str(GMAIL_TOKEN), scopes)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        GMAIL_TOKEN.write_text(creds.to_json(), encoding="utf-8")
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def gmail_consent() -> None:
    from google_auth_oauthlib.flow import InstalledAppFlow

    if not GMAIL_CLIENT.exists():
        raise SystemExit(
            f"Save the Gmail desktop OAuth client to {GMAIL_CLIENT} first. "
            "The installer steps name that file."
        )
    scopes = ["https://www.googleapis.com/auth/gmail.readonly"]
    flow = InstalledAppFlow.from_client_secrets_file(str(GMAIL_CLIENT), scopes)
    creds = flow.run_local_server(port=0, open_browser=True)
    GMAIL_TOKEN.parent.mkdir(parents=True, exist_ok=True)
    GMAIL_TOKEN.write_text(creds.to_json(), encoding="utf-8")
    print(f"Gmail consent saved for {WD_ADDRESS}. The token is only on this PC.")


def read_gmail(since: datetime) -> list[dict]:
    service = gmail_service()
    query = (
        f"after:{since.strftime('%Y/%m/%d')} "
        '(subject:"payment received" OR subject:deposit OR subject:AvidPay '
        'OR subject:"Bank of America") -subject:"money on its way"'
    )
    response = service.users().messages().list(userId="me", q=query, maxResults=100).execute()
    found: list[dict] = []
    for item in response.get("messages", []):
        message = service.users().messages().get(userId="me", id=item["id"], format="full").execute()
        headers = {h["name"].casefold(): h["value"] for h in message.get("payload", {}).get("headers", [])}
        subject = headers.get("subject", "")
        body = _gmail_body(message.get("payload", {}))
        internal = datetime.fromtimestamp(int(message.get("internalDate", "0")) / 1000, AZ)
        parsed = parse_message(subject, body, internal, "WD")
        if parsed:
            found.append(parsed)
    return found


def _gmail_body(payload: dict) -> str:
    import base64

    parts = [payload]
    text = []
    while parts:
        part = parts.pop(0)
        parts.extend(part.get("parts") or [])
        data = (part.get("body") or {}).get("data")
        if data and part.get("mimeType", "").startswith("text/"):
            text.append(base64.urlsafe_b64decode(data.encode()).decode("utf-8", "replace"))
    return "\n".join(text)


def git(*args: str, cwd: Path) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=cwd,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout


def publish(repo: Path, added: int) -> None:
    status = git("status", "--porcelain", cwd=repo)
    dirty = [line for line in status.splitlines() if line.strip() and not line.endswith(" received.html")]
    if any(not line.startswith("??") and "received.html" not in line for line in status.splitlines()):
        raise RuntimeError("The repo has other changes, so this run will not publish")
    if dirty:
        raise RuntimeError("The repo has other changes, so this run will not publish")
    git("add", "--", "received.html", cwd=repo)
    staged = subprocess.run(
        ["git", "diff", "--cached", "--quiet", "--", "received.html"],
        cwd=repo,
    )
    if staged.returncode == 0:
        return
    sealed = (repo / "received.html").read_text(encoding="utf-8")
    if "AES-GCM" not in sealed or "const PAYMENTS" in sealed:
        raise RuntimeError("received.html is not sealed, so it was not committed")
    git("commit", "-m", f"Seal Cash Flow Carnival update ({added} payments)", cwd=repo)
    git("push", "origin", "HEAD:main", cwd=repo)


def prepare_repo(repo: Path) -> None:
    git("checkout", "main", cwd=repo)
    git("pull", "--ff-only", "origin", "main", cwd=repo)


def build(repo: Path, workbook: Path) -> None:
    sys.path.insert(0, str(repo / "tools"))
    from build_monies_received import seal_workbook

    seal_workbook(workbook, repo / "received.html", repo / "data" / "rep-customer-map.csv")


def check_setup() -> int:
    problems = []
    if not os.environ.get("FIELD_LEDGER_PASSCODE"):
        problems.append("FIELD_LEDGER_PASSCODE is not set for this Windows user")
    if not WORKBOOK.exists():
        problems.append(f"Workbook not found: {WORKBOOK}")
    if not (REPO / "tools" / "build_monies_received.py").exists():
        problems.append(f"Repo not found: {REPO}")
    if not GMAIL_TOKEN.exists():
        problems.append(f"Gmail consent has not been saved: {GMAIL_TOKEN}")
    try:
        import win32com.client  # noqa: F401
    except Exception:
        problems.append("pywin32 is not installed, so Outlook cannot be read")
    if problems:
        for problem in problems:
            print(problem)
        return 1
    print("Setup looks ready. The passcode is set and was not printed.")
    return 0


def run() -> int:
    state = load_state()
    since = since_when(state)
    try:
        incoming = read_outlook(since) + read_gmail(since)
        qb_ids, prints, banks = workbook_keys(WORKBOOK)
        qb_ids.update(item.upper() for item in state.get("qb_ids", []))
        fresh = []
        seen: set[tuple] = set()
        for item in incoming:
            key = (item["qb_id"].upper(), fingerprint(item), bank_key(item) if item["kind"] == "bank" else None)
            if key in seen or not is_new(item, qb_ids, prints, banks):
                continue
            seen.add(key)
            fresh.append(item)
            if item["qb_id"]:
                qb_ids.add(item["qb_id"].upper())
            prints.add(fingerprint(item))
            if item["kind"] == "bank":
                banks.add(bank_key(item))
        if fresh:
            append_rows(WORKBOOK, fresh)
            state["pending_publish"] = True
            state["pending_summary"] = [
                {
                    "co": item["co"],
                    "received": item["received"],
                    "customer": item["customer"],
                    "invoice": item["invoice"],
                    "amount": f"{item['amount_cents'] / 100:.2f}",
                }
                for item in fresh
            ]
            state["qb_ids"] = sorted(qb_ids)
            save_state(state)
        if not state.get("pending_publish"):
            log_line("OK nothing new")
            state["last_success"] = datetime.now(AZ).isoformat(timespec="seconds")
            save_state(state)
            return 0
        prepare_repo(REPO)
        build(REPO, WORKBOOK)
        added = len(state.get("pending_summary") or [])
        publish(REPO, added)
    except Exception as err:
        log_line(f"FAIL {err}")
        return 1
    lines = ["OK added " + str(len(state.get("pending_summary") or [])) + " and pushed the sealed page"]
    for item in state.get("pending_summary") or []:
        lines.append(
            f"  {item['co']} {item['received']} invoice {item['invoice'] or '-'} "
            f"${item['amount']} {item['customer']}"
        )
    log_line("\n".join(lines))
    state["pending_publish"] = False
    state["pending_summary"] = []
    state["last_success"] = datetime.now(AZ).isoformat(timespec="seconds")
    save_state(state)
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--gmail-consent", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    if args.gmail_consent:
        gmail_consent()
        return
    if args.check:
        raise SystemExit(check_setup())
    raise SystemExit(run())


if __name__ == "__main__":
    main()
