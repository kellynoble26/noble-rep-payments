#!/usr/bin/env python3
"""Build the combined Monies Received report.

Reads the GC and WD tabs of Monies_Received_2026.xlsx. Writes an unlocked
review page (received.html) and, when FIELD_LEDGER_PASSCODE is set, seals
that page with the same PBKDF2 + AES-GCM gate the REP Payment Sheet uses.

The passcode is read from the environment only. It is never printed and
never written into the page.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
ITERATIONS = 2_000_000

# Figures printed on the workbook Summary tab. The builder refuses to
# write a page unless the payment rows add up to these exactly.
EXPECTED_ALL = (421, 401_678_340)  # count, cents
EXPECTED_CO = {
    "GC": (221, 242_203_413),
    "WD": (200, 159_474_927),
}
EXPECTED_MONTHS = {
    "2026-01": {"GC": 24_761_834, "WD": 10_218_285},
    "2026-02": {"GC": 25_013_310, "WD": 10_774_023},
    "2026-03": {"GC": 25_787_338, "WD": 20_567_429},
    "2026-04": {"GC": 25_499_577, "WD": 11_519_861},
    "2026-05": {"GC": 25_717_371, "WD": 12_125_778},
    "2026-06": {"GC": 22_925_915, "WD": 15_866_305},
    "2026-07": {"GC": 22_996_889, "WD": 13_315_996},
    "2026-08": {"GC": 17_474_882, "WD": 17_324_080},
    "2026-09": {"GC": 17_346_998, "WD": 13_788_027},
    "2026-10": {"GC": 25_254_756, "WD": 17_169_972},
    "2026-11": {"GC": 9_424_543, "WD": 16_805_171},
}

LONG_DIGITS = re.compile(r"\d{7,}")
SENSITIVE = re.compile(
    r"\b\d{3}-\d{2}-\d{4}\b|\b\d{2}-\d{7}\b|\b\d{9}\b"
)


def cents(value) -> int | None:
    if value is None or value == "":
        return None
    quantized = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(quantized * 100)


def text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def received_display(value: str) -> str:
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        parsed = datetime.strptime(value, "%Y-%m-%d")
        return parsed.strftime("%m/%d/%Y")
    parsed = datetime.strptime(value, "%Y-%m-%d %H:%M")
    return parsed.strftime("%m/%d/%Y %H:%M")


def load_rows(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    rows: list[dict] = []
    index = 0
    for sheet in ("GC", "WD"):
        ws = wb[sheet]
        headers = [ws.cell(1, col).value for col in range(1, 11)]
        expected_headers = [
            "Mailbox",
            "Date Received (AZ)",
            "Customer",
            "Invoice #",
            "Invoice Amount",
            "Amount Received",
            "Deposit/Paid Date",
            "Method",
            "Source",
            "Subject",
        ]
        if headers != expected_headers:
            raise SystemExit(f"{sheet} headers changed: {headers}")
        for excel_row in range(2, ws.max_row + 1):
            values = [ws.cell(excel_row, col).value for col in range(1, 11)]
            if all(v is None for v in values):
                continue
            mailbox, received, customer, invoice, invoice_amount, amount, deposit, method, source, subject = values
            if mailbox != sheet:
                raise SystemExit(f"{sheet} row {excel_row} mailbox is {mailbox!r}")
            if not isinstance(received, str):
                raise SystemExit(f"{sheet} row {excel_row} date received is {received!r}")
            amount_cents = cents(amount)
            if amount_cents is None:
                raise SystemExit(f"{sheet} row {excel_row} has no amount received")
            record = {
                "i": index,
                "co": sheet,
                "when": received.replace(" ", "T"),
                "whenText": received_display(received),
                "customer": text(customer),
                "invoice": text(invoice),
                "invoiceCents": cents(invoice_amount),
                "amountCents": amount_cents,
                "deposit": text(deposit),
                "method": text(method),
                "source": text(source),
                "subject": text(subject),
            }
            for field in ("customer", "invoice", "deposit", "method", "source", "subject"):
                blob = record[field]
                check_number = field == "subject" and "Check #" in blob
                if SENSITIVE.search(blob) or (LONG_DIGITS.search(blob) and not check_number):
                    raise SystemExit(
                        f"Refusing to publish {sheet} row {excel_row} {field}: "
                        "possible account, routing, SSN, or EIN"
                    )
            if record["deposit"] and not re.fullmatch(r"\d{2}/\d{2}/\d{4}", record["deposit"]):
                raise SystemExit(f"{sheet} row {excel_row} deposit date is {record['deposit']!r}")
            rows.append(record)
            index += 1
    return rows


def load_csv_rows(gc_path: Path, wd_path: Path) -> list[dict]:
    rows: list[dict] = []
    index = 0
    for sheet, path in (("GC", gc_path), ("WD", wd_path)):
        with path.open(newline="", encoding="utf-8-sig") as handle:
            for excel_row, rec in enumerate(csv.DictReader(handle), start=2):
                values = [rec.get(name) for name in (
                    "Mailbox",
                    "Date Received (AZ)",
                    "Customer",
                    "Invoice #",
                    "Invoice Amount",
                    "Amount Received",
                    "Deposit/Paid Date",
                    "Method",
                    "Source",
                    "Subject",
                )]
                if all(v is None or str(v).strip() == "" for v in values):
                    continue
                mailbox, received, customer, invoice, invoice_amount, amount, deposit, method, source, subject = values
                if text(mailbox) != sheet:
                    raise SystemExit(f"{sheet} row {excel_row} mailbox is {mailbox!r}")
                received = text(received)
                if re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", received):
                    when = received.replace(" ", "T")
                elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", received):
                    when = received
                else:
                    raise SystemExit(f"{sheet} row {excel_row} date received is {received!r}")
                amount_cents = cents(amount)
                if amount_cents is None:
                    raise SystemExit(f"{sheet} row {excel_row} has no amount received")
                record = {
                    "i": index,
                    "co": sheet,
                    "when": when,
                    "whenText": received_display(received),
                    "customer": text(customer),
                    "invoice": text(invoice),
                    "invoiceCents": cents(invoice_amount),
                    "amountCents": amount_cents,
                    "deposit": text(deposit),
                    "method": text(method),
                    "source": text(source),
                    "subject": text(subject),
                    "rep": "Unassigned",
                }
                for field in ("customer", "invoice", "deposit", "method", "source", "subject"):
                    blob = record[field]
                    check_number = field == "subject" and "Check #" in blob
                    if SENSITIVE.search(blob) or (LONG_DIGITS.search(blob) and not check_number):
                        raise SystemExit(
                            f"Refusing to publish {sheet} row {excel_row} {field}: "
                            "possible account, routing, SSN, or EIN"
                        )
                if record["deposit"] and not re.fullmatch(r"\d{2}/\d{2}/\d{4}", record["deposit"]):
                    raise SystemExit(f"{sheet} row {excel_row} deposit date is {record['deposit']!r}")
                rows.append(record)
                index += 1
    return rows


def rep_name(label: str) -> str:
    label = label.strip()
    if re.fullmatch(r"[A-Za-z]+ \([A-Za-z]+\)", label):
        return label
    match = re.match(r"([A-Za-z]+)", label)
    return match.group(1) if match else "Unassigned"


def norm_customer(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())


SUBJECT_REP = re.compile(r"Rep:\s*([A-Za-z]+)\s*\((\d+)\)")
REP_CODES = {
    "222": "Nick",
    "333": "Eric",
    "007": "George",
    "444": "Christian",
    "555": "Daniel",
}


def subject_rep(subject: str) -> str | None:
    match = SUBJECT_REP.search(subject)
    if not match:
        return None
    name, code = match.group(1), match.group(2)
    if REP_CODES.get(code) != name:
        return None
    return name


def load_exclusions(path: Path) -> set[tuple[str, str]]:
    excluded: set[tuple[str, str]] = set()
    if not path.exists():
        return excluded
    with path.open(newline="", encoding="utf-8") as handle:
        for rec in csv.DictReader(handle):
            mailbox = (rec.get("mailbox") or "").strip()
            customer = (rec.get("customer") or "").strip()
            if mailbox not in {"GC", "WD"} or not customer:
                raise SystemExit(f"Bad exclusion row: {rec}")
            excluded.add((mailbox, norm_customer(customer)))
    return excluded


def standalone_token(text: str, token: str) -> bool:
    if not text or not token:
        return False
    if norm_customer(text) == norm_customer(token):
        return True
    pattern = rf"(?:^|[\s\-_.,/()]){re.escape(token)}(?:$|[\s\-_.,/()])"
    return re.search(pattern, text, re.IGNORECASE) is not None


def apply_rep_map(rows: list[dict], map_path: Path) -> None:
    assigned: dict[tuple[str, str], str] = {}
    confirmed: dict[tuple[str, str], str] = {}
    variants: list[tuple[str, str, str]] = []
    excluded = load_exclusions(map_path.parent / "excluded-customers.csv")
    with map_path.open(newline="", encoding="utf-8") as handle:
        for rec in csv.DictReader(handle):
            name = rep_name(rec["rep"])
            customers = [part.strip() for part in (rec.get("report_customers") or "").split("|")]
            mailbox = rec["mailbox"]
            if rec["match_type"] in {"exact", "fuzzy"}:
                for customer in customers:
                    if not customer:
                        continue
                    if (mailbox, norm_customer(customer)) in excluded:
                        raise SystemExit(f"Refusing to assign excluded customer {customer}")
                    key = (mailbox, customer)
                    previous = assigned.get(key)
                    if previous and previous != name:
                        raise SystemExit(f"Rep conflict for {key}: {previous} and {name}")
                    assigned[key] = name
            elif rec["match_type"] == "confirmed":
                if mailbox not in {"GC", "WD"}:
                    raise SystemExit(f"Confirmed rep row has no mailbox: {rec['rep_file_customer']}")
                for customer in [rec["rep_file_customer"], *customers]:
                    if not customer:
                        continue
                    if (mailbox, norm_customer(customer)) in excluded:
                        raise SystemExit(f"Refusing to assign excluded customer {customer}")
                    key = (mailbox, norm_customer(customer))
                    previous = confirmed.get(key)
                    if previous and previous != name:
                        raise SystemExit(f"Rep conflict for {key}: {previous} and {name}")
                    confirmed[key] = name
            elif rec["match_type"] == "variant":
                if mailbox not in {"GC", "WD"}:
                    raise SystemExit(f"Variant rep row has no mailbox: {rec['rep_file_customer']}")
                for token in customers:
                    if token:
                        variants.append((mailbox, token, name))
    variants.sort(key=lambda item: len(item[1]), reverse=True)
    for row in rows:
        exact = assigned.get((row["co"], row["customer"]))
        if exact:
            row["rep"] = exact
        else:
            row["rep"] = confirmed.get((row["co"], norm_customer(row["customer"])), "Unassigned")
        if row["rep"] == "Unassigned":
            chosen = ""
            for mailbox, token, name in variants:
                if mailbox == row["co"] and standalone_token(row["customer"], token):
                    if chosen and chosen != name:
                        raise SystemExit(f"Variant rep conflict for {row['customer']}: {chosen} and {name}")
                    chosen = name
            if chosen:
                row["rep"] = chosen
        if row["rep"] == "Unassigned" and row["co"] == "WD":
            noted = subject_rep(row["subject"])
            if noted:
                row["rep"] = noted
        if row["rep"] == "Unassigned" and APPLY_REP_CODES:
            reps = code_reps(row)
            if len(reps) == 1:
                row["rep"] = reps[0]
        if (row["co"], norm_customer(row["customer"])) in excluded:
            row["rep"] = "Unassigned"


# Explicit name assignments above win. Codes fill only a still-unassigned payment,
# and only when the customer and subject contain one rep code.
APPLY_REP_CODES = True
REP_CODE_RULES = (
    ("333", "Eric"),
    ("222", "Nick"),
    ("007", "George"),
    ("444", "Christian"),
)
_CODE_PATTERN = re.compile(
    r"(?:^|(?<=[\s\-_.,/()]))(333|222|007|444)(?=$|[\s\-_.,/()])"
)
_DIRECTION = re.compile(r"(?i)^(N|S|E|W|NE|NW|SE|SW|North|South|East|West)$")


def code_reps(row: dict) -> list[str]:
    reps = []
    for hit in find_rep_codes(row["customer"]) + find_rep_codes(row["subject"]):
        if hit["rep"] not in reps:
            reps.append(hit["rep"])
    return reps


def find_rep_codes(text: str) -> list[dict]:
    hits = []
    if not text:
        return hits
    names = dict(REP_CODE_RULES)
    for match in _CODE_PATTERN.finditer(text):
        code = match.group(1)
        before = text[: match.start()]
        after = text[match.end() :]
        before_token = re.search(r"([A-Za-z]+)[\s\-_.,/()]*$", before)
        after_token = re.match(r"[\s\-_.,/()]*([A-Za-z]+)", after)
        street = bool(before_token and _DIRECTION.match(before_token.group(1))) or bool(
            after_token and _DIRECTION.match(after_token.group(1))
        )
        hits.append({
            "code": code,
            "rep": names[code],
            "street": street,
            "snippet": text[max(0, match.start() - 18): match.end() + 18],
        })
    return hits


def write_rep_code_review(rows: list[dict], path: Path) -> None:
    review = []
    for row in rows:
        hits = find_rep_codes(row["customer"]) + find_rep_codes(row["subject"])
        if not hits:
            continue
        codes = []
        reps = []
        streets = []
        snippets = []
        for hit in hits:
            if hit["code"] not in codes:
                codes.append(hit["code"])
                reps.append(hit["rep"])
            if hit["street"]:
                streets.append(hit["snippet"])
            snippets.append(hit["snippet"])
        unique_reps = []
        for name in reps:
            if name not in unique_reps:
                unique_reps.append(name)
        proposed = " | ".join(unique_reps)
        flags = []
        if len(unique_reps) > 1:
            flags.append("two_codes")
        if streets:
            flags.append("possible_street")
        disagrees = row["rep"] != "Unassigned" and proposed != row["rep"]
        if disagrees:
            flags.append("disagrees")
        would_assign = row["rep"] == "Unassigned"
        if not (would_assign or disagrees or "two_codes" in flags or "possible_street" in flags):
            continue
        review.append({
            "mailbox": row["co"],
            "date": row["whenText"],
            "customer": row["customer"],
            "amount": f"{row['amountCents'] / 100:.2f}",
            "current_rep": row["rep"],
            "matched_codes": " | ".join(codes),
            "proposed_rep": proposed,
            "flags": "; ".join(flags),
            "evidence": " | ".join(snippets),
        })
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=[
            "mailbox", "date", "customer", "amount", "current_rep",
            "matched_codes", "proposed_rep", "flags", "evidence",
        ])
        writer.writeheader()
        writer.writerows(review)


def apply_branch_deposits(rows: list[dict]) -> None:
    fixes = {
        ("GC", "2026-04-16T11:18", 714500): (
            "Lake Pleasant",
            "Check #118237 (branch deposit)",
        ),
        ("GC", "2026-04-30T09:10", 10524750): (
            "Desert Cove Nursing (paid by Chandler Unified Medical Investors)",
            "Check #28055 (branch deposit)",
        ),
    }
    found = set()
    for row in rows:
        key = (row["co"], row["when"], row["amountCents"])
        if key not in fixes:
            continue
        customer, method = fixes[key]
        if row["customer"] not in {"", customer}:
            raise SystemExit(f"Branch deposit customer is {row['customer']!r}")
        row["customer"] = customer
        row["method"] = method
        for blob in (customer, method):
            if SENSITIVE.search(blob) or LONG_DIGITS.search(blob):
                raise SystemExit(f"Branch deposit correction looks sensitive: {blob}")
        found.add(key)
    missing = set(fixes) - found
    if missing:
        raise SystemExit(f"Branch deposits not found: {missing}")


def verify(rows: list[dict]) -> dict:
    by_co = defaultdict(lambda: {"n": 0, "cents": 0})
    by_month = defaultdict(lambda: {"GC": 0, "WD": 0, "n": 0})
    for row in rows:
        by_co[row["co"]]["n"] += 1
        by_co[row["co"]]["cents"] += row["amountCents"]
        month = row["when"][:7]
        by_month[month][row["co"]] += row["amountCents"]
        by_month[month]["n"] += 1

    all_n = len(rows)
    all_cents = sum(row["amountCents"] for row in rows)
    if (all_n, all_cents) != EXPECTED_ALL:
        raise SystemExit(f"All total {all_n} {all_cents} != {EXPECTED_ALL}")
    for co, expected in EXPECTED_CO.items():
        got = (by_co[co]["n"], by_co[co]["cents"])
        if got != expected:
            raise SystemExit(f"{co} total {got} != {expected}")
    if set(by_month) != set(EXPECTED_MONTHS):
        raise SystemExit(f"Months {sorted(by_month)} != {sorted(EXPECTED_MONTHS)}")
    for month, expected in EXPECTED_MONTHS.items():
        got = {"GC": by_month[month]["GC"], "WD": by_month[month]["WD"]}
        if got != expected:
            raise SystemExit(f"{month} {got} != {expected}")

    return {
        "all": {"n": EXPECTED_ALL[0], "cents": EXPECTED_ALL[1]},
        "GC": {"n": EXPECTED_CO["GC"][0], "cents": EXPECTED_CO["GC"][1]},
        "WD": {"n": EXPECTED_CO["WD"][0], "cents": EXPECTED_CO["WD"][1]},
        "months": EXPECTED_MONTHS,
    }


def js_payload(data) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")


LOGO_SRC = "data:image/png;base64,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"

def money_text(cents_value: int) -> str:
    return f"${cents_value / 100:,.2f}"


def slash_day(iso: str) -> str:
    year, month, day = iso[:10].split("-")
    return f"{month}/{day}/{year}"


def render_report(rows: list[dict], expected: dict) -> str:
    payments = js_payload(rows)
    expect = js_payload(expected)
    whens = [row["when"] for row in rows]
    date_range = f"{slash_day(min(whens))} through {slash_day(max(whens))}"
    check = (
        "Checked against the workbook: "
        f"All {EXPECTED_ALL[0]} / {money_text(EXPECTED_ALL[1])} · "
        f"GC {EXPECTED_CO['GC'][0]} / {money_text(EXPECTED_CO['GC'][1])} · "
        f"WD {EXPECTED_CO['WD'][0]} / {money_text(EXPECTED_CO['WD'][1])}."
    )
    return (
        REPORT_HTML.replace("__PAYMENTS__", payments)
        .replace("__EXPECT__", expect)
        .replace("__LOGO_SRC__", LOGO_SRC)
        .replace("__UPDATED__", "11/18/2026")
        .replace("__RANGE__", date_range)
        .replace("__CHECK_LINE__", check)
    )


REPORT_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Receipts and Reps</title>
<style>
@page { size: letter; margin: 0.45in; }
* { box-sizing: border-box; }
html, body { margin: 0; }
body {
  background: #f3f4f6;
  color: #1c1917;
  font: 14px/1.4 "Segoe UI", system-ui, -apple-system, Roboto, sans-serif;
}
header.top { background: #1f4e79; color: #fff; }
.bar, .wrap { max-width: 1280px; margin: 0 auto; }
.bar {
  display: flex;
  justify-content: space-between;
  gap: 16px;
  align-items: center;
  padding: 14px 22px;
}
.brand { display: flex; align-items: center; gap: 14px; min-width: 0; }
.logo-badge {
  display: flex;
  align-items: center;
  flex-shrink: 0;
  background: #fff;
  border-radius: 10px;
  padding: 4px 10px;
  line-height: 0;
}
.logo-badge img { height: 56px; width: auto; display: block; }
header.top h1 { margin: 0; font-size: 26px; font-weight: 700; letter-spacing: 0; }
.subhead { margin: 4px 0 0; font-size: 14px; opacity: .92; }
.actions { display: flex; gap: 8px; align-items: center; flex-shrink: 0; }
.actions a, header button {
  font: inherit;
  border-radius: 8px;
  padding: 8px 12px;
  cursor: pointer;
  text-decoration: none;
}
.actions a {
  color: #fff;
  border: 1px solid rgba(255,255,255,.75);
}
header button {
  background: #fff;
  color: #1f4e79;
  border: 0;
  font-weight: 600;
}
header button:hover { background: #f5f5f4; }
.wrap { padding: 16px 18px 48px; }
.toolbar {
  position: sticky;
  top: 0;
  z-index: 5;
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 20px rgba(0,0,0,.08);
  padding: 10px 12px;
}
.seg { display: flex; border: 1px solid #1f4e79; border-radius: 8px; overflow: hidden; }
.seg button {
  font: inherit;
  margin: 0;
  border: 0;
  border-right: 1px solid #1f4e79;
  background: #fff;
  color: #1f4e79;
  padding: 8px 14px;
  cursor: pointer;
  font-weight: 600;
}
.seg button:last-child { border-right: 0; }
.seg button[aria-pressed="true"] { background: #1f4e79; color: #fff; }
.seg button:focus-visible, .actions a:focus-visible, header button:focus-visible,
input:focus-visible, select:focus-visible { outline: 2px solid #1f4e79; outline-offset: 2px; }
label.search, label.sort { display: flex; align-items: center; gap: 6px; color: #57534e; font-size: 13px; }
label.search { flex: 1; min-width: 200px; }
input, select {
  font: inherit;
  color: #1c1917;
  background: #fff;
  border: 1px solid #d6d3d1;
  border-radius: 8px;
  padding: 8px 10px;
}
label.search input { width: 100%; }
.stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 14px; }
.stat, .card {
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 20px rgba(0,0,0,.08);
}
.stat { padding: 16px 18px; }
.stat .k {
  font-size: 12px;
  letter-spacing: .04em;
  text-transform: uppercase;
  color: #57534e;
  font-weight: 600;
}
.stat .v {
  margin-top: 4px;
  font-size: 26px;
  font-weight: 700;
  font-variant-numeric: tabular-nums;
  letter-spacing: 0;
}
.stat .s { margin-top: 6px; color: #444; font-size: 13px; }
.split { display: flex; gap: 18px; margin-top: 8px; }
.split .amt { display: block; font-size: 18px; font-weight: 700; font-variant-numeric: tabular-nums; }
.split .s { margin-top: 2px; }
.card { margin-top: 14px; padding: 14px 14px 8px; }
h2 { font-size: 14px; margin: 0 0 8px; color: #1f4e79; }
.scroll { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; margin: 0 0 8px; }
th, td { border: 1px solid #d6d3d1; padding: 6px 8px; vertical-align: top; }
th {
  background: #1f4e79;
  color: #fff;
  font-weight: 600;
  text-align: left;
  white-space: nowrap;
}
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.co { font-weight: 700; color: #1f4e79; white-space: nowrap; }
tr.total td { background: #f5f5f4; font-weight: 700; }
tbody tr:hover td { background: #fafaf9; }
tr.total:hover td, tr.total td { background: #f5f5f4; }
button.drill, button.back {
  font: inherit;
  font-weight: inherit;
  color: #1f4e79;
  background: none;
  border: 0;
  padding: 0;
  cursor: pointer;
  text-align: inherit;
  text-decoration: underline;
  text-underline-offset: 2px;
}
button.back { font-weight: 700; margin: 0 0 10px; }
.subject { min-width: 220px; }
.empty { color: #57534e; padding: 12px 4px; }
.note { margin: 0 0 8px; color: #444; }
#summary-card[hidden], #month-card[hidden], #payment-card[hidden],
#year-filter[hidden], #month-filter[hidden], #rep-filter[hidden], #company-filter[hidden] { display: none; }
.summary-block { margin: 0 0 14px; }
.summary-block h3 { margin: 8px 0 6px; font-size: 13px; color: #1f4e79; }
.banner {
  display: none;
  background: #7f1d1d;
  color: #fff;
  padding: 8px 12px;
  font-weight: 700;
}
.printonly { display: none; }
.check { color: #166534; font-size: 12px; margin: 8px 2px 0; }
@media (max-width: 800px) {
  .bar { flex-direction: column; align-items: flex-start; }
  .logo-badge img { height: 48px; }
  .stats { grid-template-columns: 1fr; }
  header.top h1 { font-size: 22px; }
}
@media print {
  header.top, .noprint, .check { display: none !important; }
  .printonly { display: block; }
  .printonly h1 { font-size: 18px; margin: 0 0 2px; color: #1c1917; }
  .printonly p { margin: 0 0 8px; color: #444; }
  body { background: #fff; }
  .wrap { max-width: none; padding: 0; }
  .stat, .card { box-shadow: none; border-radius: 0; }
  .stats { gap: 0; }
  .stat { border: 1px solid #d6d3d1; }
  a, button.drill { color: inherit; text-decoration: none; }
  table { font-size: 10.5px; }
  th, td { padding: 3px 4px; }
  th { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .card, .stat, tr { break-inside: avoid; }
  thead { display: table-header-group; }
}
</style>
</head>
<body>
<div class="banner" id="banner" role="alert"></div>
<header class="top">
  <div class="bar">
    <div class="brand">
      <span class="logo-badge"><img src="__LOGO_SRC__" alt="Noble" height="56"></span>
      <div>
        <h1>Receipts and Reps</h1>
        <p class="subhead">Updated __UPDATED__ · __RANGE__ (Arizona time)</p>
      </div>
    </div>
    <div class="actions noprint">
      <a href="https://kellynoble26.github.io/noble-rep-payments/index.html" target="_blank" rel="noopener noreferrer">REP Payment Sheet</a>
      <button type="button" id="print">Print</button>
    </div>
  </div>
</header>
<main class="wrap">
  <div class="printonly">
    <h1>Receipts and Reps</h1>
    <p>Noble General Contracting (GC) and Noble Windows &amp; Doors (WD). Updated __UPDATED__. __RANGE__ (Arizona time).</p>
    <p id="printmeta"></p>
  </div>
  <div class="toolbar noprint">
    <div class="seg" role="group" aria-label="Summary">
      <button type="button" data-summary="overview" aria-pressed="true">Summary</button>
    </div>
    <div class="seg" role="group" aria-label="Mailbox">
      <button type="button" data-view="all" aria-pressed="false">All</button>
      <button type="button" data-view="GC" aria-pressed="false">GC only</button>
      <button type="button" data-view="WD" aria-pressed="false">WD only</button>
    </div>
    <div class="seg" role="group" aria-label="Breakdown">
      <button type="button" data-summary="ytd" aria-pressed="false">Year-to-Date</button>
      <button type="button" data-summary="year" aria-pressed="false">Year</button>
      <button type="button" data-summary="month" aria-pressed="false">By Month</button>
      <button type="button" data-summary="rep" aria-pressed="false">By Rep</button>
      <button type="button" data-summary="company" aria-pressed="false">By Company</button>
    </div>
    <label class="sort" id="year-filter" hidden>Year
      <select id="year-pick" aria-label="Year"></select>
    </label>
    <label class="sort" id="month-filter" hidden>Month
      <select id="month-pick" aria-label="Month">
        <option value="all">All</option>
        <option value="1">January</option>
        <option value="2">February</option>
        <option value="3">March</option>
        <option value="4">April</option>
        <option value="5">May</option>
        <option value="6">June</option>
        <option value="7">July</option>
        <option value="8">August</option>
        <option value="9">September</option>
        <option value="10">October</option>
        <option value="11">November</option>
        <option value="12">December</option>
      </select>
    </label>
    <label class="sort" id="rep-filter" hidden>Rep
      <select id="rep-pick" aria-label="Rep">
        <option value="all">All</option>
        <option value="Unassigned">Unassigned</option>
        <option value="George">George</option>
        <option value="Nick">Nick</option>
        <option value="Eric">Eric</option>
        <option value="Eric (George)">Eric (George)</option>
        <option value="Eric (Christian)">Eric (Christian)</option>
        <option value="Daniel">Daniel</option>
        <option value="Christian">Christian</option>
      </select>
    </label>
    <label class="sort" id="company-filter" hidden>Company
      <select id="company-pick" aria-label="Company">
        <option value="all">All</option>
        <option value="GC">Noble General Contracting</option>
        <option value="WD">Noble Windows and Doors</option>
      </select>
    </label>
    <label class="search">Search
      <input id="q" type="search" placeholder="Client or invoice" autocomplete="off" aria-label="Search by client or invoice">
    </label>
    <label class="sort">Sort
      <select id="sort" aria-label="Sort payments">
        <option value="date-asc">Date received, oldest first</option>
        <option value="date-desc">Date received, newest first</option>
        <option value="amount-desc">Amount received, high to low</option>
        <option value="amount-asc">Amount received, low to high</option>
      </select>
    </label>
  </div>
  <section class="stats" id="stats" aria-live="polite"></section>
  <section class="card" id="month-card" hidden>
    <h2>By month</h2>
    <div class="scroll"><table id="months"></table></div>
  </section>
  <section class="card" id="payment-card" hidden>
    <h2>Payments</h2>
    <div class="scroll"><table id="payments"></table></div>
  </section>
  <section class="card" id="summary-card">
    <h2 id="summary-title"></h2>
    <p class="note" id="summary-note"></p>
    <div id="summary-body"></div>
  </section>
  <p class="check noprint" id="check"></p>
</main>
<script>
const PAYMENTS = __PAYMENTS__;
const EXPECT = __EXPECT__;
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const COMPANY = { GC: "Noble General Contracting", WD: "Noble Windows & Doors" };
const VIEW_LABEL = { all: "All", GC: "GC only", WD: "WD only" };

const money = cents => {
  if (cents == null) return "";
  const sign = cents < 0 ? "-" : "";
  const abs = Math.abs(cents);
  const dollars = Math.floor(abs / 100).toLocaleString("en-US");
  const frac = String(abs % 100).padStart(2, "0");
  return sign + "$" + dollars + "." + frac;
};
const monthLabel = ym => MONTHS[Number(ym.slice(5)) - 1] + " " + ym.slice(0, 4);

function verifyData() {
  const problems = [];
  const count = { all: 0, GC: 0, WD: 0 };
  const total = { all: 0, GC: 0, WD: 0 };
  const months = {};
  for (const row of PAYMENTS) {
    count.all += 1;
    count[row.co] += 1;
    total.all += row.amountCents;
    total[row.co] += row.amountCents;
    const ym = row.when.slice(0, 7);
    months[ym] = months[ym] || { GC: 0, WD: 0 };
    months[ym][row.co] += row.amountCents;
  }
  for (const key of ["all", "GC", "WD"]) {
    if (count[key] !== EXPECT[key].n || total[key] !== EXPECT[key].cents) {
      problems.push(key + " " + count[key] + " " + total[key]);
    }
  }
  for (const ym of Object.keys(EXPECT.months)) {
    const got = months[ym] || { GC: 0, WD: 0 };
    if (got.GC !== EXPECT.months[ym].GC || got.WD !== EXPECT.months[ym].WD) problems.push(ym);
  }
  return problems;
}

let view = "all";
let query = "";
let sort = "date-asc";
let summary = "overview";
let drill = null;
const SUMMARY_LABEL = {
  ytd: "Year-to-Date",
  year: "Year",
  month: "By Month",
  rep: "By Rep",
  company: "By Company",
  overview: "Summary"
};
const repOf = row => row.rep || "Unassigned";
const REP_ORDER = ["Nick", "Eric", "Eric (George)", "Eric (Christian)", "George", "Christian", "Daniel", "Unassigned"];
function repGroups(rows) {
  return REP_ORDER.filter(name => rows.some(row => repOf(row) === name)).map(name => {
    const mine = rows.filter(row => repOf(row) === name);
    return { name, n: mine.length, total: sum(mine), key: name };
  });
}

function mailboxScope() {
  if (summary && companyPick !== "all") return companyPick;
  return view;
}

function selected() {
  const q = query.trim().toLowerCase();
  const scope = mailboxScope();
  let rows = PAYMENTS.filter(row => scope === "all" || row.co === scope);
  if (summary === "rep" && repPick !== "all") rows = rows.filter(row => repOf(row) === repPick);
  if (summary === "month" && monthPick !== "all") rows = rows.filter(row => String(Number(row.when.slice(5, 7))) === monthPick);
  if (summary === "year" && yearPick !== "all") rows = rows.filter(row => row.when.slice(0, 4) === yearPick);
  if (q) {
    rows = rows.filter(row =>
      row.customer.toLowerCase().includes(q) || row.invoice.toLowerCase().includes(q)
    );
  }
  rows = rows.slice().sort((a, b) => {
    if (sort === "date-asc" || sort === "date-desc") {
      if (a.when < b.when) return sort === "date-asc" ? -1 : 1;
      if (a.when > b.when) return sort === "date-asc" ? 1 : -1;
      return a.i - b.i;
    }
    if (a.amountCents !== b.amountCents) {
      return sort === "amount-asc" ? a.amountCents - b.amountCents : b.amountCents - a.amountCents;
    }
    if (a.when < b.when) return -1;
    if (a.when > b.when) return 1;
    return a.i - b.i;
  });
  return rows;
}

function sum(rows) { return rows.reduce((n, row) => n + row.amountCents, 0); }

function drillSpec(title, paymentRows) {
  return paymentRows && paymentRows.length ? { title, rows: paymentRows } : null;
}

function openDrill(spec) {
  drill = spec;
  render();
}

function makeTable(headers, lines, emptyText) {
  const wrap = el("div", "scroll");
  const table = el("table");
  const head = el("thead");
  const hr = el("tr");
  headers.forEach((label, i) => hr.append(el("th", i ? "num" : "", label)));
  head.append(hr);
  const body = el("tbody");
  if (!lines.length) {
    const tr = el("tr");
    const td = el("td", "empty", emptyText || "No payments in this view.");
    td.colSpan = headers.length;
    tr.append(td);
    body.append(tr);
  }
  for (const line of lines) {
    const tr = el("tr", line.total ? "total" : "");
    line.cells.forEach((value, i) => {
      const td = el("td", i ? "num" : "");
      const spec = line.drills && line.drills[i];
      if (spec) {
        const button = el("button", "drill", value);
        button.type = "button";
        button.addEventListener("click", () => openDrill(spec));
        td.append(button);
      } else {
        td.textContent = value;
      }
      tr.append(td);
    });
    body.append(tr);
  }
  table.append(head, body);
  wrap.append(table);
  return wrap;
}

function totalLine(label, paymentRows) {
  const spec = drillSpec(label, paymentRows);
  return {
    cells: [label, String(paymentRows.length), money(sum(paymentRows))],
    drills: [spec, spec, spec],
    total: label === "Total"
  };
}

function summaryTable(headers, lines, emptyText) {
  const block = el("div", "summary-block");
  block.append(makeTable(headers, lines, emptyText));
  document.getElementById("summary-body").append(block);
}

function summaryBlock(heading, headers, lines) {
  const block = el("div", "summary-block");
  block.append(el("h3", "", heading), makeTable(headers, lines));
  document.getElementById("summary-body").append(block);
}

function monthKeysOf(rows) {
  return [...new Set(rows.map(row => row.when.slice(0, 7)))].sort();
}

function amountGrid(rowLabels, rowKey, rows) {
  const months = monthKeysOf(rows);
  const headers = [""].concat(months.map(monthLabel), ["Total"]);
  const lines = rowLabels.map(label => {
    const mine = rows.filter(row => rowKey(row) === label.key);
    const cells = [label.name];
    const drills = [drillSpec(label.name, mine)];
    for (const ym of months) {
      const group = mine.filter(row => row.when.slice(0, 7) === ym);
      const amt = sum(group);
      cells.push(amt ? money(amt) : "");
      drills.push(amt ? drillSpec(label.name + " · " + monthLabel(ym), group) : null);
    }
    cells.push(money(sum(mine)));
    drills.push(drillSpec(label.name, mine));
    return { cells, drills };
  });
  if (lines.length) {
    const totalCells = ["Total"];
    const totalDrills = [drillSpec("Total", rows)];
    for (const ym of months) {
      const group = rows.filter(row => row.when.slice(0, 7) === ym);
      totalCells.push(group.length ? money(sum(group)) : "");
      totalDrills.push(drillSpec("Total · " + monthLabel(ym), group));
    }
    totalCells.push(money(sum(rows)));
    totalDrills.push(drillSpec("Total", rows));
    lines.push({ cells: totalCells, drills: totalDrills, total: true });
  }
  return { headers, lines };
}

function renderDrill(spec) {
  const title = document.getElementById("summary-title");
  const note = document.getElementById("summary-note");
  const body = document.getElementById("summary-body");
  body.replaceChildren();
  const listed = spec.rows.slice().sort((a, b) => a.when < b.when ? -1 : a.when > b.when ? 1 : a.i - b.i);
  title.textContent = spec.title;
  note.textContent = listed.length + " payments. Total received " + money(sum(listed)) + ".";
  const back = el("button", "back noprint", "Back to Summary");
  back.type = "button";
  back.id = "back-summary";
  back.addEventListener("click", () => { drill = null; render(); });
  const lines = listed.map(row => ({
    cells: [row.customer, row.whenText, money(row.amountCents), row.co, row.invoice, row.method, row.source]
  }));
  body.append(back, makeTable(
    ["Customer", "Date", "Amount", "GC/WD", "Invoice #", "Method", "Source"],
    lines,
    "No payments in this total."
  ));
}

function renderSummary(rows, total) {
  if (drill) {
    renderDrill(drill);
    return;
  }
  const title = document.getElementById("summary-title");
  const note = document.getElementById("summary-note");
  document.getElementById("summary-body").replaceChildren();
  title.textContent = SUMMARY_LABEL[summary];
  if (summary === "ytd") {
    note.textContent = "__RANGE__ (Arizona time). " + VIEW_LABEL[mailboxScope()] + ".";
    const gc = rows.filter(row => row.co === "GC");
    const wd = rows.filter(row => row.co === "WD");
    const lines = [totalLine(VIEW_LABEL[mailboxScope()], rows)];
    if (mailboxScope() === "all") {
      lines.unshift(totalLine("Noble General Contracting (GC)", gc));
      lines.splice(1, 0, totalLine("Noble Windows & Doors (WD)", wd));
      lines[2].total = true;
    }
    summaryTable(["", "Payments", "Amount received"], lines);
    return;
  }
  if (summary === "year") {
    const yearName = yearPick === "all" ? "All years" : yearPick;
    note.textContent = "Totals for each calendar year in these payments. " + VIEW_LABEL[mailboxScope()] + ". " + yearName + ".";
    const buckets = new Map();
    for (const row of rows) {
      const year = row.when.slice(0, 4);
      if (!buckets.has(year)) buckets.set(year, { n: 0, total: 0 });
      const bucket = buckets.get(year);
      bucket.n += 1;
      bucket.total += row.amountCents;
    }
    const lines = [...buckets.keys()].sort().map(year => totalLine(year, rows.filter(row => row.when.slice(0, 4) === year)));
    if (lines.length) lines.push(totalLine("Total", rows));
    if (!lines.length && yearPick !== "all") note.textContent = "No data yet for this year.";
    summaryTable(["Year", "Payments", "Total"], lines, "No data yet for this year.");
    return;
  }
  if (summary === "month") {
    const monthName = monthPick === "all" ? "All months" : document.getElementById("month-pick").selectedOptions[0].textContent;
    note.textContent = "One row per month for " + VIEW_LABEL[mailboxScope()] + ". " + monthName + ". Count and amount received.";
    const buckets = new Map();
    for (const row of rows) {
      const ym = row.when.slice(0, 7);
      if (!buckets.has(ym)) buckets.set(ym, { n: 0, total: 0 });
      const bucket = buckets.get(ym);
      bucket.n += 1;
      bucket.total += row.amountCents;
    }
    const lines = [...buckets.keys()].sort().map(ym => totalLine(monthLabel(ym), rows.filter(row => row.when.slice(0, 7) === ym)));
    if (lines.length) lines.push(totalLine("Total", rows));
    summaryTable(["Month", "Payments", "Total"], lines);
    return;
  }
  if (summary === "rep") {
    const repName = repPick === "all" ? "All reps" : repPick;
    note.textContent = "Grouped by rep. " + VIEW_LABEL[mailboxScope()] + ". " + repName + ".";
    const lines = repGroups(rows).map(group => totalLine(group.name, rows.filter(row => repOf(row) === group.name)));
    if (lines.length) lines.push(totalLine("Total", rows));
    summaryTable(["Rep", "Payments", "Total"], lines);
    return;
  }
  if (summary === "overview") {
    note.textContent = "Company totals, rep totals, rep by month, and company by month. " + VIEW_LABEL[mailboxScope()] + ".";
    const companyRows = [];
    const scope = mailboxScope();
    if (scope !== "WD") companyRows.push({ key: "GC", name: "Noble General Contracting" });
    if (scope !== "GC") companyRows.push({ key: "WD", name: "Noble Windows and Doors" });
    const companyLines = companyRows.map(item => totalLine(item.name, rows.filter(row => row.co === item.key)));
    if (companyLines.length) companyLines.push(totalLine("Total", rows));
    summaryBlock("Each company", ["Company", "Payments", "Total"], companyLines);
    const groups = repGroups(rows);
    const repLines = groups.map(group => totalLine(group.name, rows.filter(row => repOf(row) === group.name)));
    if (repLines.length) repLines.push(totalLine("Total", rows));
    summaryBlock("Each rep", ["Rep", "Payments", "Total"], repLines);
    const repGrid = amountGrid(groups.map(group => ({ key: group.key, name: group.name })), row => repOf(row), rows);
    summaryBlock("Rep by month", repGrid.headers, repGrid.lines);
    const companyGrid = amountGrid(companyRows, row => row.co, rows);
    summaryBlock("Company by month", companyGrid.headers, companyGrid.lines);
    return;
  }
  const companyName = companyPick === "all" ? "All companies" : document.getElementById("company-pick").selectedOptions[0].textContent;
  note.textContent = "Grouped by the Customer on each payment. " + VIEW_LABEL[mailboxScope()] + ". " + companyName + ".";
  const buckets = new Map();
  for (const row of rows) {
    const name = row.customer;
    if (!buckets.has(name)) buckets.set(name, { n: 0, total: 0 });
    const bucket = buckets.get(name);
    bucket.n += 1;
    bucket.total += row.amountCents;
  }
  const lines = [...buckets.entries()]
    .sort((a, b) => b[1].total - a[1].total || a[0].localeCompare(b[0]))
    .map(([name, bucket]) => totalLine(name, rows.filter(row => row.customer === name)));
  if (lines.length) lines.push(totalLine("Total", rows));
  summaryTable(["Customer", "Payments", "Total"], lines);
}

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text != null) node.textContent = text;
  return node;
}

function render() {
  const rows = selected();
  const total = sum(rows);
  const gc = rows.filter(row => row.co === "GC");
  const wd = rows.filter(row => row.co === "WD");
  document.body.dataset.view = view;
  document.body.dataset.summary = summary;
  document.body.dataset.count = String(rows.length);
  document.body.dataset.totalCents = String(total);
  document.getElementById("month-card").hidden = summary !== "";
  document.getElementById("payment-card").hidden = summary !== "";
  document.getElementById("summary-card").hidden = summary === "";
  document.getElementById("year-filter").hidden = summary !== "year";
  document.getElementById("month-filter").hidden = summary !== "month";
  document.getElementById("rep-filter").hidden = summary !== "rep";
  document.getElementById("company-filter").hidden = !summary;
  if (summary) renderSummary(rows, total);

  const stats = document.getElementById("stats");
  stats.replaceChildren();
  const cards = [
    ["Payments", String(rows.length), VIEW_LABEL[mailboxScope()]],
    ["Total received", money(total), "Amount received"],
  ];
  for (const [k, v, s] of cards) {
    const card = el("article", "stat");
    card.append(el("div", "k", k), el("div", "v", v), el("div", "s", s));
    stats.append(card);
  }
  const mailbox = el("article", "stat");
  if (mailboxScope() === "all") {
    mailbox.append(el("div", "k", "By mailbox"));
    const split = el("div", "split");
    for (const [code, group] of [["GC", gc], ["WD", wd]]) {
      const part = el("div");
      part.append(el("div", "co", code), el("div", "amt", money(sum(group))), el("div", "s", group.length + " payments"));
      split.append(part);
    }
    mailbox.append(split);
  } else {
    mailbox.append(el("div", "k", "Mailbox"), el("div", "v", mailboxScope()), el("div", "s", COMPANY[mailboxScope()]));
  }
  stats.append(mailbox);

  const byMonth = new Map();
  for (const row of rows) {
    const ym = row.when.slice(0, 7);
    if (!byMonth.has(ym)) byMonth.set(ym, { n: 0, GC: 0, WD: 0, total: 0 });
    const bucket = byMonth.get(ym);
    bucket.n += 1;
    bucket[row.co] += row.amountCents;
    bucket.total += row.amountCents;
  }
  const monthKeys = [...byMonth.keys()].sort();
  const months = document.getElementById("months");
  const showSplit = mailboxScope() === "all";
  const head = el("thead");
  const hr = el("tr");
  const headers = showSplit
    ? ["Month", "Payments", "GC", "WD", "Total"]
    : ["Month", "Payments", "Total"];
  headers.forEach((label, i) => hr.append(el("th", i ? "num" : "", label)));
  head.append(hr);
  const body = el("tbody");
  if (!monthKeys.length) {
    const tr = el("tr");
    const td = el("td", "", "No payments in this view.");
    td.colSpan = headers.length;
    tr.append(td);
    body.append(tr);
  }
  for (const ym of monthKeys) {
    const bucket = byMonth.get(ym);
    const tr = el("tr");
    tr.append(el("td", "", monthLabel(ym)), el("td", "num", String(bucket.n)));
    if (showSplit) tr.append(el("td", "num", bucket.GC ? money(bucket.GC) : ""), el("td", "num", bucket.WD ? money(bucket.WD) : ""));
    tr.append(el("td", "num", money(bucket.total)));
    body.append(tr);
  }
  const foot = el("tr", "total");
  foot.append(el("td", "", "Total"), el("td", "num", String(rows.length)));
  if (showSplit) foot.append(el("td", "num", sum(gc) ? money(sum(gc)) : ""), el("td", "num", sum(wd) ? money(sum(wd)) : ""));
  foot.append(el("td", "num", money(total)));
  if (monthKeys.length) body.append(foot);
  months.replaceChildren(head, body);

  const payments = document.getElementById("payments");
  const phead = el("thead");
  const pr = el("tr");
  ["Mailbox", "Date Received (AZ)", "Customer", "Invoice #", "Invoice Amount", "Amount Received", "Deposit/Paid Date", "Method", "Source", "Subject"]
    .forEach((label, i) => pr.append(el("th", (i === 4 || i === 5) ? "num" : "", label)));
  phead.append(pr);
  const pbody = el("tbody");
  if (!rows.length) {
    const tr = el("tr");
    const td = el("td", "empty", "No payments match.");
    td.colSpan = 10;
    tr.append(td);
    pbody.append(tr);
  }
  for (const row of rows) {
    const tr = el("tr");
    tr.append(el("td", "co", row.co));
    tr.append(el("td", "", row.whenText));
    tr.append(el("td", "", row.customer));
    tr.append(el("td", "", row.invoice));
    tr.append(el("td", "num", money(row.invoiceCents)));
    tr.append(el("td", "num", money(row.amountCents)));
    tr.append(el("td", "", row.deposit));
    tr.append(el("td", "", row.method));
    tr.append(el("td", "", row.source));
    tr.append(el("td", "subject", row.subject));
    pbody.append(tr);
  }
  if (rows.length) {
    const totalRow = el("tr", "total");
    const label = el("td", "", "Total");
    label.colSpan = 5;
    totalRow.append(label, el("td", "num", money(total)));
    const rest = el("td", "");
    rest.colSpan = 4;
    totalRow.append(rest);
    pbody.append(totalRow);
  }
  payments.replaceChildren(phead, pbody);

  const sortLabel = document.getElementById("sort").selectedOptions[0].textContent;
  const searchNote = query.trim() ? " Search: " + query.trim() + "." : "";
  const summaryNote = summary ? " Summary: " + SUMMARY_LABEL[summary] + "." : " Sorted by " + sortLabel + ".";
  document.getElementById("printmeta").textContent =
    "View: " + VIEW_LABEL[mailboxScope()] + ". " + rows.length + " payments. Total received " + money(total) + "." + searchNote + summaryNote;
}

function pressOnly(button) {
  document.querySelectorAll("button[data-summary], button[data-view]").forEach(other => {
    other.setAttribute("aria-pressed", other === button ? "true" : "false");
  });
}

function clearPicks() {
  yearPick = "all";
  monthPick = "all";
  repPick = "all";
  companyPick = "all";
  document.getElementById("year-pick").value = "all";
  document.getElementById("month-pick").value = "all";
  document.getElementById("rep-pick").value = "all";
  document.getElementById("company-pick").value = "all";
}

document.querySelectorAll("button[data-view]").forEach(button => {
  button.addEventListener("click", () => {
    drill = null;
    view = button.dataset.view;
    summary = "";
    clearPicks();
    pressOnly(button);
    const hash = view === "all" ? "" : "#" + view.toLowerCase();
    try { history.replaceState(null, "", location.pathname + location.search + hash); } catch (err) {}
    render();
  });
});
const hashView = location.hash.replace("#", "").toLowerCase();
if (hashView === "gc" || hashView === "wd") {
  view = hashView.toUpperCase();
  summary = "";
  pressOnly(document.querySelector('button[data-view="' + view + '"]'));
}
const params = new URLSearchParams(location.search);
if (params.get("q")) {
  query = params.get("q");
  document.getElementById("q").value = query;
}
if (["date-asc", "date-desc", "amount-desc", "amount-asc"].includes(params.get("sort"))) {
  sort = params.get("sort");
  document.getElementById("sort").value = sort;
}
if (["ytd", "year", "month", "rep", "company", "overview"].includes(params.get("summary"))) {
  summary = params.get("summary");
  view = "all";
  pressOnly(document.querySelector('button[data-summary="' + summary + '"]'));
}
document.querySelectorAll("button[data-summary]").forEach(button => {
  button.addEventListener("click", () => {
    drill = null;
    summary = button.dataset.summary;
    view = "all";
    clearPicks();
    pressOnly(button);
    try { history.replaceState(null, "", location.pathname + location.search); } catch (err) {}
    render();
  });
});
let yearPick = "all";
let monthPick = "all";
let repPick = "all";
let companyPick = "all";
const yearSelect = document.getElementById("year-pick");
yearSelect.append(el("option", "", "All"));
yearSelect.lastChild.value = "all";
["2024", "2025", "2026", "2027"].forEach(year => {
  const option = el("option", "", year);
  option.value = year;
  yearSelect.append(option);
});
yearSelect.addEventListener("change", event => { drill = null; yearPick = event.target.value; render(); });
document.getElementById("month-pick").addEventListener("change", event => { drill = null; monthPick = event.target.value; render(); });
document.getElementById("rep-pick").addEventListener("change", event => { drill = null; repPick = event.target.value; render(); });
document.getElementById("company-pick").addEventListener("change", event => { drill = null; companyPick = event.target.value; render(); });
document.getElementById("q").addEventListener("input", event => { drill = null; query = event.target.value; render(); });
document.getElementById("sort").addEventListener("change", event => { sort = event.target.value; render(); });
document.getElementById("print").addEventListener("click", () => window.print());

const problems = verifyData();
const banner = document.getElementById("banner");
if (problems.length) {
  banner.style.display = "block";
  banner.textContent = "Totals do not match the workbook Summary tab.";
  document.getElementById("check").textContent = "";
} else {
  document.getElementById("check").textContent =
    "__CHECK_LINE__";
}
render();
</script>
</body>
</html>
"""


GATE_HTML = r"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"><title>Monies Received 2026</title>
<style>html,body{margin:0;height:100%}body{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:#f3f4f6;display:flex;min-height:100vh;align-items:center;justify-content:center}
.box{background:#fff;padding:32px 28px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.08);width:300px;text-align:center}
h1{font-size:20px;margin:0 0 18px;color:#111}input{width:100%;box-sizing:border-box;font-size:22px;letter-spacing:8px;text-align:center;padding:10px;border:1px solid #ccc;border-radius:8px}
button{margin-top:14px;width:100%;padding:11px;font-size:16px;border:0;border-radius:8px;background:#1f4e79;color:#fff;cursor:pointer}button:disabled{opacity:.6}
#msg{min-height:20px;margin-top:10px;font-size:14px;color:#b91c1c}
.alt{display:inline-block;margin-top:14px;color:#1f4e79;font-size:14px}
.switch{position:fixed;top:0;left:0;right:0;height:36px;display:flex;align-items:center;justify-content:flex-end;gap:14px;padding:0 14px;background:#1f4e79;color:#fff;font:13px/1 system-ui,Segoe UI,sans-serif;z-index:20}
.switch a{color:#fff}
iframe.ledger{border:0;position:fixed;inset:0;width:100%;height:100%;background:#fff;z-index:10}</style></head>
<body><form class="box" id="f" autocomplete="off"><h1>Monies Received 2026</h1>
<input id="p" type="password" inputmode="numeric" maxlength="4" placeholder="Passcode" autocomplete="off" autofocus aria-label="Passcode">
<button id="b" type="submit">Open</button><div id="msg"></div>
<a class="alt" href="https://kellynoble26.github.io/noble-rep-payments/index.html" target="_blank" rel="noopener noreferrer">REP Payment Sheet</a></form>
<script>
const S="__SALT__",V="__IV__",I=__ITERATIONS__,D="__DATA__";
const u=s=>Uint8Array.from(atob(s),ch=>ch.charCodeAt(0));
document.getElementById('f').addEventListener('submit',async e=>{e.preventDefault();
const p=document.getElementById('p'),b=document.getElementById('b'),m=document.getElementById('msg');
b.disabled=true;m.style.color='#555';m.textContent='Unlocking…';
try{const k0=await crypto.subtle.importKey('raw',new TextEncoder().encode(p.value),'PBKDF2',false,['deriveKey']);
const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt:u(S),iterations:I,hash:'SHA-256'},k0,{name:'AES-GCM',length:256},false,['decrypt']);
const pt=await crypto.subtle.decrypt({name:'AES-GCM',iv:u(V)},k,u(D));
const h=new TextDecoder().decode(pt);
const url=URL.createObjectURL(new Blob([h],{type:'text/html;charset=utf-8'}));
const frame=document.createElement('iframe');
frame.className='ledger';frame.title='Monies Received';frame.src=url;
frame.style.top='36px';frame.style.height='calc(100% - 36px)';
const bar=document.createElement('div');bar.className='switch';
bar.innerHTML='<a href="https://kellynoble26.github.io/noble-rep-payments/index.html" target="_blank" rel="noopener noreferrer">REP Payment Sheet</a>';
document.body.replaceChildren(bar, frame);
}catch(err){m.style.color='#b91c1c';m.textContent='Incorrect passcode';b.disabled=false;p.value='';p.focus();}});
</script></body></html>
"""


def seal(inner_html: str, password: str) -> str:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    salt = os.urandom(16)
    iv = os.urandom(12)
    key = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=ITERATIONS,
    ).derive(password.encode())
    ciphertext = AESGCM(key).encrypt(iv, inner_html.encode(), None)
    data = base64.b64encode(ciphertext).decode()
    # Keep the ciphertext off one enormous line so GitHub Pages can serve it.
    chunked = "\\\n".join(data[i : i + 76] for i in range(0, len(data), 76))
    page = GATE_HTML
    page = page.replace("__SALT__", base64.b64encode(salt).decode(), 1)
    page = page.replace("__IV__", base64.b64encode(iv).decode(), 1)
    page = page.replace("__ITERATIONS__", str(ITERATIONS), 1)
    page = page.replace("__DATA__", chunked, 1)
    head, _, _tail = page.partition('D="')
    if f'"{password}"' in head or f"'{password}'" in head:
        raise SystemExit("Refusing to write a page that contains the passcode")
    return page


def self_test_crypto() -> None:
    """Round-trip the gate with a dummy password. Does not use Kelly's passcode."""
    sample = "<!doctype html><title>ok</title>"
    sealed = seal(sample, "1357")
    salt_b64 = sealed.split('S="', 1)[1].split('"', 1)[0]
    iv_b64 = sealed.split('V="', 1)[1].split('"', 1)[0]
    data_b64 = "".join(sealed.split('D="', 1)[1].split('";', 1)[0].split("\\\n"))
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    key = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=base64.b64decode(salt_b64),
        iterations=ITERATIONS,
    ).derive(b"1357")
    plain = AESGCM(key).decrypt(base64.b64decode(iv_b64), base64.b64decode(data_b64), None)
    if plain.decode() != sample:
        raise SystemExit("crypto self-test failed")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--gc", type=Path, required=False, help="GC payments CSV")
    parser.add_argument("--wd", type=Path, required=False, help="WD payments CSV")
    parser.add_argument("--reps", type=Path, default=ROOT / "data" / "rep-customer-map.csv")
    parser.add_argument("--self-test", action="store_true", help="Check the seal round-trip and exit")
    args = parser.parse_args()
    if args.self_test:
        self_test_crypto()
        print("crypto self-test ok")
        return
    if not args.gc or not args.wd:
        parser.error("--gc and --wd payment CSVs are required")
    rows = load_csv_rows(args.gc, args.wd)
    apply_branch_deposits(rows)
    apply_rep_map(rows, args.reps)
    write_rep_code_review(rows, ROOT / "data" / "rep-code-review.csv")
    expected = verify(rows)
    html = render_report(rows, expected)
    out = ROOT / "received.html"
    password = os.environ.get("FIELD_LEDGER_PASSCODE", "")
    if password:
        out.write_text(seal(html, password), encoding="utf-8")
        print(f"sealed {out} ({len(rows)} payments)")
    else:
        out.write_text(html, encoding="utf-8")
        print(f"wrote unlocked review {out} ({len(rows)} payments)")
        print("FIELD_LEDGER_PASSCODE is not set, so this copy is not sealed.")
    print(
        f"All {EXPECTED_ALL[0]} ${EXPECTED_ALL[1]/100:,.2f} | "
        f"GC {EXPECTED_CO['GC'][0]} ${EXPECTED_CO['GC'][1]/100:,.2f} | "
        f"WD {EXPECTED_CO['WD'][0]} ${EXPECTED_CO['WD'][1]/100:,.2f}"
    )


if __name__ == "__main__":
    try:
        main()
    except BrokenPipeError:
        sys.exit(0)
