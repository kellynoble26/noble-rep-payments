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
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
ITERATIONS = 2_000_000

# Figures printed on the workbook Summary tab. The builder refuses to
# write a page unless the payment rows add up to these exactly.
EXPECTED_ALL = (390, 356_925_832)  # count, cents
EXPECTED_CO = {
    "GC": (223, 245_384_381),
    "WD": (167, 111_541_451),
}
EXPECTED_MONTHS = {
    "2026-01": {"GC": 24_761_834, "WD": 2_477_696},
    "2026-02": {"GC": 25_013_310, "WD": 9_363_080},
    "2026-03": {"GC": 25_787_338, "WD": 13_820_897},
    "2026-04": {"GC": 25_499_577, "WD": 9_400_636},
    "2026-05": {"GC": 25_717_371, "WD": 8_382_682},
    "2026-06": {"GC": 25_589_563, "WD": 11_692_609},
    "2026-07": {"GC": 22_996_889, "WD": 6_584_383},
    "2026-08": {"GC": 17_474_882, "WD": 15_875_985},
    "2026-09": {"GC": 17_346_998, "WD": 10_840_375},
    "2026-10": {"GC": 25_772_076, "WD": 11_291_922},
    "2026-11": {"GC": 9_424_543, "WD": 11_811_186},
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
                if LONG_DIGITS.search(blob) or SENSITIVE.search(blob):
                    raise SystemExit(
                        f"Refusing to publish {sheet} row {excel_row} {field}: "
                        "possible account, routing, SSN, or EIN"
                    )
            if record["deposit"] and not re.fullmatch(r"\d{2}/\d{2}/\d{4}", record["deposit"]):
                raise SystemExit(f"{sheet} row {excel_row} deposit date is {record['deposit']!r}")
            rows.append(record)
            index += 1
    return rows


def verify(rows: list[dict], workbook: Path) -> dict:
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

    wb = load_workbook(workbook, data_only=True)
    summary = wb["Summary"]
    summary_all = (
        int(summary["B7"].value),
        cents(summary["C7"].value),
    )
    if summary_all != EXPECTED_ALL:
        raise SystemExit(f"Summary All cell {summary_all} != {EXPECTED_ALL}")
    for cell, co in (("B5", "GC"), ("B6", "WD")):
        count = int(summary[cell].value)
        total = cents(summary[cell.replace("B", "C")].value)
        if (count, total) != EXPECTED_CO[co]:
            raise SystemExit(f"Summary {co} {(count, total)} != {EXPECTED_CO[co]}")

    return {
        "all": {"n": EXPECTED_ALL[0], "cents": EXPECTED_ALL[1]},
        "GC": {"n": EXPECTED_CO["GC"][0], "cents": EXPECTED_CO["GC"][1]},
        "WD": {"n": EXPECTED_CO["WD"][0], "cents": EXPECTED_CO["WD"][1]},
        "months": EXPECTED_MONTHS,
    }


def js_payload(data) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")


def render_report(rows: list[dict], expected: dict) -> str:
    payments = js_payload(rows)
    expect = js_payload(expected)
    return REPORT_HTML.replace("__PAYMENTS__", payments).replace("__EXPECT__", expect)


REPORT_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Monies Received 2026</title>
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
  align-items: flex-end;
  padding: 18px 22px 16px;
}
.eyebrow {
  font-size: 12px;
  letter-spacing: .06em;
  text-transform: uppercase;
  opacity: .85;
}
header.top h1 { margin: 2px 0 0; font-size: 26px; font-weight: 700; letter-spacing: 0; }
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
.subject { min-width: 220px; }
.empty { color: #57534e; padding: 12px 4px; }
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
  a { color: inherit; text-decoration: none; }
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
    <div>
      <div class="eyebrow">Noble General Contracting · Noble Windows &amp; Doors</div>
      <h1>Monies Received</h1>
      <p class="subhead">Updated 11/17/2026 · 01/01/2026 through 11/17/2026 (Arizona time)</p>
    </div>
    <div class="actions noprint">
      <a href="index.html">REP Payment Sheet</a>
      <button type="button" id="print">Print</button>
    </div>
  </div>
</header>
<main class="wrap">
  <div class="printonly">
    <h1>Monies Received</h1>
    <p>Noble General Contracting (GC) and Noble Windows &amp; Doors (WD). Updated 11/17/2026. 01/01/2026 through 11/17/2026 (Arizona time).</p>
    <p id="printmeta"></p>
  </div>
  <div class="toolbar noprint">
    <div class="seg" role="group" aria-label="Mailbox">
      <button type="button" data-view="all" aria-pressed="true">All</button>
      <button type="button" data-view="GC" aria-pressed="false">GC only</button>
      <button type="button" data-view="WD" aria-pressed="false">WD only</button>
    </div>
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
  <section class="card">
    <h2>By month</h2>
    <div class="scroll"><table id="months"></table></div>
  </section>
  <section class="card">
    <h2>Payments</h2>
    <div class="scroll"><table id="payments"></table></div>
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

function selected() {
  const q = query.trim().toLowerCase();
  let rows = PAYMENTS.filter(row => view === "all" || row.co === view);
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
  document.body.dataset.count = String(rows.length);
  document.body.dataset.totalCents = String(total);

  const stats = document.getElementById("stats");
  stats.replaceChildren();
  const cards = [
    ["Payments", String(rows.length), VIEW_LABEL[view]],
    ["Total received", money(total), "Amount received"],
  ];
  for (const [k, v, s] of cards) {
    const card = el("article", "stat");
    card.append(el("div", "k", k), el("div", "v", v), el("div", "s", s));
    stats.append(card);
  }
  const mailbox = el("article", "stat");
  if (view === "all") {
    mailbox.append(el("div", "k", "By mailbox"));
    const split = el("div", "split");
    for (const [code, group] of [["GC", gc], ["WD", wd]]) {
      const part = el("div");
      part.append(el("div", "co", code), el("div", "amt", money(sum(group))), el("div", "s", group.length + " payments"));
      split.append(part);
    }
    mailbox.append(split);
  } else {
    mailbox.append(el("div", "k", "Mailbox"), el("div", "v", view), el("div", "s", COMPANY[view]));
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
  const showSplit = view === "all";
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
  document.getElementById("printmeta").textContent =
    "View: " + VIEW_LABEL[view] + ". " + rows.length + " payments. Total received " + money(total) + "." + searchNote + " Sorted by " + sortLabel + ".";
}

document.querySelectorAll(".seg button").forEach(button => {
  button.addEventListener("click", () => {
    view = button.dataset.view;
    document.querySelectorAll(".seg button").forEach(other => {
      other.setAttribute("aria-pressed", other === button ? "true" : "false");
    });
    const hash = view === "all" ? "" : "#" + view.toLowerCase();
    try { history.replaceState(null, "", location.pathname + location.search + hash); } catch (err) {}
    render();
  });
});
const hashView = location.hash.replace("#", "").toLowerCase();
if (hashView === "gc" || hashView === "wd") {
  view = hashView.toUpperCase();
  document.querySelectorAll(".seg button").forEach(button => {
    button.setAttribute("aria-pressed", button.dataset.view === view ? "true" : "false");
  });
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
document.getElementById("q").addEventListener("input", event => { query = event.target.value; render(); });
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
    "Checked against the workbook: All 390 / $3,569,258.32 · GC 223 / $2,453,843.81 · WD 167 / $1,115,414.51.";
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
<a class="alt" href="index.html">REP Payment Sheet</a></form>
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
bar.innerHTML='<a href="index.html">REP Payment Sheet</a>';
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
    parser.add_argument("workbook", type=Path, nargs="?", help="Monies_Received_2026.xlsx")
    parser.add_argument("--self-test", action="store_true", help="Check the seal round-trip and exit")
    args = parser.parse_args()
    if args.self_test:
        self_test_crypto()
        print("crypto self-test ok")
        return
    if not args.workbook:
        parser.error("workbook path is required")
    rows = load_rows(args.workbook)
    expected = verify(rows, args.workbook)
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
