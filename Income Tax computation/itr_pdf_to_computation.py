#!/usr/bin/env python3
"""
Convert an Income Tax portal ITR PDF into a Word-openable Computation of Income.

The generated file is HTML saved with a .doc extension, matching the style of
many computation documents that open cleanly in Microsoft Word.
"""

from __future__ import annotations

import argparse
import html
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

try:
    from pypdf import PdfReader
except ImportError as exc:  # pragma: no cover - helpful runtime message
    raise SystemExit("Missing dependency: install pypdf with `pip install pypdf`.") from exc

logging.getLogger("pypdf").setLevel(logging.ERROR)


def money_to_int(value: str | None) -> int:
    if not value:
        return 0
    cleaned = re.sub(r"[^\d\-]", "", value)
    if cleaned in {"", "-"}:
        return 0
    return int(cleaned)


def fmt_money(value: int | str | None) -> str:
    if isinstance(value, str):
        value = money_to_int(value)
    value = int(value or 0)
    sign = "-" if value < 0 else ""
    value = abs(value)
    digits = str(value)
    if len(digits) <= 3:
        return sign + digits
    last3 = digits[-3:]
    rest = digits[:-3]
    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        parts.insert(0, rest)
    return sign + ",".join(parts + [last3])


def safe(text: object) -> str:
    return html.escape("" if text is None else str(text), quote=True)


def normalize_lines(text: str) -> list[str]:
    text = text.replace("\u00a0", " ").replace("\u2002", " ").replace("�", "")
    lines = []
    for line in text.splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            lines.append(line)
    return lines


def extract_pdf_text(pdf_path: Path) -> tuple[str, list[str]]:
    reader = PdfReader(str(pdf_path))
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n".join(pages)
    return text, normalize_lines(text)


def next_line_after(lines: list[str], marker: str, skip: Iterable[str] = ()) -> str:
    skip_set = {s.lower() for s in skip}
    for index, line in enumerate(lines):
        if line == marker:
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
    return ""


def next_value_after(lines: list[str], marker: str, skip: Iterable[str] = ()) -> str:
    skip_set = {s.lower() for s in skip}
    for index, line in enumerate(lines):
        if line == marker:
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
    return ""


def value_after_code(lines: list[str], code: str) -> str:
    for index, line in enumerate(lines):
        if line == code:
            for candidate in lines[index + 1 :]:
                if candidate == code:
                    continue
                if re.fullmatch(r"-?[\d,]+", candidate):
                    return candidate
                if candidate == "0":
                    return candidate
                # Stop when another labelled field begins.
                if re.fullmatch(r"[A-Z]\d+[a-z]?", candidate):
                    break
    return "0"


def value_after_label(lines: list[str], label: str) -> str:
    for index, line in enumerate(lines):
        if line == label and index + 1 < len(lines):
            return lines[index + 1]
    return ""


def collect_between(lines: list[str], start: str, end: str) -> list[str]:
    try:
        start_index = lines.index(start)
    except ValueError:
        return []
    try:
        end_index = lines.index(end, start_index + 1)
    except ValueError:
        end_index = len(lines)
    return lines[start_index + 1 : end_index]


def last_money_between(lines: list[str], start: str, end: str) -> int:
    block = collect_between(lines, start, end)
    values = [money_to_int(line) for line in block if re.fullmatch(r"-?[\d,]+", line)]
    return values[-1] if values else 0


@dataclass
class Business:
    name: str = ""
    code: str = ""
    description: str = ""
    gross_receipt: int = 0
    banking_receipt: int = 0
    cash_receipt: int = 0
    other_receipt: int = 0
    income_6: int = 0
    income_8: int = 0
    income_total: int = 0


@dataclass
class TaxData:
    assessment_year: str = ""
    first_name: str = ""
    middle_name: str = ""
    last_name: str = ""
    full_name: str = ""
    pan: str = ""
    dob: str = ""
    status: str = ""
    address: str = ""
    city: str = ""
    state: str = ""
    pin: str = ""
    email: str = ""
    father_name: str = ""
    place: str = ""
    filing_date: str = ""
    acknowledgement: str = ""
    business_income: int = 0
    salary_income: int = 0
    house_property_income: int = 0
    other_sources_income: int = 0
    gross_total_income: int = 0
    total_deductions: int = 0
    total_income: int = 0
    tax_on_income: int = 0
    rebate: int = 0
    cess: int = 0
    total_tax_cess: int = 0
    interest_234a: int = 0
    interest_234b: int = 0
    interest_234c: int = 0
    fee_234f: int = 0
    total_tax_fee_interest: int = 0
    taxes_paid: int = 0
    amount_payable: int = 0
    refund: int = 0
    bank_ifsc: str = ""
    bank_name: str = ""
    bank_account: str = ""
    bank_type: str = ""
    other_income_description: str = "Other Income"
    business: Business = field(default_factory=Business)


def parse_itr_pdf(pdf_path: Path) -> TaxData:
    text, lines = extract_pdf_text(pdf_path)
    data = TaxData()

    data.assessment_year = next_line_after(lines, "Year") or value_after_label(lines, "Assessment Year")
    data.first_name = next_line_after(lines, "(A1) First Name")
    data.middle_name = next_line_after(lines, "(A2) Middle Name")
    data.last_name = next_line_after(lines, "(A3) Last Name")
    data.full_name = " ".join(x for x in [data.first_name, data.middle_name, data.last_name] if x).strip()
    data.pan = next_line_after(lines, "(A4) Permanent Account Number")
    data.dob = next_line_after(lines, "(A5) Date of Birth/Formation (DD/MM/YYYY)")
    data.status = next_line_after(lines, "(A15) Status")
    data.email = next_line_after(lines, "(A18a) Primary Email ID of the taxpayer")
    data.acknowledgement = re_search(text, r"Acknowledgement Number\s*:\s*([0-9]+)")
    data.filing_date = re_search(text, r"Date of Filing\s*:\s*([0-9A-Za-z\-]+)")

    flat = next_line_after(lines, "(A6a) Flat/Door/Block No.")
    premise = next_value_after(lines, "(A7a) Name of", skip=("Premises/Building/Village",))
    road = next_line_after(lines, "(A8a) Road/Street/Post Office")
    locality = next_line_after(lines, "(A9a) Area/Locality")
    data.city = next_line_after(lines, "(A10a) Town/City/District")
    data.state = re.sub(r"^\d+\-", "", next_line_after(lines, "(A11a) State")).strip()
    data.pin = next_line_after(lines, "(A13a) PIN Code/ZIP Code")
    data.address = " ".join(x for x in [flat, premise, road, locality] if x).strip()

    data.father_name = re_search(text, r"son/ daughter of\s*\n?([A-Z][A-Z ]+)")
    data.place = re_search(text, r"Place:\s*([A-Z ]+)")

    data.business_income = money_to_int(value_after_code(lines, "B1"))
    data.salary_income = money_to_int(value_after_code(lines, "v"))
    data.house_property_income = money_to_int(value_after_code(lines, "B3"))
    data.other_sources_income = money_to_int(value_after_code(lines, "B4"))
    data.gross_total_income = money_to_int(value_after_code(lines, "B5"))
    data.total_deductions = money_to_int(value_after_code(lines, "C19"))
    data.total_income = last_money_between(lines, "C20", "PART D - TAX COMPUTATIONS AND TAX STATUS") or money_to_int(value_after_code(lines, "C20"))
    data.tax_on_income = money_to_int(value_after_code(lines, "D1"))
    data.rebate = money_to_int(value_after_code(lines, "D2"))
    data.cess = money_to_int(value_after_code(lines, "D4"))
    data.total_tax_cess = money_to_int(value_after_code(lines, "D5"))
    data.interest_234a = money_to_int(value_after_code(lines, "D8"))
    data.interest_234b = money_to_int(value_after_code(lines, "D9"))
    data.interest_234c = money_to_int(value_after_code(lines, "D10"))
    data.fee_234f = money_to_int(value_after_code(lines, "D11"))
    data.total_tax_fee_interest = money_to_int(value_after_code(lines, "D12"))
    data.taxes_paid = money_to_int(value_after_code(lines, "D17"))
    data.amount_payable = money_to_int(value_after_code(lines, "D18"))
    data.refund = money_to_int(value_after_code(lines, "D19"))

    other_block = collect_between(lines, "Nature of Income", "Please provide Quarterly breakup of Dividend Income")
    for idx, line in enumerate(other_block):
        if line == "Any Other" and idx + 2 < len(other_block):
            data.other_income_description = other_block[idx + 1].title()
            break

    business = Business()
    bp = collect_between(lines, "COMPUTATION OF PRESUMPTIVE BUSINESS INCOME UNDER SECTION 44AD", "COMPUTATION OF PRESUMPTIVE INCOME FROM PROFESSIONS UNDER SECTION 44ADA")
    business.name = find_business_name(bp)
    code_desc = next((line for line in bp if re.match(r"^\d{5}\-", line)), "")
    if code_desc:
        business.code, business.description = code_desc.split("-", 1)
    # The next non-code business text after code is commonly the trade description.
    if business.name and not business.description:
        business.description = business.name
    business.gross_receipt = money_to_int(value_after_code(lines, "E1"))
    business.banking_receipt = money_to_int(value_after_code(lines, "E1a"))
    business.cash_receipt = money_to_int(value_after_code(lines, "E1b"))
    business.other_receipt = money_to_int(value_after_code(lines, "E1c"))
    business.income_6 = money_to_int(value_after_code(lines, "E2a"))
    business.income_8 = money_to_int(value_after_code(lines, "E2b"))
    business.income_total = money_to_int(value_after_code(lines, "E2c")) or data.business_income
    data.business = business

    data.bank_ifsc, data.bank_name, data.bank_account, data.bank_type = parse_bank(lines)

    return data


def re_search(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    return re.sub(r"\s+", " ", match.group(1)).strip() if match else ""


def find_business_name(bp_lines: list[str]) -> str:
    for index, line in enumerate(bp_lines):
        if line == "1":
            pieces = []
            for candidate in bp_lines[index + 1 :]:
                if re.match(r"^\d{5}\-", candidate):
                    break
                if candidate not in {"(1)", "(2)", "(3)", "(4)"}:
                    pieces.append(candidate)
            return " ".join(pieces).strip()
    return ""


def parse_bank(lines: list[str]) -> tuple[str, str, str, str]:
    block = collect_between(lines, "(D21) DETAILS OF ALL BANK ACCOUNT DETAILS HELD IN INDIA AT ANY TIME DURING THE PREVIOUS YEAR (EXCLUDING DORMANT", "SCHEDULE BP - DETAILS OF INCOME FROM BUSINESS OR PROFESSION")
    for index, line in enumerate(block):
        if re.fullmatch(r"[A-Z]{4}0[A-Z0-9]{6}", line) and index + 3 < len(block):
            return line, block[index + 1], block[index + 2], block[index + 3]
    return "", "", "", ""


def compute_new_regime_tax(total_income: int) -> int:
    slabs = [
        (400000, 0.00),
        (400000, 0.05),
        (400000, 0.10),
        (400000, 0.15),
        (400000, 0.20),
        (10**12, 0.30),
    ]
    remaining = max(0, total_income)
    tax = 0.0
    for width, rate in slabs:
        taxable = min(remaining, width)
        tax += taxable * rate
        remaining -= taxable
        if remaining <= 0:
            break
    return round(tax)


def compute_old_regime_tax(total_income: int) -> int:
    slabs = [(250000, 0.00), (250000, 0.05), (500000, 0.20), (10**12, 0.30)]
    remaining = max(0, total_income)
    tax = 0.0
    for width, rate in slabs:
        taxable = min(remaining, width)
        tax += taxable * rate
        remaining -= taxable
        if remaining <= 0:
            break
    return round(tax)


def new_regime_slab_rows(total_income: int) -> list[tuple[str, str, int]]:
    slabs = [
        ("0 to 4 lakh", "0%", 400000, 0.00),
        ("4 lakh to 8 lakh", "5%", 400000, 0.05),
        ("8 lakh to 12 lakh", "10%", 400000, 0.10),
        ("12 lakh to 16 lakh", "15%", 400000, 0.15),
        ("16 lakh to 20 lakh", "20%", 400000, 0.20),
        ("Above 20 lakh", "30%", 10**12, 0.30),
    ]
    remaining = max(0, total_income)
    rows = []
    for label, rate_label, width, rate in slabs:
        taxable = min(remaining, width)
        amount = round(taxable * rate)
        if taxable > 0 or rows:
            rows.append((label, rate_label, amount))
        remaining -= taxable
        if remaining <= 0:
            break
    return rows


def output_name(data: TaxData, pdf_path: Path) -> str:
    ay = data.assessment_year.replace("/", "-") or "AY"
    name = re.sub(r"[^A-Za-z0-9]+", " ", data.full_name).strip().replace(" ", " ")
    name = re.sub(r"\s+", " ", name).strip()
    stem = f"AY{ay} {name}-{data.pan}-Computation".strip(" -")
    stem = re.sub(r'[<>:"/\\|?*]', "", stem)
    return stem + ".doc"


def render_doc(data: TaxData) -> str:
    tax_after_rebate = max(0, data.tax_on_income - data.rebate)
    old_tax = compute_old_regime_tax(data.total_income)
    old_cess = round(old_tax * 0.04)
    old_total = old_tax + old_cess
    new_slab_rows = new_regime_slab_rows(data.total_income)

    income_rows = [
        ("Business and Profession", data.business_income),
        ("Salary", data.salary_income),
        ("House Property", data.house_property_income),
        ("Other Sources", data.other_sources_income),
    ]
    income_rows = [row for row in income_rows if row[1] != 0 or row[0] in {"Business and Profession", "Other Sources"}]

    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="ProgId" content="Word.Document">
<title>{safe(data.full_name)} Computation</title>
<style>
@page {{ size: A4; margin: .5in; }}
body {{ font-family: Cambria, "Times New Roman", serif; font-size: 12pt; color: #111; }}
h1, h2, h3 {{ margin: 0; text-align: center; }}
h1 {{ font-size: 18pt; }}
h2 {{ font-size: 13.5pt; margin-top: 10pt; }}
h3 {{ font-size: 12pt; margin: 12pt 0 6pt; }}
p {{ margin: 0 0 5pt; }}
table {{ border-collapse: collapse; width: 100%; margin: 5pt 0 10pt; }}
td, th {{ padding: 4pt 6pt; vertical-align: top; }}
.meta td {{ border: none; padding: 2pt 6pt; }}
.box td, .box th {{ border: 1px solid #999; }}
.label {{ font-weight: bold; }}
.amount {{ text-align: right; white-space: nowrap; }}
.center {{ text-align: center; }}
.total td {{ border-top: 1.5pt solid #000; font-weight: bold; }}
.section {{ background: #eee; font-weight: bold; text-align: center; }}
.small {{ font-size: 10pt; }}
.signature {{ margin-top: 35pt; text-align: right; }}
</style>
</head>
<body>
<h1>{safe(data.full_name)}</h1>
<h2>AY {safe(data.assessment_year)}</h2>
<table class="meta">
<tr><td class="label">Address:</td><td colspan="3">{safe(data.address)}<br>{safe(data.city)}, {safe(data.state)} - {safe(data.pin)}</td></tr>
<tr><td class="label">E-Mail:</td><td>{safe(data.email)}</td><td class="label">PAN:</td><td>{safe(data.pan)}</td></tr>
<tr><td class="label">Status:</td><td>{safe(data.status)}</td><td class="label">Date of Birth:</td><td>{safe(data.dob)}</td></tr>
<tr><td class="label">Residential Status:</td><td>Resident</td><td class="label">Father's Name:</td><td>{safe(data.father_name)}</td></tr>
<tr><td class="label">Bank A/C no.:</td><td>{safe(data.bank_account)}</td><td class="label">IFSC code:</td><td>{safe(data.bank_ifsc)}</td></tr>
<tr><td class="label">E-Filing Status:</td><td>Filed</td><td class="label">Selected tax regime:</td><td>New Regime</td></tr>
<tr><td class="label">Acknowledgement:</td><td>{safe(data.acknowledgement)}</td><td class="label">Date of Filing:</td><td>{safe(data.filing_date)}</td></tr>
</table>

<h2>Computation of Income (ITR4)</h2>
<h3>Tax Summary (Amount in Rs.)</h3>
<table class="box">
{''.join(f'<tr><td>{safe(label)}</td><td class="amount">{fmt_money(amount)}</td></tr>' for label, amount in income_rows)}
<tr class="total"><td>Gross Total Income</td><td class="amount">{fmt_money(data.gross_total_income)}</td></tr>
<tr><td>Less: Total Deductions</td><td class="amount">{fmt_money(data.total_deductions)}</td></tr>
<tr class="total"><td>Total Income (Taxable)<br><span class="small">Rounded off as per Section 288A</span></td><td class="amount">{fmt_money(data.total_income)}</td></tr>
</table>
<p>Taxes are applicable as per normal provision. Please refer Annexure for details.</p>

<h3>Business and Profession</h3>
<table class="box">
<tr><th>Particulars</th><th class="amount">Amount</th></tr>
<tr><td>Presumptive Income u/s 44AD</td><td class="amount">{fmt_money(data.business.income_total)}</td></tr>
<tr class="total"><td>Net Income under the head "Business and Profession"</td><td class="amount">{fmt_money(data.business_income)}</td></tr>
</table>

<table class="box">
<tr><th>Business nature</th><th>Business code</th><th>Trade Name</th></tr>
<tr><td>{safe(data.business.description)}</td><td class="center">{safe(data.business.code)}</td><td>{safe(data.business.name)}</td></tr>
</table>

<table class="box">
<tr><th>Particulars</th><th class="amount">Cash Transactions (8%)</th><th class="amount">Any Other Mode Transactions (8%)</th><th class="amount">Banking Mode Transactions (6%)</th><th class="amount">Total</th></tr>
<tr><td>Gross Receipt</td><td class="amount">{fmt_money(data.business.cash_receipt)}</td><td class="amount">{fmt_money(data.business.other_receipt)}</td><td class="amount">{fmt_money(data.business.banking_receipt)}</td><td class="amount">{fmt_money(data.business.gross_receipt)}</td></tr>
<tr><td>Income u/s 44AD</td><td class="amount">{fmt_money(0 if data.business.cash_receipt == 0 else data.business.income_8)}</td><td class="amount">{fmt_money(data.business.income_8)}</td><td class="amount">{fmt_money(data.business.income_6)}</td><td class="amount">{fmt_money(data.business.income_total)}</td></tr>
</table>

<h3>Other Income</h3>
<table class="box">
<tr><td>{safe(data.other_income_description)}</td><td class="amount">{fmt_money(data.other_sources_income)}</td></tr>
<tr class="total"><td>Total</td><td class="amount">{fmt_money(data.other_sources_income)}</td></tr>
</table>

<h3>Income Tax</h3>
<table class="box">
<tr><td>Total Income</td><td class="amount">{fmt_money(data.total_income)}</td></tr>
<tr><td>Basic Exemption</td><td class="amount">4,00,000</td></tr>
<tr><td>Income Tax</td><td class="amount">{fmt_money(data.tax_on_income)}</td></tr>
<tr><td>Rebate u/s 87A</td><td class="amount">{fmt_money(data.rebate)}</td></tr>
<tr><td>Health and Education Cess</td><td class="amount">{fmt_money(data.cess)}</td></tr>
<tr class="total"><td>Tax after rebate</td><td class="amount">{fmt_money(tax_after_rebate + data.cess)}</td></tr>
<tr class="total"><td>Payable</td><td class="amount">{fmt_money(data.amount_payable)}</td></tr>
</table>

<h3>Normal Tax Breakup</h3>
<table class="box">
<tr><th>Income Slab</th><th>Rate</th><th class="amount">Tax Amount</th></tr>
{''.join(f'<tr><td>{safe(label)}</td><td class="center">{safe(rate)}</td><td class="amount">{fmt_money(amount)}</td></tr>' for label, rate, amount in new_slab_rows)}
<tr class="total"><td colspan="2">Total</td><td class="amount">{fmt_money(data.tax_on_income)}</td></tr>
</table>

<h2>Annexures</h2>
<h3>Tax Computation Comparison - New Regime vs Old Regime</h3>
<p class="small">This computation has been prepared under the <b>New Tax Regime</b>. The table below shows a basic comparison using extracted income figures.</p>
<table class="box">
<tr><th>Particulars</th><th class="amount">New Regime (Rs.)</th><th class="amount">Old Regime (Rs.)</th></tr>
<tr><td>Business &amp; Profession</td><td class="amount">{fmt_money(data.business_income)}</td><td class="amount">{fmt_money(data.business_income)}</td></tr>
<tr><td>Income From Other Sources</td><td class="amount">{fmt_money(data.other_sources_income)}</td><td class="amount">{fmt_money(data.other_sources_income)}</td></tr>
<tr class="total"><td>Gross Total Income</td><td class="amount">{fmt_money(data.gross_total_income)}</td><td class="amount">{fmt_money(data.gross_total_income)}</td></tr>
<tr class="total"><td>Total Income</td><td class="amount">{fmt_money(data.total_income)}</td><td class="amount">{fmt_money(data.total_income)}</td></tr>
<tr><td>Income Tax at normal rates</td><td class="amount">{fmt_money(data.tax_on_income)}</td><td class="amount">{fmt_money(old_tax)}</td></tr>
<tr><td>Rebate u/s 87A</td><td class="amount">-{fmt_money(data.rebate)}</td><td class="amount">-</td></tr>
<tr><td>Health and Education Cess</td><td class="amount">{fmt_money(data.cess)}</td><td class="amount">{fmt_money(old_cess)}</td></tr>
<tr class="total"><td>Tax Due</td><td class="amount">{fmt_money(data.total_tax_cess)}</td><td class="amount">{fmt_money(old_total)}</td></tr>
<tr class="total"><td>Tax Payable</td><td class="amount">{fmt_money(data.amount_payable)}</td><td class="amount">{fmt_money(old_total)}</td></tr>
</table>

<p class="small"><b>Note:</b> Under the new tax regime, many Chapter VI-A deductions such as 80C, 80D, 80E, 80G, and 80CCD(1B) are generally not available, except where specifically allowed by law.</p>

<h3>Bank Account Details</h3>
<table class="box">
<tr><th>SI No.</th><th>IFSC Code</th><th>Name of the Bank</th><th>Account No.</th><th>Type</th></tr>
<tr><td class="center">1</td><td>{safe(data.bank_ifsc)}</td><td>{safe(data.bank_name)}</td><td>{safe(data.bank_account)}</td><td>{safe(data.bank_type)}</td></tr>
</table>

<div class="signature">
<p>Signature</p>
<p>For {safe(data.full_name)}</p>
</div>
</body>
</html>
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Income Tax portal ITR PDF to Computation of Income Word document.")
    parser.add_argument("pdf", type=Path, help="Input PDF downloaded from the income tax portal")
    parser.add_argument("-o", "--output", type=Path, help="Output .doc path. Defaults beside the PDF name in the current folder.")
    parser.add_argument("--json", action="store_true", help="Print extracted fields as JSON instead of writing a document")
    args = parser.parse_args()

    data = parse_itr_pdf(args.pdf)
    if args.json:
        import json
        from dataclasses import asdict

        print(json.dumps(asdict(data), indent=2, ensure_ascii=False))
        return 0

    output = args.output or Path.cwd() / output_name(data, args.pdf)
    output.write_text(render_doc(data), encoding="utf-8")
    print(f"Wrote {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
