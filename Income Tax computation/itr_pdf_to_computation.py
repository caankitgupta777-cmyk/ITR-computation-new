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
        # Exact match
        if line == marker:
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
        # Starts with match, e.g. "(A2) First Name SANKALP"
        if line.startswith(marker):
            val = line[len(marker):].strip()
            if val and val.lower() not in skip_set:
                return val
            # Check next lines
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
        # Substring match if it contains the marker (e.g. spaces/slashes mismatch)
        if marker in line:
            idx = line.find(marker)
            val = line[idx + len(marker):].strip()
            if val and val.lower() not in skip_set:
                return val
            # Check next lines
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
    return ""


def next_value_after(lines: list[str], marker: str, skip: Iterable[str] = ()) -> str:
    skip_set = {s.lower() for s in skip}
    for index, line in enumerate(lines):
        if line == marker:
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
        if line.startswith(marker):
            val = line[len(marker):].strip()
            if val and val.lower() not in skip_set:
                return val
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
        if marker in line:
            idx = line.find(marker)
            val = line[idx + len(marker):].strip()
            if val and val.lower() not in skip_set:
                return val
            for candidate in lines[index + 1 :]:
                if candidate.lower() not in skip_set:
                    return candidate
            break
    return ""


def has_code(lines: list[str], code: str) -> bool:
    pattern = r"\b" + re.escape(code) + r"\b"
    for line in lines:
        if line.strip() == code or re.search(pattern, line):
            return True
    return False


def value_after_code(lines: list[str], code: str) -> str:
    # 1. Exact match
    for index, line in enumerate(lines):
        if line.strip() == code:
            for candidate in lines[index + 1 :]:
                if candidate.strip() == code:
                    continue
                if re.fullmatch(r"[A-Z]\d+[a-z]?", candidate):
                    break
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return candidate
            break

    # 2. Substring search with word boundary
    pattern = r"\b" + re.escape(code) + r"\b"
    for index, line in enumerate(lines):
        if re.search(pattern, line):
            m = re.findall(r"-?[\d,]+", line)
            if m:
                cleaned_line = line.strip()
                end_match = re.search(r"-?[\d,]+$", cleaned_line)
                if end_match:
                    return end_match.group(0)
            for candidate in lines[index + 1 :]:
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return candidate
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


def find_value_by_label_exact(lines: list[str], label: str) -> int:
    label_lower = label.strip().lower()
    for idx, line in enumerate(lines):
        if line.strip().lower() == label_lower:
            for candidate in lines[idx + 1 : idx + 6]:
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return money_to_int(candidate)
    return 0


def find_value_by_label_substring(lines: list[str], label_sub: str, exclude_sub: str = None) -> int:
    label_sub_lower = label_sub.lower()
    exclude_lower = exclude_sub.lower() if exclude_sub else None
    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if label_sub_lower in line_lower:
            if exclude_lower and exclude_lower in line_lower:
                continue
            for candidate in lines[idx + 1 : idx + 6]:
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return money_to_int(candidate)
    return 0


def find_refund_value(lines: list[str]) -> int:
    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if line_lower.startswith("refund") and "interest" not in line_lower:
            for candidate in lines[idx + 1 : idx + 6]:
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return money_to_int(candidate)
    return 0


def find_salary_income(lines: list[str]) -> int:
    for idx, line in enumerate(lines):
        if "income chargeable under the head" in line.lower() and "salaries" in line.lower():
            # Check if there is a number at the end of the same line
            cleaned = line.strip()
            end_match = re.search(r"-?[\d,]+$", cleaned)
            if end_match:
                return money_to_int(end_match.group(0))
            
            # If not, check subsequent lines
            for candidate in lines[idx + 1 : idx + 6]:
                if re.fullmatch(r"[A-Z]\d+[a-z]?", candidate):
                    break
                if re.fullmatch(r"-?[\d,]+", candidate) or candidate == "0":
                    return money_to_int(candidate)
    return money_to_int(value_after_code(lines, "v"))


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
    itr_form: str = "ITR-4"
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


def next_email_after_marker(lines: list[str], marker: str) -> str:
    for index, line in enumerate(lines):
        if marker in line:
            m = re.search(r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b", line)
            if m:
                return m.group(0)
            for candidate in lines[index + 1 : index + 5]:
                m = re.search(r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b", candidate)
                if m:
                    return m.group(0)
    return ""


def next_date_after_marker(lines: list[str], marker: str) -> str:
    for index, line in enumerate(lines):
        if marker in line:
            m = re.search(r"\b\d{2}/\d{2}/\d{4}\b", line)
            if m:
                return m.group(0)
            for candidate in lines[index + 1 : index + 4]:
                m = re.search(r"\b\d{2}/\d{2}/\d{4}\b", candidate)
                if m:
                    return m.group(0)
    return ""


def extract_address_field(line: str, marker: str) -> str:
    idx = line.find(marker)
    if idx == -1:
        return ""
    line_clean = line[idx + len(marker):].strip()
    labels = [
        r"Flat/Door/Block\s+No\b\.?",
        r"Name\s+of\s+Premises\s*/\s*Building\s*/\s*Village",
        r"Name\s+of\s+Premises\b",
        r"Road/Street/Post\s+Office,\s+Area/Locality",
        r"Road/Street/Post\s+Office\b",
        r"Area/Locality\b",
        r"Town/City/District\b",
        r"State\b",
        r"PIN\s+Code/ZIP\s+Code",
        r"PIN\s+Code\b",
    ]
    for label in labels:
        line_clean = re.sub(r"^" + label + r"\s*", "", line_clean, flags=re.IGNORECASE).strip()
    # Clean leading punctuation like commas or slashes
    line_clean = re.sub(r"^[\s,\-\./]+", "", line_clean).strip()
    return line_clean


def next_address_line_after(lines: list[str], marker: str) -> str:
    for index, line in enumerate(lines):
        if marker in line:
            val = extract_address_field(line, marker)
            if val and len(val) > 2:
                return val
            if index + 1 < len(lines):
                next_line = lines[index + 1].strip()
                if not re.match(r"^\([A-Z]\d+[a-z]?\)", next_line):
                    return next_line
            break
    return ""


def parse_itr_pdf(pdf_path: Path) -> TaxData:
    text, lines = extract_pdf_text(pdf_path)
    data = TaxData()

    # Detect Form Type
    text_upper = text.upper()
    if "ITR-1" in text_upper or "ITR1" in text_upper or "SAHAJ" in text_upper:
        data.itr_form = "ITR-1"
    else:
        data.itr_form = "ITR-4"

    data.assessment_year = next_line_after(lines, "Year") or value_after_label(lines, "Assessment Year")
    
    if data.itr_form == "ITR-1":
        data.first_name = next_line_after(lines, "(A2) First Name")
        data.middle_name = next_line_after(lines, "(A2a) Middle Name")
        data.last_name = next_line_after(lines, "(A3) Last Name")
    else:
        data.first_name = next_line_after(lines, "(A1) First Name")
        data.middle_name = next_line_after(lines, "(A2) Middle Name")
        data.last_name = next_line_after(lines, "(A3) Last Name")
        
    data.full_name = " ".join(x for x in [data.first_name, data.middle_name, data.last_name] if x).strip()
    
    if data.itr_form == "ITR-1":
        data.pan = next_line_after(lines, "(A1) PAN")
        data.dob = next_date_after_marker(lines, "(A4) Date of Birth")
        data.status = next_line_after(lines, "(A15) Status") or "Individual"
        data.email = (
            next_email_after_marker(lines, "(A7)(a) Primary Email ID") or
            next_email_after_marker(lines, "Primary Email ID") or
            next_line_after(lines, "(A7) Email Address")
        )
    else:
        data.pan = next_line_after(lines, "(A4) Permanent Account Number")
        data.dob = next_line_after(lines, "(A5) Date of Birth/Formation (DD/MM/YYYY)")
        data.status = next_line_after(lines, "(A15) Status")
        data.email = next_line_after(lines, "(A18a) Primary Email ID of the taxpayer")
        
    data.acknowledgement = re_search(text, r"Acknowledgement Number\s*:\s*([0-9]+)")
    data.filing_date = re_search(text, r"Date of Filing\s*:\s*([0-9A-Za-z\-]+)")

    if data.itr_form == "ITR-1":
        flat = next_address_line_after(lines, "(A8a)") or next_address_line_after(lines, "(A8)")
        premise = next_address_line_after(lines, "(A9a)") or next_address_line_after(lines, "(A9)")
        road = next_address_line_after(lines, "(A10a)") or next_address_line_after(lines, "(A10)")
        data.city = next_address_line_after(lines, "(A11a)") or next_address_line_after(lines, "(A11)")
        data.state = re.sub(r"^\d+\-", "", next_address_line_after(lines, "(A12a)") or next_address_line_after(lines, "(A12)")).strip()
        data.pin = next_address_line_after(lines, "(A14a)") or next_address_line_after(lines, "(A14)")
        data.address = " ".join(x for x in [flat, premise, road] if x).strip()
    else:
        flat = next_line_after(lines, "(A6a) Flat/Door/Block No.")
        premise = next_value_after(lines, "(A7a) Name of", skip=("Premises/Building/Village",))
        road = next_line_after(lines, "(A8a) Road/Street/Post Office")
        locality = next_line_after(lines, "(A9a) Area/Locality")
        data.city = next_line_after(lines, "(A10a) Town/City/District")
        data.state = re.sub(r"^\d+\-", "", next_line_after(lines, "(A11a) State")).strip()
        data.pin = next_line_after(lines, "(A13a) PIN Code/ZIP Code")
        data.address = " ".join(x for x in [flat, premise, road, locality] if x).strip()

    father_match = re.search(r"son/\s*daughter\s*of\s*([^,\n\.\(]+)", text, re.IGNORECASE)
    if father_match:
        father_name = father_match.group(1).strip()
        father_name = re.split(r"\b(solemnly|declare|do)\b", father_name, flags=re.IGNORECASE)[0].strip()
        data.father_name = father_name
    else:
        data.father_name = re_search(text, r"son/ daughter of\s*\n?([A-Z][A-Z ]+)")

    data.place = re_search(text, r"Place:\s*([A-Za-z0-9\.\- ]+)")

    # Income & Deductions & Taxes
    data.salary_income = find_salary_income(lines)

    if data.itr_form == "ITR-1":
        data.business_income = 0
        data.house_property_income = money_to_int(value_after_code(lines, "B2"))
        data.other_sources_income = money_to_int(value_after_code(lines, "B3"))
        data.gross_total_income = money_to_int(value_after_code(lines, "B4"))
        
        c1_val = value_after_code(lines, "C1")
        if c1_val != "0" or has_code(lines, "C1"):
            data.total_deductions = money_to_int(c1_val)
        else:
            data.total_deductions = money_to_int(value_after_code(lines, "C21")) or money_to_int(value_after_code(lines, "C22"))
    else:
        data.business_income = money_to_int(value_after_code(lines, "B1"))
        data.house_property_income = money_to_int(value_after_code(lines, "B3"))
        data.other_sources_income = money_to_int(value_after_code(lines, "B4"))
        data.gross_total_income = money_to_int(value_after_code(lines, "B5"))
        data.total_deductions = money_to_int(value_after_code(lines, "C19"))

    # Taxable Total Income u/s 288A
    data.total_income = find_value_by_label_exact(lines, "Total Income")
    if data.total_income == 0 or data.total_income == data.total_deductions:
        if data.itr_form == "ITR-1":
            data.total_income = money_to_int(value_after_code(lines, "C2")) or money_to_int(value_after_code(lines, "C22")) or money_to_int(value_after_code(lines, "C23"))
        else:
            data.total_income = money_to_int(value_after_code(lines, "C20"))
            
    if data.total_income == 0:
        data.total_income = max(0, data.gross_total_income - data.total_deductions)
        
    data.total_income = round(data.total_income / 10) * 10

    # Tax Calculations
    data.tax_on_income = find_value_by_label_substring(lines, "Tax payable on total income") or money_to_int(value_after_code(lines, "D1"))
    data.rebate = find_value_by_label_substring(lines, "Rebate u/s 87A") or money_to_int(value_after_code(lines, "D2"))
    data.cess = find_value_by_label_substring(lines, "education Cess") or money_to_int(value_after_code(lines, "D4"))
    data.total_tax_cess = find_value_by_label_substring(lines, "Total Tax and Cess") or money_to_int(value_after_code(lines, "D5"))
    
    if data.itr_form == "ITR-1":
        data.interest_234a = money_to_int(value_after_code(lines, "D7")) or find_value_by_label_substring(lines, "Interest u/s 234A")
        data.interest_234b = money_to_int(value_after_code(lines, "D8")) or find_value_by_label_substring(lines, "Interest u/s 234B")
        data.interest_234c = money_to_int(value_after_code(lines, "D9")) or find_value_by_label_substring(lines, "Interest u/s 234C")
        data.fee_234f = money_to_int(value_after_code(lines, "D10")) or find_value_by_label_substring(lines, "Fee u/s 234F") or find_value_by_label_substring(lines, "Late Fee")
    else:
        data.interest_234a = money_to_int(value_after_code(lines, "D8")) or find_value_by_label_substring(lines, "Interest u/s 234A")
        data.interest_234b = money_to_int(value_after_code(lines, "D9")) or find_value_by_label_substring(lines, "Interest u/s 234B")
        data.interest_234c = money_to_int(value_after_code(lines, "D10")) or find_value_by_label_substring(lines, "Interest u/s 234C")
        data.fee_234f = money_to_int(value_after_code(lines, "D11")) or find_value_by_label_substring(lines, "Fee u/s 234F")

    data.total_tax_fee_interest = find_value_by_label_substring(lines, "Total Tax, Fee and Interest") or find_value_by_label_substring(lines, "Total Interest and Fee Payable")
    if data.total_tax_fee_interest == 0:
        if data.itr_form == "ITR-1":
            data.total_tax_fee_interest = money_to_int(value_after_code(lines, "D11")) or money_to_int(value_after_code(lines, "D12"))
        else:
            data.total_tax_fee_interest = money_to_int(value_after_code(lines, "D12"))

    data.taxes_paid = find_value_by_label_substring(lines, "Total Taxes Paid") or find_value_by_label_substring(lines, "Total Tax Paid")
    if data.taxes_paid == 0:
        if data.itr_form == "ITR-1":
            data.taxes_paid = money_to_int(value_after_code(lines, "D12")) or money_to_int(value_after_code(lines, "D18"))
        else:
            data.taxes_paid = money_to_int(value_after_code(lines, "D17"))

    data.amount_payable = find_value_by_label_substring(lines, "Amount payable")
    if data.amount_payable == 0:
        if data.itr_form == "ITR-1":
            data.amount_payable = money_to_int(value_after_code(lines, "D13")) or money_to_int(value_after_code(lines, "D19"))
        else:
            data.amount_payable = money_to_int(value_after_code(lines, "D18"))

    data.refund = find_refund_value(lines)
    if data.refund == 0:
        if data.itr_form == "ITR-1":
            data.refund = money_to_int(value_after_code(lines, "D14")) or money_to_int(value_after_code(lines, "D20"))
        else:
            data.refund = money_to_int(value_after_code(lines, "D19"))

    other_block = []
    start_idx = -1
    for idx, line in enumerate(lines):
        if "Nature of Income" in line:
            start_idx = idx
            break
    if start_idx != -1:
        end_idx = len(lines)
        for idx in range(start_idx + 1, len(lines)):
            if "Quarterly breakup of Dividend" in lines[idx] or "Dividend Income" in lines[idx]:
                end_idx = idx
                break
        other_block = lines[start_idx + 1 : end_idx]
    
    for idx, line in enumerate(other_block):
        if "Any Other" in line:
            for candidate in other_block[idx + 1 : idx + 4]:
                if candidate and not re.fullmatch(r"-?[\d,]+", candidate) and not re.match(r"^\d+$", candidate):
                    data.other_income_description = candidate.title()
                    break
            break

    business = Business()
    if data.itr_form == "ITR-4":
        bp = collect_between(lines, "COMPUTATION OF PRESUMPTIVE BUSINESS INCOME UNDER SECTION 44AD", "COMPUTATION OF PRESUMPTIVE INCOME FROM PROFESSIONS UNDER SECTION 44ADA")
        business.name = find_business_name(bp)
        code_desc = next((line for line in bp if re.match(r"^\d{5}\-", line)), "")
        if code_desc:
            business.code, business.description = code_desc.split("-", 1)
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
    for index, line in enumerate(lines):
        line_clean = line.strip()
        ifsc_match = re.search(r"\b([A-Z]{4}0[A-Z0-9]{6})\b", line_clean, re.IGNORECASE)
        if ifsc_match:
            # 1. Check if the line ONLY contains the IFSC. If so, bank name/acct/type are on subsequent lines.
            if re.fullmatch(r"[A-Z]{4}0[A-Z0-9]{6}", line_clean, re.IGNORECASE) and index + 3 < len(lines):
                return line_clean.upper(), lines[index + 1].strip(), lines[index + 2].strip(), lines[index + 3].strip()
            
            # 2. Otherwise, parse inline using the IFSC and Account Number as anchors
            ifsc = ifsc_match.group(1).upper()
            # Look for account number (a digit block of length 9-18)
            acct_match = re.search(r"\b(\d{9,18})\b", line_clean)
            if acct_match:
                acct_num = acct_match.group(1)
                # Bank name is between IFSC and Account Number
                ifsc_end = ifsc_match.end()
                acct_start = acct_match.start()
                bank_name = line_clean[ifsc_end:acct_start].strip()
                # Account type is after Account Number
                acct_end = acct_match.end()
                acct_type = line_clean[acct_end:].strip()
                # Clean up any checkmarks, labels, or trailing noise
                acct_type = re.sub(r"\b(Select Account|Refund Credit|o|Tick|Yes|No|[\u2611\u2610])\b", "", acct_type, flags=re.IGNORECASE).strip()
                acct_type = re.sub(r"\s+", " ", acct_type).strip()
                return ifsc, bank_name, acct_num, acct_type
            
            # 3. Fallback to space split if account number was not a simple digit block
            parts = [p for p in re.split(r"\s{2,}", line_clean) if p.strip()]
            if len(parts) < 2:
                parts = line_clean.split(" ")
            for idx, part in enumerate(parts):
                if re.fullmatch(r"[A-Z]{4}0[A-Z0-9]{6}", part, re.IGNORECASE):
                    remaining = parts[idx:]
                    if len(remaining) >= 4:
                        return remaining[0].upper(), remaining[1].strip(), remaining[2].strip(), remaining[3].strip()
                    elif len(remaining) == 3:
                        return remaining[0].upper(), remaining[1].strip(), remaining[2].strip(), ""
                    elif len(remaining) == 2:
                        return remaining[0].upper(), remaining[1].strip(), "", ""
                    return remaining[0].upper(), "", "", ""
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

    # Dynamic Financial Year extraction from Assessment Year
    ay = data.assessment_year or ""
    fy = ""
    match = re.search(r"(\d{4})[-/](\d{2,4})", ay)
    if match:
        start_yr = int(match.group(1))
        fy_start = start_yr - 1
        fy_end = start_yr
        fy = f"{fy_start}-{str(fy_end)[-2:]}"
    else:
        fy = "Preceding F.Y."

    # Build Income Heads Multi-column rows
    income_rows_html = []
    # Business & Profession
    if data.business_income != 0:
        income_rows_html.append(f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 500;">Profits & Gains of Business or Profession</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center; font-size: 8.5pt; color: #475569;">Sec 44AD</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.business_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold; color: #0f172a;">{fmt_money(data.business_income)}</td>
    </tr>""")
    # Salary
    if data.salary_income != 0:
        income_rows_html.append(f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 500;">Income under the head "Salaries"</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center; font-size: 8.5pt; color: #475569;">Sec 15-17</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.salary_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold; color: #0f172a;">{fmt_money(data.salary_income)}</td>
    </tr>""")
    # House Property
    if data.house_property_income != 0:
        income_rows_html.append(f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 500;">Income / (Loss) from House Property</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center; font-size: 8.5pt; color: #475569;">Sec 22-27</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.house_property_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold; color: #0f172a;">{fmt_money(data.house_property_income)}</td>
    </tr>""")
    # Other Sources
    if data.other_sources_income != 0:
        income_rows_html.append(f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 500;">Income from Other Sources <span style="font-size: 8pt; font-weight: normal; color: #64748b;">({safe(data.other_income_description)})</span></td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center; font-size: 8.5pt; color: #475569;">Sec 56</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.other_sources_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold; color: #0f172a;">{fmt_money(data.other_sources_income)}</td>
    </tr>""")

    # Fallback if no income fields parsed to ensure table isn't blank
    if not income_rows_html:
        income_rows_html.append(f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; color: #94a3b8; font-style: italic;">No specific income heads reported</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center; color: #94a3b8;">-</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #94a3b8;">0</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #94a3b8;">0</td>
    </tr>""")

    income_rows_html_str = "".join(income_rows_html)

    # Balance row html (Payable / Refund)
    if data.refund > 0:
        balance_row_html = f"""
    <tr style="font-weight: bold; background-color: #ecfdf5; border-top: 1.5pt solid #10b981; border-bottom: 3px double #10b981;">
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; color: #047857; font-size: 10.5pt;">NET REFUND DUE TO TAXPAYER</td>
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: center; color: #047857; font-size: 8.5pt;">Sec 244A</td>
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right; color: #059669; font-size: 11.5pt;">{fmt_money(data.refund)}</td>
    </tr>"""
    else:
        balance_row_html = f"""
    <tr style="font-weight: bold; background-color: #fef2f2; border-top: 1.5pt solid #ef4444; border-bottom: 3px double #ef4444;">
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; color: #b91c1c; font-size: 10.5pt;">NET TAX PAYABLE / (REMAINING LIABILITY)</td>
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: center; color: #b91c1c; font-size: 8.5pt;">-</td>
      <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right; color: #dc2626; font-size: 11.5pt;">{fmt_money(data.amount_payable)}</td>
    </tr>"""

    # Slab Rows HTML
    slab_rows_html = "".join(f"""
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; color: #334155;">{safe(label)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; color: #475569;">{safe(rate)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(amount)}</td>
    </tr>""" for label, rate, amount in new_slab_rows)

    # Optional Schedules
    business_schedule_html = ""
    if data.business_income != 0 or data.business.gross_receipt != 0:
        business_schedule_html = f"""
<h3 style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11.5pt; font-weight: bold; color: #1e3a8a; margin: 10pt 0in 4pt; mso-margin-top-alt: 10pt; mso-margin-bottom-alt: 4pt; border-bottom: 2px solid #cbd5e1; padding-bottom: 4px; text-transform: uppercase;">SCHEDULE BP: PRESUMPTIVE BUSINESS INCOME (SECTION 44AD)</h3>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 10pt; font-family: 'Segoe UI', Calibri, Arial, sans-serif; border: 1px solid #cbd5e1;">
  <tr style="background-color: #f8fafc; font-weight: bold; color: #334155;">
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; width: 33%;">TRADE / BUSINESS NAME</td>
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; width: 22%;">BUSINESS CODE</td>
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; width: 45%;">NATURE OF BUSINESS / TRADE</td>
  </tr>
  <tr style="color: #0f172a;">
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; font-weight: bold;">{safe(data.business.name or "N/A")}</td>
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: #1e3a8a;">{safe(data.business.code or "N/A")}</td>
    <td style="padding: 6px 10px; border: 1px solid #cbd5e1; color: #475569;">{safe(data.business.description or "N/A")}</td>
  </tr>
</table>

<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 10pt; font-family: 'Segoe UI', Calibri, Arial, sans-serif; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #f1f5f9; font-weight: bold; color: #334155;">
      <th style="text-align: left; padding: 7px 10px; border: 1px solid #cbd5e1; width: 42%;">TRANSACTION RECEIPTS CLASSIFICATION</th>
      <th style="text-align: center; padding: 7px 10px; border: 1px solid #cbd5e1; width: 13%;">MIN %</th>
      <th style="text-align: right; padding: 7px 10px; border: 1px solid #cbd5e1; width: 22%;">GROSS RECEIPTS (Rs.)</th>
      <th style="text-align: right; padding: 7px 10px; border: 1px solid #cbd5e1; width: 23%;">PRESUMPTIVE INCOME (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; padding-left: 15px; color: #334155;">Digital / Banking Channels (u/s 44AD(1) Proviso)</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; color: #16a34a; font-weight: bold;">6%</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.business.banking_receipt)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.business.income_6)}</td>
    </tr>
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; padding-left: 15px; color: #334155;">Cash / Standard Transactions</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; color: #3b82f6; font-weight: bold;">8%</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.business.cash_receipt)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(0 if data.business.cash_receipt == 0 else data.business.income_8)}</td>
    </tr>
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; padding-left: 15px; color: #334155;">Other Mode Receipts</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center; color: #3b82f6; font-weight: bold;">8%</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.business.other_receipt)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.business.income_8)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f8fafc; border-top: 1.5pt solid #475569;">
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; color: #0f172a;">Total Presumptive Receipts &amp; Profit u/s 44AD</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.business.gross_receipt)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #1e3a8a; font-size: 10.5pt;">{fmt_money(data.business.income_total)}</td>
    </tr>
  </tbody>
</table>"""

    other_income_schedule_html = ""
    if data.other_sources_income != 0:
        other_income_schedule_html = f"""
<h3 style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11.5pt; font-weight: bold; color: #1e3a8a; margin: 10pt 0in 4pt; mso-margin-top-alt: 10pt; mso-margin-bottom-alt: 4pt; border-bottom: 2px solid #cbd5e1; padding-bottom: 4px; text-transform: uppercase;">SCHEDULE OS: INCOME FROM OTHER SOURCES</h3>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 10pt; font-family: 'Segoe UI', Calibri, Arial, sans-serif; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #f1f5f9; color: #334155; font-weight: bold;">
      <th style="text-align: left; padding: 7px 10px; border: 1px solid #cbd5e1; width: 75%;">NATURE / DETAILS OF OTHER INCOME</th>
      <th style="text-align: right; padding: 7px 10px; border: 1px solid #cbd5e1; width: 25%;">NET AMOUNT (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; padding-left: 15px; color: #0f172a;">{safe(data.other_income_description)}</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.other_sources_income)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f8fafc; border-top: 1.5pt solid #cbd5e1;">
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; color: #0f172a;">Total Income from Other Sources</td>
      <td style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.other_sources_income)}</td>
    </tr>
  </tbody>
</table>"""

    business_comparison_html = ""
    if data.itr_form == "ITR-4" or data.business_income != 0:
        business_comparison_html = f"""
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px;">Income from Business &amp; Profession</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; font-weight: 500;">{fmt_money(data.business_income)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.business_income)}</td>
    </tr>"""

    salary_comparison_html = ""
    if data.salary_income != 0:
        salary_comparison_html = f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; padding-left: 15px;">Income under the head "Salaries"</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; font-weight: bold; color: #0f172a;">{fmt_money(data.salary_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.salary_income)}</td>
    </tr>"""

    house_property_comparison_html = ""
    if data.house_property_income != 0:
        house_property_comparison_html = f"""
    <tr>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; padding-left: 15px;">Income / (Loss) from House Property</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; font-weight: bold; color: #0f172a;">{fmt_money(data.house_property_income)}</td>
      <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.house_property_income)}</td>
    </tr>"""

    savings = old_total - data.total_tax_cess
    if savings > 0:
        savings_str = f"Net Saving Benefit of <b>Rs. {fmt_money(savings)}</b> by selecting New Regime"
    elif savings < 0:
        savings_str = f"Old Regime would have been cheaper by <b>Rs. {fmt_money(abs(savings))}</b>"
    else:
        savings_str = "Neutral - Tax liabilities are identical under both regimes"

    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="ProgId" content="Word.Document">
<title>{safe(data.full_name)} - Tax Computation</title>
<style>
@page {{ size: A4; margin: 0.4in; }}
body {{ font-family: "Segoe UI", Calibri, Arial, sans-serif; font-size: 10pt; color: #1e293b; line-height: 1.3; margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
h1 {{ font-family: "Segoe UI", Calibri, Arial, sans-serif; font-weight: bold; margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
h2 {{ font-family: "Segoe UI", Calibri, Arial, sans-serif; font-weight: bold; margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
h3 {{ font-family: "Segoe UI", Calibri, Arial, sans-serif; font-weight: bold; margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
p {{ margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
table {{ border-collapse: collapse; margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
tr {{ margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
td {{ margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
th {{ margin: 0in 0in 0.0001pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; padding: 0pt; }}
.amount {{ text-align: right; white-space: nowrap; }}
.center {{ text-align: center; }}
.label {{ font-weight: bold; color: #475569; }}
</style>
</head>
<body>

<table style="width: 100%; border-collapse: collapse; margin-top: 1in; mso-margin-top-alt: 72pt; margin-bottom: 2pt;">
  <tr><td style="background-color: #1e3a8a; height: 3px; padding: 0; line-height: 1px;">&nbsp;</td></tr>
</table>

<h1 style="font-size: 13.5pt; font-weight: bold; color: #0f172a; text-align: center; margin: 0in 0in 2pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 2pt; text-transform: uppercase; letter-spacing: 0.5px;">{safe(data.full_name)}</h1>
<h2 style="font-size: 10pt; font-weight: 600; color: #475569; text-align: center; margin: 0in 0in 6pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 6pt;">STATEMENT OF COMPUTATION OF TOTAL INCOME &amp; TAX LIABILITY</h2>

<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 9.5pt; border: 1px solid #cbd5e1;">
  <tr>
    <th colspan="4" style="background-color: #1e3a8a; color: #ffffff; text-align: left; padding: 4px 6px; font-weight: bold; font-size: 9pt; text-transform: uppercase; border: 1px solid #1e3a8a;">Taxpayer Profile Summary</th>
  </tr>
  <tr>
    <td style="width: 18%; border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Assessment Year</td>
    <td style="width: 32%; border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a; font-weight: bold;">{safe(data.assessment_year)}</td>
    <td style="width: 18%; border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Financial Year</td>
    <td style="width: 32%; border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a; font-weight: bold;">{safe(fy)}</td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">PAN of Taxpayer</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a; font-weight: bold; font-family: monospace; letter-spacing: 0.5px;">{safe(data.pan)}</td>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Date of Birth</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;">{safe(data.dob)}</td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Address</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;" colspan="3">{safe(data.address)}<br>{safe(data.city)}, {safe(data.state)} - {safe(data.pin)}</td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Father's Name</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;">{safe(data.father_name)}</td>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Assessee Status</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;">{safe(data.status)} / Resident</td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Filing Status</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #16a34a; font-weight: bold;">E-Filed <span style="font-weight: normal; font-size: 7.5pt; color: #64748b;">(Form {safe(data.itr_form)})</span></td>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Selected Regime</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #1e3a8a; font-weight: bold;">New Tax Regime <span style="font-weight: normal; font-size: 7.5pt; color: #64748b;">(u/s 115BAC)</span></td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">E-Filing Ack No.</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a; font-family: monospace;">{safe(data.acknowledgement)}</td>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Date of Filing</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;">{safe(data.filing_date)}</td>
  </tr>
  <tr>
    <td style="border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 3px 5px; font-weight: bold; color: #334155;">Primary Email ID</td>
    <td style="border: 1px solid #cbd5e1; padding: 3px 5px; color: #0f172a;" colspan="3">{safe(data.email)}</td>
  </tr>
</table>

<h2 style="font-size: 10.5pt; font-weight: bold; color: #0f172a; margin: 8pt 0in 4pt; mso-margin-top-alt: 8pt; mso-margin-bottom-alt: 4pt; text-transform: uppercase;">I. Computation of Total Income</h2>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 10pt; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #1e3a8a; color: #ffffff;">
      <th style="text-align: left; padding: 4px 6px; border: 1px solid #1e3a8a; width: 45%;">PARTICULARS OF INCOME HEADS</th>
      <th style="text-align: center; padding: 4px 6px; border: 1px solid #1e3a8a; width: 15%;">SECTION</th>
      <th style="text-align: right; padding: 4px 6px; border: 1px solid #1e3a8a; width: 20%;">INNER AMOUNT (Rs.)</th>
      <th style="text-align: right; padding: 4px 6px; border: 1px solid #1e3a8a; width: 20%;">NET AMOUNT (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    {income_rows_html_str}
    
    <tr style="font-weight: bold; background-color: #f1f5f9; border-top: 1.2pt solid #475569; border-bottom: 1.2pt solid #475569;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">GROSS TOTAL INCOME</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.gross_total_income)}</td>
    </tr>
    
    <tr style="color: #475569;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Less: Deductions under Chapter VI-A</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Chapter VI-A</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #b91c1c;">-{fmt_money(data.total_deductions)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right;">&nbsp;</td>
    </tr>
    
    <tr style="font-weight: bold; background-color: #f8fafc; border-top: 1.2pt solid #475569; border-bottom: 3px double #1e3a8a;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; color: #1e3a8a;">TOTAL TAXABLE INCOME<br><span style="font-size: 7.5pt; font-weight: normal; color: #64748b; font-style: italic;">(Rounded off to nearest Rs. 10 u/s 288A)</span></td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt; color: #1e3a8a;">Sec 288A</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #1e3a8a; font-size: 10pt;">{fmt_money(data.total_income)}</td>
    </tr>
  </tbody>
</table>

<h2 style="font-size: 10.5pt; font-weight: bold; color: #0f172a; margin: 8pt 0in 4pt; mso-margin-top-alt: 8pt; mso-margin-bottom-alt: 4pt; text-transform: uppercase;">II. Computation of Tax Liability</h2>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 10pt; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #1e3a8a; color: #ffffff;">
      <th style="text-align: left; padding: 4px 6px; border: 1px solid #1e3a8a; width: 60%;">TAX COMPUTATION PARTICULARS</th>
      <th style="text-align: center; padding: 4px 6px; border: 1px solid #1e3a8a; width: 18%;">SECTION</th>
      <th style="text-align: right; padding: 4px 6px; border: 1px solid #1e3a8a; width: 22%;">TAX LIABILITY (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; font-weight: 500;">Basic Income Tax on Total Taxable Income</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 115BAC</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.tax_on_income)}</td>
    </tr>
    <tr style="color: #475569;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Less: Tax Rebate on Net Income</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 87A</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #dc2626;">-{fmt_money(data.rebate)}</td>
    </tr>
    <tr style="font-weight: 500;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">Net Basic Tax Payable</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(tax_after_rebate)}</td>
    </tr>
    <tr style="color: #475569;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Add: Health and Education Cess (4% of Tax)</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Cess</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right;">{fmt_money(data.cess)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f8fafc; border-top: 1.2pt solid #cbd5e1;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;"><b>Gross Tax &amp; Cess Liability</b></td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.total_tax_cess)}</td>
    </tr>
    
    <tr style="color: #475569;">
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Add: Interest for delay in filing return</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 234A</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: right;">{fmt_money(data.interest_234a)}</td>
    </tr>
    <tr style="color: #475569;">
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Add: Interest for default in payment of advance tax</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 234B</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: right;">{fmt_money(data.interest_234b)}</td>
    </tr>
    <tr style="color: #475569;">
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Add: Interest for deferment of advance tax</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 234C</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: right;">{fmt_money(data.interest_234c)}</td>
    </tr>
    <tr style="color: #475569;">
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; padding-left: 15px;">Add: Late Fee for delay in filing return</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Sec 234F</td>
      <td style="padding: 3px 6px; border: 1px solid #cbd5e1; text-align: right;">{fmt_money(data.fee_234f)}</td>
    </tr>
    
    <tr style="font-weight: bold; background-color: #f1f5f9; border-top: 1.2pt solid #475569; border-bottom: 1.2pt solid #475569;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">TOTAL LIABILTY (TAX, INTEREST &amp; LATE FEE)</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.total_tax_fee_interest)}</td>
    </tr>
    
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 15px; color: #475569;">Less: Prepaid Taxes Paid (TDS/TCS/Advance/Self-Assessment)</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 7.5pt;">Prepaid</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #16a34a; font-weight: bold;">-{fmt_money(data.taxes_paid)}</td>
    </tr>
    
    {balance_row_html}
  </tbody>
</table>

<h3 style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 10pt; font-weight: bold; color: #334155; margin: 8pt 0in 4pt; mso-margin-top-alt: 8pt; mso-margin-bottom-alt: 4pt; border-bottom: 2px solid #cbd5e1; padding-bottom: 3px; text-transform: uppercase;">Schedule: Progressive Tax Slab Details (New Regime - u/s 115BAC)</h3>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 9.5pt; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #f8fafc; color: #334155; font-weight: bold;">
      <th style="text-align: left; padding: 4px 6px; border: 1px solid #cbd5e1; width: 45%;">INCOME SLAB RANGE</th>
      <th style="text-align: center; padding: 4px 6px; border: 1px solid #cbd5e1; width: 20%;">RATE</th>
      <th style="text-align: right; padding: 4px 6px; border: 1px solid #cbd5e1; width: 35%;">TAX AMOUNT (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    {slab_rows_html}
    <tr style="font-weight: bold; background-color: #fafafa; border-top: 1.2pt solid #cbd5e1;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">Total Basic Tax on Slabs</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center;">&nbsp;</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.tax_on_income)}</td>
    </tr>
  </tbody>
</table>

{business_schedule_html}

{other_income_schedule_html}

<h2 style="font-size: 11pt; font-weight: bold; color: #0f172a; margin: 10pt 0in 4pt; mso-margin-top-alt: 10pt; mso-margin-bottom-alt: 4pt; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; text-transform: uppercase;">Annexure A: Tax Regime Comparison Sheet</h2>
<p style="font-size: 8pt; color: #475569; margin: 0in 0in 4pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 4pt; line-height: 1.3;">
  This comparison analysis is provided for validation. Chapter VI-A deductions are assumed standard.
</p>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 9.5pt; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #f1f5f9; color: #334155; font-weight: bold;">
      <th style="text-align: left; padding: 5px 6px; border: 1px solid #cbd5e1; width: 40%;">PARTICULARS</th>
      <th style="text-align: right; padding: 5px 6px; border: 1px solid #cbd5e1; width: 30%; background-color: #e0f2fe; color: #0369a1; border-bottom: 2px solid #0284c7;">NEW REGIME (Rs.) &nbsp;★ OPTIMAL</th>
      <th style="text-align: right; padding: 5px 6px; border: 1px solid #cbd5e1; width: 30%;">OLD REGIME (Rs.)</th>
    </tr>
  </thead>
  <tbody>
    {business_comparison_html}
    {salary_comparison_html}
    {house_property_comparison_html}
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px;">Income from Other Sources</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; font-weight: 500;">{fmt_money(data.other_sources_income)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(data.other_sources_income)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f8fafc;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">Gross Total Income</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #e0f2fe; color: #0f172a;">{fmt_money(data.gross_total_income)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.gross_total_income)}</td>
    </tr>
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px; color: #475569;">Less: Chapter VI-A Deductions</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; color: #94a3b8;">-{fmt_money(data.total_deductions)} <span style="font-size: 7pt; font-weight: normal; font-style: italic; display: block;">(Restricted u/s 115BAC)</span></td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">-{fmt_money(data.total_deductions)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f8fafc;">
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1;">Net Taxable Income</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #e0f2fe; color: #1e3a8a;">{fmt_money(data.total_income)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #0f172a;">{fmt_money(data.total_income)}</td>
    </tr>
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px;">Basic Tax Liability at Slab Rates</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff;">{fmt_money(data.tax_on_income)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(old_tax)}</td>
    </tr>
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px;">Less: Tax Rebate u/s 87A</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff; color: #dc2626;">-{fmt_money(data.rebate)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #94a3b8;">-</td>
    </tr>
    <tr>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; padding-left: 10px;">Add: Health &amp; Education Cess (4%)</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #f0f9ff;">{fmt_money(data.cess)}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: right; color: #475569;">{fmt_money(old_cess)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #f1f5f9; border-top: 1.2pt solid #cbd5e1; border-bottom: 3px double #0f172a;">
      <td style="padding: 5px 6px; border: 1px solid #cbd5e1;">TOTAL REGIME TAX LIABILITY</td>
      <td style="padding: 5px 6px; border: 1px solid #cbd5e1; text-align: right; background-color: #e0f2fe; color: #1e3a8a; font-size: 9.5pt;">{fmt_money(data.total_tax_cess)}</td>
      <td style="padding: 5px 6px; border: 1px solid #cbd5e1; text-align: right; color: #b91c1c; font-size: 9pt;">{fmt_money(old_total)}</td>
    </tr>
    <tr style="font-weight: bold; background-color: #ecfdf5; color: #047857;">
      <td style="padding: 5px 6px; border: 1px solid #cbd5e1;" colspan="3">Benefit Analysis: {savings_str}</td>
    </tr>
  </tbody>
</table>

<h3 style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 9.5pt; font-weight: bold; color: #334155; margin: 8pt 0in 4pt; mso-margin-top-alt: 8pt; mso-margin-bottom-alt: 4pt; border-bottom: 2px solid #e2e8f0; padding-bottom: 3px; text-transform: uppercase;">Schedule BA: Primary Bank Account Details</h3>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 6pt; mso-margin-bottom-alt: 6pt; font-size: 9.5pt; border: 1px solid #cbd5e1;">
  <thead>
    <tr style="background-color: #f8fafc; color: #334155; font-weight: bold;">
      <th style="text-align: center; padding: 4px 5px; border: 1px solid #cbd5e1; width: 8%;">S.NO</th>
      <th style="text-align: left; padding: 4px 6px; border: 1px solid #cbd5e1; width: 35%;">BANK NAME</th>
      <th style="text-align: center; padding: 4px 6px; border: 1px solid #cbd5e1; width: 20%;">IFSC CODE</th>
      <th style="text-align: left; padding: 4px 6px; border: 1px solid #cbd5e1; width: 22%;">ACCOUNT NUMBER</th>
      <th style="text-align: center; padding: 4px 6px; border: 1px solid #cbd5e1; width: 15%;">A/C TYPE</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="padding: 4px 5px; border: 1px solid #cbd5e1; text-align: center; color: #64748b;">1</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; font-weight: bold; color: #0f172a;">{safe(data.bank_name or "N/A")}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: #1e3a8a; letter-spacing: 0.5px;">{safe(data.bank_ifsc or "N/A")}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; font-family: monospace; color: #0f172a; letter-spacing: 0.5px;">{safe(data.bank_account or "N/A")}</td>
      <td style="padding: 4px 6px; border: 1px solid #cbd5e1; text-align: center; color: #475569;">{safe(data.bank_type or "N/A")}</td>
    </tr>
  </tbody>
</table>

<table style="width: 100%; border-collapse: collapse; margin-top: 15pt; mso-margin-top-alt: 15pt; font-size: 9.5pt;">
  <tr>
    <td style="width: 50%;">&nbsp;</td>
    <td style="width: 50%; text-align: right; vertical-align: bottom;">
      <div style="margin-bottom: 60px; mso-margin-bottom-alt: 60px; color: #334155; font-weight: bold;">For {safe(data.full_name)}</div>
      <div style="border-top: 1px solid #94a3b8; width: 160px; display: inline-block; margin-bottom: 4px;">&nbsp;</div>
      <div style="color: #475569; font-size: 8.5pt; padding-right: 15px; font-weight: bold;">Taxpayer Signature</div>
    </td>
  </tr>
</table>

</body>
</html>"""


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
