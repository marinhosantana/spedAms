from __future__ import annotations

import datetime as dt
import re
from decimal import Decimal
from pathlib import Path

from app.parsers.sped_parser import get_field, normalize_sped_line
from app.services.tax_rules import normalize_operation_type

def filter_details_by_operation_scope(
    detailed_sales: list[dict[str, object]],
    operation_scope: str,
) -> list[dict[str, object]]:
    scope = str(operation_scope or "").strip().lower()
    if scope == "entrada":
        return [item for item in detailed_sales if str(item["operation_type"]).strip().lower() == "entrada"]
    if scope == "saida":
        return [item for item in detailed_sales if str(item["operation_type"]).strip().lower() == "saida"]
    return list(detailed_sales)

def parse_sped_document_date(value: object) -> dt.date | None:
    digits = "".join(char for char in str(value or "").strip() if char.isdigit())
    if len(digits) != 8:
        return None
    try:
        if int(digits[:4]) >= 1900:
            return dt.datetime.strptime(digits, "%Y%m%d").date()
        return dt.datetime.strptime(digits, "%d%m%Y").date()
    except ValueError:
        return None

def infer_sped_period_label(file_path: Path, detailed_sales: list[dict[str, object]]) -> str:
    filename_match = re.search(r"-(\d{4})(\d{2})\d{2}-(\d{4})(\d{2})\d{2}-", file_path.name)
    if filename_match:
        return f"{filename_match.group(2)}/{filename_match.group(1)}"

    try:
        with file_path.open("r", encoding="latin-1", errors="ignore") as sped_file:
            for raw_line in sped_file:
                if not raw_line.startswith("|0000|"):
                    continue
                fields = normalize_sped_line(raw_line)
                period_start = parse_sped_document_date(get_field(fields, 4))
                period_end = parse_sped_document_date(get_field(fields, 5))
                period_date = period_start or period_end
                if period_date is not None:
                    return period_date.strftime("%m/%Y")
                break
    except OSError:
        pass

    dates = [
        parsed_date
        for item in detailed_sales
        if (parsed_date := parse_sped_document_date(item.get("document_date")))
    ]
    if dates:
        return min(dates).strftime("%m/%Y")
    match = re.search(r"(0[1-9]|1[0-2])([12]\d{3})", file_path.stem)
    if match:
        return f"{match.group(1)}/{match.group(2)}"
    return file_path.stem

def calculate_abc_curve_labels(
    totals_by_code: dict[str, Decimal],
    threshold_a: Decimal = Decimal("80"),
    threshold_b: Decimal = Decimal("95"),
) -> dict[str, str]:
    positive_items = [
        (str(code).strip(), Decimal(value))
        for code, value in totals_by_code.items()
        if str(code).strip() and Decimal(value) > 0
    ]
    grand_total = sum((value for _, value in positive_items), Decimal("0"))
    if grand_total <= 0:
        return {code: "C" for code in totals_by_code}

    curve_labels: dict[str, str] = {}
    cumulative_percent = Decimal("0")
    for code, value in sorted(positive_items, key=lambda item: item[1], reverse=True):
        cumulative_percent += (value * Decimal("100") / grand_total)
        if cumulative_percent <= threshold_a:
            curve_labels[code] = "A"
        elif cumulative_percent <= threshold_b:
            curve_labels[code] = "B"
        else:
            curve_labels[code] = "C"

    for code in totals_by_code:
        curve_labels.setdefault(str(code).strip(), "C")
    return curve_labels
