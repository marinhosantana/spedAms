from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from decimal import Decimal

from app.parsers.sped_parser import parse_decimal


def normalize_text(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value or "")
    no_accents = "".join(char for char in decomposed if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", no_accents).strip().upper()


def normalize_tax_code(value: object, digits: int) -> str:
    raw_text = str(value or "").strip()
    if not raw_text:
        return ""
    only_digits = "".join(char for char in raw_text if char.isdigit())
    if not only_digits:
        return ""
    return only_digits.zfill(digits)[-digits:]


VALID_PIS_COFINS_CST_SAIDA = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "49"}
VALID_PIS_COFINS_CST_ENTRADA = {
    "50", "51", "52", "53", "54", "55", "56",
    "60", "61", "62", "63", "64", "65", "66", "67",
    "70", "71", "72", "73", "74", "75",
    "98",
}
VALID_PIS_COFINS_CST_AMBOS = {"99"}


def classify_pis_cofins_cst(value: object) -> str:
    normalized = normalize_tax_code(value, 2)
    if normalized in VALID_PIS_COFINS_CST_AMBOS:
        return "ambos"
    if normalized in VALID_PIS_COFINS_CST_ENTRADA:
        return "entrada"
    if normalized in VALID_PIS_COFINS_CST_SAIDA:
        return "saida"
    return ""


def build_pis_cofins_side_values(cst_value: object, aliquota_value: object) -> tuple[str, str, Decimal, Decimal, list[str]]:
    normalized_cst = normalize_tax_code(cst_value, 2)
    aliquota = Decimal(aliquota_value or Decimal("0"))
    cst_kind = classify_pis_cofins_cst(normalized_cst)
    warnings: list[str] = []
    if normalized_cst and not cst_kind:
        warnings.append(f"CST {normalized_cst} de PIS/COFINS fora das faixas validas de entrada/saida.")
    if cst_kind == "entrada":
        return normalized_cst, "", aliquota, Decimal("0"), warnings
    if cst_kind == "saida":
        return "", normalized_cst, Decimal("0"), aliquota, warnings
    if cst_kind == "ambos":
        return normalized_cst, normalized_cst, aliquota, aliquota, warnings
    return "", "", Decimal("0"), Decimal("0"), warnings


def normalize_operation_type(value: str) -> str:
    normalized = normalize_text(value)
    if normalized.startswith("SAID"):
        return "Saida"
    if normalized.startswith("ENTRAD"):
        return "Entrada"
    return str(value or "").strip()


def format_decimal_sped(value: Decimal) -> str:
    normalized = value.quantize(Decimal("0.01"))
    return format(normalized, "f").replace(".", ",")


def compute_display_icms_rate(
    nominal_rate: object,
    sale_value: object,
    base_icms: object,
    icms_value: object,
) -> Decimal:
    rate = Decimal(nominal_rate) if isinstance(nominal_rate, Decimal) else parse_decimal(str(nominal_rate or "0"))
    sale = Decimal(sale_value) if isinstance(sale_value, Decimal) else parse_decimal(str(sale_value or "0"))
    base = Decimal(base_icms) if isinstance(base_icms, Decimal) else parse_decimal(str(base_icms or "0"))
    icms = Decimal(icms_value) if isinstance(icms_value, Decimal) else parse_decimal(str(icms_value or "0"))
    if sale > 0 and base > 0 and icms >= 0 and sale.quantize(Decimal("0.01")) != base.quantize(Decimal("0.01")):
        effective_rate = (icms * Decimal("100") / sale).quantize(Decimal("0.01"))
        if Decimal("0") < effective_rate < Decimal("1") and rate > Decimal("0"):
            scaled_effective_rate = (effective_rate * Decimal("100")).quantize(Decimal("0.01"))
            if scaled_effective_rate <= Decimal("100"):
                return scaled_effective_rate
        return effective_rate
    if rate > 0:
        return rate.quantize(Decimal("0.01"))
    if base > 0 and icms >= 0:
        return (icms * Decimal("100") / base).quantize(Decimal("0.01"))
    return Decimal("0.00")


def extract_representative_icms_rate(row: dict[str, object]) -> Decimal:
    for detail in row.get("launch_details", []):
        detail_rate = detail.get("icms_rate", Decimal("0"))
        rate = detail_rate if isinstance(detail_rate, Decimal) else parse_decimal(str(detail_rate or "0"))
        if rate > 0:
            return rate.quantize(Decimal("0.01"))
    current_rate = row.get("icms_rate", Decimal("0"))
    rate = current_rate if isinstance(current_rate, Decimal) else parse_decimal(str(current_rate or "0"))
    return rate.quantize(Decimal("0.01"))


def extract_effective_icms_rate(row: dict[str, object]) -> Decimal:
    stored_rate = row.get("display_icms_rate")
    if isinstance(stored_rate, Decimal):
        return stored_rate.quantize(Decimal("0.01"))
    if stored_rate is not None:
        return parse_decimal(str(stored_rate)).quantize(Decimal("0.01"))
    return compute_display_icms_rate(
        extract_representative_icms_rate(row),
        Decimal(row.get("sale_value", Decimal("0"))),
        Decimal(row.get("base_icms", Decimal("0"))),
        Decimal(row.get("icms_value", Decimal("0"))),
    )


def has_icms_reduction(
    nominal_rate: object,
    sale_value: object,
    base_icms: object,
    icms_value: object,
) -> bool:
    nominal = Decimal(nominal_rate) if isinstance(nominal_rate, Decimal) else parse_decimal(str(nominal_rate or "0"))
    effective = compute_display_icms_rate(nominal_rate, sale_value, base_icms, icms_value)
    if nominal <= 0 or effective <= 0:
        return False
    return nominal.quantize(Decimal("0.01")) != effective.quantize(Decimal("0.01"))


def normalize_cst_icms_for_sped(value: str) -> str:
    digits = "".join(char for char in str(value or "").strip() if char.isdigit())
    if not digits:
        return "000"
    if len(digits) >= 3:
        return digits[-3:]
    return digits.zfill(3)


def merge_product_rate(
    product_rates: defaultdict[str, set[Decimal]] | dict[str, set[Decimal]],
    code: str,
    rate: Decimal | None,
) -> None:
    if code and rate is not None:
        product_rates[code].add(rate)


def merge_product_cst(
    product_csts: defaultdict[str, set[str]] | dict[str, set[str]],
    code: str,
    cst_icms: str,
) -> None:
    if code and cst_icms:
        product_csts[code].add(cst_icms)


def normalize_header(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", normalize_text(value))


def resolve_header(row: dict[str, str], *aliases: str) -> str:
    for alias in aliases:
        if alias in row and str(row[alias]).strip():
            return str(row[alias]).strip()
    return ""
