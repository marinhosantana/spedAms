from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path


def parse_decimal(value: str) -> Decimal:
    text = str(value or "").strip()
    if not text:
        return Decimal("0")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def parse_rate(value: str) -> Decimal | None:
    cleaned = (value or "").strip()
    if not cleaned:
        return None
    return parse_decimal(cleaned)


def normalize_sped_line(line: str) -> list[str]:
    return line.rstrip("\n\r").split("|")


def get_field(fields: list[str], index: int) -> str:
    if index < len(fields):
        return fields[index].strip()
    return ""


def first_non_empty(*values: str) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def normalize_document_key(value: str) -> str:
    return "".join(char for char in str(value or "").strip() if char.isdigit())


def normalize_cfop(value: str) -> str:
    return "".join(char for char in str(value or "").strip() if char.isdigit())


def infer_operation_type_from_cfop(cfop: str) -> str:
    normalized_cfop = normalize_cfop(cfop)
    if normalized_cfop == "1605":
        return "Saida"
    if normalized_cfop == "5605":
        return "Entrada"
    if normalized_cfop.startswith(("1", "2", "3")):
        return "Entrada"
    if normalized_cfop.startswith(("5", "6", "7")):
        return "Saida"
    return ""


def read_sped_plain_lines(file_path: Path) -> list[str]:
    lines = file_path.read_text(encoding="latin-1").splitlines()
    plain_lines: list[str] = []
    for line in lines:
        plain_lines.append(line)
        if line.startswith("|9999|"):
            break
    return plain_lines


def read_sped_e110_summary(file_path: Path) -> dict[str, Decimal]:
    summary = {
        "debitos": Decimal("0"),
        "ajustes_debito_doc": Decimal("0"),
        "ajustes_debito": Decimal("0"),
        "estornos_credito": Decimal("0"),
        "creditos": Decimal("0"),
        "ajustes_credito_doc": Decimal("0"),
        "ajustes_credito": Decimal("0"),
        "estornos_debito": Decimal("0"),
        "saldo_credor_anterior": Decimal("0"),
        "saldo_devedor": Decimal("0"),
        "deducoes": Decimal("0"),
        "icms_recolher": Decimal("0"),
        "saldo_credor_transportar": Decimal("0"),
        "extra_apuracao": Decimal("0"),
    }
    if not file_path.exists() or not file_path.is_file():
        return summary

    with file_path.open("r", encoding="latin-1") as sped_file:
        for raw_line in sped_file:
            if not raw_line.startswith("|E110|"):
                continue
            fields = normalize_sped_line(raw_line)
            return {
                "debitos": parse_decimal(get_field(fields, 2)),
                "ajustes_debito_doc": parse_decimal(get_field(fields, 3)),
                "ajustes_debito": parse_decimal(get_field(fields, 4)),
                "estornos_credito": parse_decimal(get_field(fields, 5)),
                "creditos": parse_decimal(get_field(fields, 6)),
                "ajustes_credito_doc": parse_decimal(get_field(fields, 7)),
                "ajustes_credito": parse_decimal(get_field(fields, 8)),
                "estornos_debito": parse_decimal(get_field(fields, 9)),
                "saldo_credor_anterior": parse_decimal(get_field(fields, 10)),
                "saldo_devedor": parse_decimal(get_field(fields, 11)),
                "deducoes": parse_decimal(get_field(fields, 12)),
                "icms_recolher": parse_decimal(get_field(fields, 13)),
                "saldo_credor_transportar": parse_decimal(get_field(fields, 14)),
                "extra_apuracao": parse_decimal(get_field(fields, 15)),
            }
    return summary


def read_combined_e110_summary(file_paths: list[Path]) -> dict[str, Decimal]:
    combined = {
        "debitos": Decimal("0"),
        "ajustes_debito_doc": Decimal("0"),
        "ajustes_debito": Decimal("0"),
        "estornos_credito": Decimal("0"),
        "creditos": Decimal("0"),
        "ajustes_credito_doc": Decimal("0"),
        "ajustes_credito": Decimal("0"),
        "estornos_debito": Decimal("0"),
        "saldo_credor_anterior": Decimal("0"),
        "saldo_devedor": Decimal("0"),
        "deducoes": Decimal("0"),
        "icms_recolher": Decimal("0"),
        "saldo_credor_transportar": Decimal("0"),
        "extra_apuracao": Decimal("0"),
    }
    for file_path in file_paths:
        current = read_sped_e110_summary(file_path)
        for key in combined:
            combined[key] += Decimal(current[key])
    return combined


SUMMARY_REGISTER_TYPES = {"C190", "C590", "D190", "D590", "D730"}


def read_sped_summary_register_rows(file_path: Path) -> list[dict[str, object]]:
    if not file_path.exists():
        return []

    rows: list[dict[str, object]] = []
    participants: dict[str, dict[str, str]] = {}
    current_document: dict[str, str] = {}
    with file_path.open("r", encoding="latin-1") as sped_file:
        for line_number, raw_line in enumerate(sped_file, start=1):
            if not raw_line.startswith("|"):
                continue
            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)

            if register == "0150":
                participant_code = get_field(fields, 2)
                participants[participant_code] = {
                    "name": get_field(fields, 3),
                    "tax_id": first_non_empty(get_field(fields, 5), get_field(fields, 4)),
                }
                continue

            if register == "C100":
                ind_oper = get_field(fields, 2)
                participant_code = get_field(fields, 4)
                participant_data = participants.get(participant_code, {})
                current_document = {
                    "operation_type": "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else "",
                    "document_number": get_field(fields, 8),
                    "document_key": get_field(fields, 9),
                    "document_date": get_field(fields, 10),
                    "document_series": get_field(fields, 7),
                    "document_model": get_field(fields, 5),
                    "participant_code": participant_code,
                    "participant_name": participant_data.get("name", ""),
                    "participant_tax_id": participant_data.get("tax_id", ""),
                }
                continue

            if register not in SUMMARY_REGISTER_TYPES:
                continue

            cfop = get_field(fields, 3)
            is_c190 = register == "C190"
            operation_type = current_document.get("operation_type", "") if is_c190 else ""
            if not operation_type:
                operation_type = infer_operation_type_from_cfop(cfop)
            rows.append(
                {
                    "source_register": register,
                    "operation_type": operation_type,
                    "document_number": current_document.get("document_number", "") if is_c190 else "",
                    "document_key": current_document.get("document_key", "") if is_c190 else "",
                    "document_date": current_document.get("document_date", "") if is_c190 else "",
                    "document_series": current_document.get("document_series", "") if is_c190 else "",
                    "document_model": current_document.get("document_model", "") if is_c190 else "",
                    "participant_code": current_document.get("participant_code", "") if is_c190 else "",
                    "participant_name": current_document.get("participant_name", "") if is_c190 else "",
                    "participant_tax_id": current_document.get("participant_tax_id", "") if is_c190 else "",
                    "cst_icms": get_field(fields, 2),
                    "cfop": cfop,
                    "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
                    "total_operation_value": parse_decimal(get_field(fields, 5)),
                    "base_icms": parse_decimal(get_field(fields, 6)),
                    "icms_value": parse_decimal(get_field(fields, 7)),
                    "base_icms_st": parse_decimal(get_field(fields, 8)),
                    "icms_st_value": parse_decimal(get_field(fields, 9)),
                    "reduction_value": parse_decimal(get_field(fields, 10)),
                    "ipi_value": parse_decimal(get_field(fields, 11)) if register == "C190" else Decimal("0"),
                    "line_number": line_number,
                    "raw_line": raw_line.rstrip("\r\n"),
                }
            )
    return rows
