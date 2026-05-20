from __future__ import annotations

import re
from decimal import Decimal

from app.exporters.rules_report_exporter import format_rule_decimal
from app.parsers.sped_parser import normalize_document_key, parse_decimal
from app.services.tax_rules import (
    normalize_cst_icms_for_sped,
    normalize_operation_type,
    normalize_text,
)


def parse_replacement_value(raw_text: str) -> str:
    return str(raw_text or "").strip()


def decimal_rule_matches(candidate: object, expected: object) -> bool:
    if not isinstance(expected, Decimal):
        return False
    if not isinstance(candidate, Decimal):
        return False
    return candidate.quantize(Decimal("0.01")) == expected.quantize(Decimal("0.01"))


def parse_bool_flag(value: str) -> bool:
    normalized = normalize_text(value)
    return normalized in {"1", "S", "SIM", "TRUE", "VERDADEIRO", "YES"}


def extract_tax_id_from_document_key(document_key: object) -> str:
    normalized_key = normalize_document_key(str(document_key or ""))
    if len(normalized_key) == 44 and normalized_key.isdigit():
        return normalized_key[6:20]
    return ""


def normalize_runtime_rule_tax_id(value: object, line_number: int) -> str:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    if not digits:
        return ""
    if len(digits) == 44:
        return extract_tax_id_from_document_key(digits)
    if len(digits) in {11, 14}:
        return digits
    raise ValueError(
        f"Regra dinamica na linha {line_number}: CNPJ/CPF deve ter 14 ou 11 digitos. "
        "Se preferir, informe a chave NF-e completa com 44 digitos."
    )


def parse_runtime_rule_lines(raw_text: str) -> list[dict[str, object]]:
    runtime_rules: list[dict[str, object]] = []
    key_aliases = {
        "tipo": "operation_type",
        "tipo_operacao": "operation_type",
        "operacao": "operation_type",
        "documento": "document_number",
        "numero_documento": "document_number",
        "numero_doc": "document_number",
        "num_doc": "document_number",
        "cnpj": "document_tax_id",
        "cpf_cnpj": "document_tax_id",
        "cnpj_cpf": "document_tax_id",
        "cst": "cst_icms",
        "cfop": "cfop",
        "aliquota": "match_rate",
        "codigo": "match_codes",
        "codigos": "match_codes",
        "novo_cst": "new_cst",
        "novo_cfop": "new_cfop",
        "nova_aliquota": "force_rate",
        "zerar_icms": "zero_icms",
        "usar_valor_operacao_como_base": "use_sale_value_as_base",
        "definir_base_icms": "set_base_icms",
        "definir_valor_icms": "set_icms_value",
        "percentual_reducao_base": "base_reduction_percent",
        "reducao_base_percentual": "base_reduction_percent",
        "perc_reducao_base": "base_reduction_percent",
        "percentual_reducao_bc": "base_reduction_percent",
        "reducao_bc_percentual": "base_reduction_percent",
        "recalcular_valor_icms": "recalculate_icms_value",
        "recalcular_base_reducao": "recalculate_reduced_base",
        "recalcular_base_calculo_reducao": "recalculate_reduced_base",
        "recalcular_bc_reducao": "recalculate_reduced_base",
    }
    decimal_keys = {
        "match_rate",
        "force_rate",
        "set_base_icms",
        "set_icms_value",
        "base_reduction_percent",
    }
    bool_keys = {"zero_icms", "use_sale_value_as_base", "recalculate_icms_value", "recalculate_reduced_base"}

    for line_number, raw_line in enumerate((raw_text or "").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        rule: dict[str, object] = {}
        for fragment in [chunk.strip() for chunk in line.split(";") if chunk.strip()]:
            if "=" not in fragment:
                raise ValueError(f"Regra dinamica na linha {line_number} invalida: use chave=valor.")
            raw_key, raw_value = fragment.split("=", 1)
            normalized_key = normalize_text(raw_key).replace(" ", "_").lower()
            target_key = key_aliases.get(normalized_key, normalized_key)
            value = raw_value.strip()
            if not value and target_key not in bool_keys:
                continue

            if target_key == "operation_type":
                operation_type = normalize_operation_type(value).strip().lower()
                if operation_type not in {"entrada", "saida"}:
                    raise ValueError(f"Regra dinamica na linha {line_number}: tipo deve ser Entrada ou Saida.")
                rule[target_key] = operation_type
                continue

            if target_key == "document_tax_id":
                rule[target_key] = normalize_runtime_rule_tax_id(value, line_number)
                continue

            if target_key in {"document_number", "cst_icms", "cfop", "new_cst", "new_cfop"}:
                rule[target_key] = value.strip()
                continue

            if target_key == "match_codes":
                rule[target_key] = {str(chunk).strip() for chunk in re.split(r"[\|,/]+", value) if str(chunk).strip()}
                continue

            if target_key in decimal_keys:
                rule[target_key] = parse_decimal(value)
                continue

            if target_key in bool_keys:
                rule[target_key] = parse_bool_flag(value)
                continue

            raise ValueError(f"Regra dinamica na linha {line_number}: chave '{raw_key}' nao suportada.")

        if "operation_type" not in rule:
            rule["operation_type"] = "entrada"
        runtime_rules.append(rule)

    return runtime_rules


def runtime_rule_summary(rule_line: str) -> str:
    summary_parts: list[str] = []
    for fragment in [chunk.strip() for chunk in str(rule_line or "").split(";") if chunk.strip()]:
        key, _, value = fragment.partition("=")
        normalized_key = normalize_text(key).replace(" ", "_").lower()
        if normalized_key in {
            "tipo",
            "documento",
            "numero_documento",
            "numero_doc",
            "num_doc",
            "cnpj",
            "cpf_cnpj",
            "cnpj_cpf",
            "cst",
            "cfop",
            "aliquota",
            "codigo",
            "codigos",
            "novo_cst",
            "novo_cfop",
            "nova_aliquota",
        }:
            summary_parts.append(f"{key.strip()}={value.strip()}")
    return "; ".join(summary_parts[:4]) or str(rule_line or "").strip()


DEFAULT_ICMS_RULE_PROFILES: dict[str, list[dict[str, object]]] = {
    "Regra Padrao Filial": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1152", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1152", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("3.09"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1101", "match_rate": Decimal("13.30"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "zero_icms": True},
        {"operation_type": "saida", "cst_icms": "040", "cfop": "5102", "new_cst": "041"},
        {"operation_type": "saida", "cst_icms": "041", "cfop": "5152", "match_rate": Decimal("0"), "new_cst": "090"},
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1253",
            "split_into": [
                {"new_cst": "090", "new_cfop": "1253", "sale_value_factor": Decimal("0.10"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1252",
                    "sale_value_factor": Decimal("0.90"),
                    "force_rate": Decimal("18"),
                    "use_sale_value_as_base": True,
                },
            ],
        },
        {"operation_type": "saida", "cst_icms": "100", "cfop": "5102", "new_cst": "200"},
        {"operation_type": "saida", "cst_icms": "120", "cfop": "5102", "new_cst": "220"},
        {"operation_type": "saida", "cst_icms": "160", "cfop": "5102", "new_cst": "260"},
        {"operation_type": "saida", "cst_icms": "141", "cfop": "5927", "new_cst": "241"},
    ],
    "Regra Padrao Matriz": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1411", "new_cfop": "1202"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1556", "new_cst": "041", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2403", "new_cfop": "2102"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2556", "new_cst": "041", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1403", "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "2403", "new_cfop": "2102"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "2401", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1407", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "2556", "zero_icms": True},
    ],
    "Regra Nova Matriz": [
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5102", "match_rates": {Decimal("4"), Decimal("12"), Decimal("18")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1152", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1152", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1101", "match_rate": Decimal("13.30"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rate": Decimal("12"), "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rate": Decimal("18"), "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1411", "match_rate": Decimal("12"), "new_cfop": "1202"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1556", "match_rate": Decimal("18"), "new_cst": "041", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2403", "match_rate": Decimal("12"), "new_cfop": "2102"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2556", "match_rate": Decimal("12"), "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1403", "match_rate": Decimal("18"), "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "2403", "match_rate": Decimal("12"), "new_cfop": "2102"},
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1253",
            "split_into": [
                {"new_cst": "090", "new_cfop": "1253", "sale_value_factor": Decimal("0.10"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1252",
                    "sale_value_factor": Decimal("0.90"),
                    "force_rate": Decimal("18"),
                    "use_sale_value_as_base": True,
                },
            ],
        },
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1407", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "2401", "new_cst": "000", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1152", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "2556", "match_rate": Decimal("12"), "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1401", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1407", "match_rate": Decimal("18"), "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "match_rate": Decimal("18"), "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.42"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.42"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2101", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.42"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2102", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.42"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
    ],
    "Nascimento e Curtarelli": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rates": {Decimal("12"), Decimal("18")}, "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "2403", "new_cfop": "2102"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "160", "cfop": "1401", "new_cst": "260", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "260", "cfop": "1401", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
    ],
    "BELLA CITTA 0001-02 - ENTRADAS-Saidas": [
        {
            "operation_type": "entrada",
            "cst_icms": "000",
            "cfop": "1253",
            "split_into": [
                {"new_cst": "090", "new_cfop": "1253", "sale_value_factor": Decimal("0.10"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1252",
                    "sale_value_factor": Decimal("0.90"),
                    "force_rate": Decimal("18"),
                    "use_sale_value_as_base": True,
                },
            ],
        },
        {
            "operation_type": "entrada",
            "cfop": "1556",
            "match_sale_value": Decimal("1638.14"),
            "new_cfop": "1101",
            "set_sale_value": Decimal("0"),
            "zero_icms": True,
        },
        {"operation_type": "entrada", "cst_icms": "010", "cfop": "1411", "match_rate": Decimal("12"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "010", "cfop": "1411", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1101", "new_cst": "040", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1407", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "061", "cfop": "1101", "new_cst": "040", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rate": Decimal("4"), "force_rate": Decimal("4"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "061", "cfop": "1401", "match_codes": {"0078957238", "0078957271"}, "new_cfop": "1651"},
        {
            "operation_type": "entrada",
            "cst_icms": "061",
            "cfop": "1653",
            "match_codes": {"0078957238", "0078957271"},
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.12"),
            "preserve_icms_value": True,
        },
        {
            "operation_type": "entrada",
            "cst_icms": "061",
            "cfop": "1651",
            "match_codes": {"0078957271"},
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.47"),
            "preserve_icms_value": True,
        },
        {
            "operation_type": "entrada",
            "cst_icms": "061",
            "cfop": "1651",
            "match_codes": {"0078957238"},
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.12"),
            "preserve_icms_value": True,
        },
        {"operation_type": "entrada", "cst_icms": "290", "cfop": "1101", "new_cst": "200"},
        {"operation_type": "entrada", "cst_icms": "200", "cfop": "1101", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1101", "new_cst": "040", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1401", "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1407", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "160", "cfop": "1401", "new_cst": "100", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "260", "cfop": "1401", "new_cst": "200", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "290", "cfop": "1401", "new_cst": "200", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "360", "cfop": "1401", "new_cst": "300", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cfop": "1556", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "560", "cfop": "1401", "new_cst": "500", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "590", "cfop": "1401", "match_sale_value": Decimal("780.00"), "new_cst": "500", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "590", "cfop": "1401", "match_sale_value": Decimal("760.00"), "new_cst": "800", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "890", "cfop": "1401", "new_cst": "800", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5101", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5102", "match_rate": Decimal("0"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5102", "match_sale_value": Decimal("83.72"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5102", "match_sale_value": Decimal("67.94"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5102", "match_rate": Decimal("100"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "000", "cfop": "5927", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "070", "cfop": "5401", "new_cst": "010"},
        {"operation_type": "saida", "cst_icms": "010", "cfop": "5401", "match_rate": Decimal("12"), "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "010", "cfop": "5401", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "070", "cfop": "5927", "new_cst": "010", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "saida", "cst_icms": "010", "cfop": "5927", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {
            "operation_type": "saida",
            "cst_icms": "020",
            "cfop": "5927",
            "set_base_icms": Decimal("2052.31"),
            "force_rate": Decimal("18"),
            "set_icms_value": Decimal("369.41"),
        },
        {"operation_type": "saida", "cst_icms": "070", "cfop": "5927", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
    ],
    "NASCIMENTO 0001-88 - ENTRADAS-Saidas": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2101", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2102", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("2.38"), Decimal("2.68"), Decimal("3.09"), Decimal("3.80")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2202", "match_rate": Decimal("3.80"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "100", "cfop": "1102", "match_rate": Decimal("1.44"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1101", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("18"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("25"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "2403", "match_rate": Decimal("12"), "new_cfop": "2401"},
        {"operation_type": "entrada", "cst_icms": "160", "cfop": "1401", "match_rate": Decimal("0"), "new_cfop": "1403"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rate": Decimal("18"), "new_cfop": "1401"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rate": Decimal("12"), "new_cst": "060"},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "match_rate": Decimal("12"), "new_cst": "000", "new_cfop": "1101", "force_rate": Decimal("12"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "match_rate": Decimal("18"), "new_cst": "000", "new_cfop": "1101", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1407", "match_rate": Decimal("0"), "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2403", "new_cst": "060", "set_sale_value": Decimal("0"), "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1403", "new_cst": "060", "set_sale_value": Decimal("0"), "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "match_rate": Decimal("0"), "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1401", "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1407", "new_cst": "000"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1653", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "061", "cfop": "1653", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1401", "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1253",
            "match_sale_value": Decimal("2140.72"),
            "split_into": [
                {"new_cst": "090", "new_cfop": "1253", "set_sale_value": Decimal("214.07"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1252",
                    "set_sale_value": Decimal("1926.65"),
                    "set_base_icms": Decimal("1926.65"),
                    "force_rate": Decimal("18"),
                    "set_icms_value": Decimal("346.80"),
                },
            ],
        },
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1102", "match_sale_value": Decimal("759.00"), "new_cst": "040", "zero_icms": True},
        {"operation_type": "saida", "cst_icms": "100", "new_cst": "200"},
        {"operation_type": "saida", "cst_icms": "160", "new_cst": "260"},
    ],
    "CASA DE PAES - ENTRADAS": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rates": {Decimal("1.25"), Decimal("1.44"), Decimal("3.94")}, "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1403", "match_rate": Decimal("18"), "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1403", "match_rate": Decimal("18"), "new_cfop": "1102"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1101", "match_rate": Decimal("0"), "new_cfop": "1403"},
        {
            "operation_type": "entrada",
            "cst_icms": "060",
            "cfop": "1401",
            "match_rate": Decimal("0"),
            "split_into": [
                {"new_cst": "060", "new_cfop": "1401", "sale_value_factor": Decimal("0.50"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1401",
                    "sale_value_factor": Decimal("0.50"),
                    "force_rate": Decimal("18"),
                    "use_sale_value_as_base": True,
                },
            ],
        },
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1403", "match_rate": Decimal("0"), "set_icms_value": Decimal("0")},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "2401", "match_rate": Decimal("0"), "new_cst": "000", "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "061", "cfop": "1401", "match_rate": Decimal("0"), "new_cfop": "1653"},
        {
            "operation_type": "entrada",
            "cst_icms": "061",
            "cfop": "1653",
            "match_rate": Decimal("0"),
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.12"),
            "preserve_icms_value": True,
        },
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1653",
            "match_rate": Decimal("0"),
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.12"),
            "preserve_icms_value": True,
        },
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1252",
            "split_into": [
                {"new_cst": "090", "new_cfop": "1253", "sale_value_factor": Decimal("0.10"), "zero_icms": True},
                {
                    "new_cst": "000",
                    "new_cfop": "1252",
                    "sale_value_factor": Decimal("0.90"),
                    "force_rate": Decimal("18"),
                    "use_sale_value_as_base": True,
                },
            ],
        },
        {"operation_type": "entrada", "cst_icms": "090", "cfop": "1556", "match_rate": Decimal("18"), "zero_icms": True},
        {
            "operation_type": "entrada",
            "cst_icms": "090",
            "cfop": "1553",
            "match_rate": Decimal("0"),
            "set_base_icms": Decimal("0"),
            "force_rate": Decimal("0"),
            "icms_per_quantity": Decimal("1.12"),
            "preserve_icms_value": True,
        },
    ],
    "NIG": [
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("12"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1121", "new_cst": "090", "new_cfop": "1551", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1250", "match_rate": Decimal("0"), "force_rate": Decimal("18"), "use_sale_value_as_base": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1556", "match_rate": Decimal("18"), "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1556", "match_rate": Decimal("7"), "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2102", "match_rate": Decimal("12"), "new_cfop": "2101"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "2102", "match_rate": Decimal("7"), "new_cfop": "2101"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1102", "match_rate": Decimal("12"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "020", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "040", "cfop": "2932", "match_rate": Decimal("0"), "new_cst": "041"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1403", "new_cst": "090", "new_cfop": "1407"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1556", "new_cst": "090", "new_cfop": "1653"},
        {"operation_type": "entrada", "cst_icms": "060", "cfop": "1653", "new_cst": "090", "new_cfop": "1653"},
        {"operation_type": "entrada", "cst_icms": "061", "cfop": "1556", "new_cfop": "1653"},
        {"operation_type": "entrada", "cst_icms": "100", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "100", "cfop": "2102", "match_rate": Decimal("4"), "new_cfop": "2101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("1.86"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("2.88"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("2.89"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("3.15"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("3.21"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("3.55"), "new_cst": "000", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1102", "match_rate": Decimal("0"), "new_cfop": "1101", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1124", "match_rate": Decimal("3.45"), "new_cst": "000"},
        {"operation_type": "entrada", "cst_icms": "101", "cfop": "1916", "new_cst": "090", "zero_icms": True},
        {"operation_type": "entrada", "cst_icms": "102", "cfop": "1102", "match_rate": Decimal("0"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "102", "cfop": "1403", "match_rate": Decimal("0"), "new_cst": "090", "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "102", "cfop": "2102", "match_rate": Decimal("0"), "new_cst": "090", "new_cfop": "2101"},
        {"operation_type": "entrada", "cst_icms": "200", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "202", "cfop": "1403", "match_rate": Decimal("0"), "new_cst": "090", "new_cfop": "1407"},
        {"operation_type": "entrada", "cst_icms": "260", "cfop": "1403", "match_rate": Decimal("0"), "new_cst": "290", "new_cfop": "1407"},
        {"operation_type": "entrada", "cst_icms": "300", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "360", "cfop": "1403", "match_rate": Decimal("0"), "new_cst": "390", "new_cfop": "1407"},
        {"operation_type": "entrada", "cst_icms": "500", "cfop": "1102", "match_rate": Decimal("18"), "new_cfop": "1101"},
        {"operation_type": "entrada", "cst_icms": "500", "cfop": "1403", "match_rate": Decimal("0"), "new_cst": "590", "new_cfop": "1407"},
        {"operation_type": "entrada", "cst_icms": "560", "cfop": "1653", "new_cst": "590"},
        {"operation_type": "entrada", "cst_icms": "900", "cfop": "1902", "new_cst": "090"},
        {"operation_type": "entrada", "cst_icms": "000", "cfop": "1102", "match_rate": Decimal("0"), "new_cst": "000", "new_cfop": "1101", "set_base_icms": Decimal("0"), "set_icms_value": Decimal("0")},
        {"operation_type": "saida", "cst_icms": "060", "cfop": "5102", "match_rate": Decimal("0"), "new_cst": "090", "new_cfop": "5662"},
        {"operation_type": "saida", "cst_icms": "060", "cfop": "5656", "match_rate": Decimal("0"), "new_cst": "090", "new_cfop": "5662"},
        {"operation_type": "saida", "cst_icms": "200", "cfop": "6102", "match_rate": Decimal("4"), "new_cst": "200", "new_cfop": "6556"},
    ],
}


def get_first_matching_icms_rule(
    rules: list[dict[str, object]],
    item: dict[str, object],
) -> dict[str, Decimal | str | bool] | None:
    operation_type = str(item.get("operation_type", "")).strip().lower()
    document_number = str(item.get("document_number", "")).strip()
    document_tax_id = str(item.get("document_tax_id", "")).strip()
    if not document_tax_id:
        document_tax_id = extract_tax_id_from_document_key(item.get("document_key", ""))
    if not document_tax_id:
        fallback_tax_id = normalize_document_key(str(item.get("participant_tax_id", "")))
        document_tax_id = fallback_tax_id if len(fallback_tax_id) in {11, 14} else ""
    cst_icms = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
    cfop = str(item.get("cfop", "")).strip()
    icms_rate = item.get("icms_rate") if isinstance(item.get("icms_rate"), Decimal) else Decimal("0")
    normalized_rate = icms_rate.quantize(Decimal("0.01"))
    sale_value = item.get("sale_value") if isinstance(item.get("sale_value"), Decimal) else Decimal("0")
    base_icms = item.get("base_icms") if isinstance(item.get("base_icms"), Decimal) else Decimal("0")
    code = str(item.get("code", "")).strip()

    for rule in rules:
        if str(rule.get("operation_type", "")).strip().lower() != operation_type:
            continue
        rule_document_number = str(rule.get("document_number", "")).strip()
        if rule_document_number and rule_document_number != document_number:
            continue
        rule_document_tax_id = str(rule.get("document_tax_id", "")).strip()
        if rule_document_tax_id and rule_document_tax_id != document_tax_id:
            continue
        rule_cst = str(rule.get("cst_icms", "")).strip()
        if rule_cst and normalize_cst_icms_for_sped(rule_cst) != cst_icms:
            continue
        rule_cfop = str(rule.get("cfop", "")).strip()
        if rule_cfop and rule_cfop != cfop:
            continue
        match_rates = rule.get("match_rates")
        if isinstance(match_rates, set):
            candidate_rates = {Decimal(str(rate)).quantize(Decimal("0.01")) for rate in match_rates}
            if normalized_rate not in candidate_rates:
                continue
        match_rate = rule.get("match_rate")
        if isinstance(match_rate, Decimal) and not decimal_rule_matches(normalized_rate, match_rate):
            continue
        match_sale_value = rule.get("match_sale_value")
        if isinstance(match_sale_value, Decimal) and not decimal_rule_matches(sale_value, match_sale_value):
            continue
        match_base_icms = rule.get("match_base_icms")
        if isinstance(match_base_icms, Decimal) and not decimal_rule_matches(base_icms, match_base_icms):
            continue
        match_codes = rule.get("match_codes")
        if isinstance(match_codes, set) and code not in {str(value).strip() for value in match_codes}:
            continue
        return rule
    return None


def build_rule_signature(rule: dict[str, object]) -> str:
    signature_parts: list[str] = []
    for key in sorted(rule.keys()):
        value = rule.get(key)
        if isinstance(value, Decimal):
            normalized_value = format_rule_decimal(value)
        elif isinstance(value, set):
            normalized_value = ",".join(sorted(str(item).strip() for item in value if str(item).strip()))
        elif isinstance(value, list):
            normalized_value = f"list:{len(value)}"
        else:
            normalized_value = str(value).strip()
        signature_parts.append(f"{key}={normalized_value}")
    return " | ".join(signature_parts)


def get_default_icms_rule(
    rule_profile: str,
    item: dict[str, object],
) -> dict[str, Decimal | str | bool] | None:
    return get_first_matching_icms_rule(DEFAULT_ICMS_RULE_PROFILES.get(rule_profile, []), item)


def get_configured_icms_rule(
    rule_profile: str,
    runtime_rules: list[dict[str, object]] | None,
    item: dict[str, object],
) -> dict[str, Decimal | str | bool] | None:
    if runtime_rules:
        runtime_match = get_first_matching_icms_rule(runtime_rules, item)
        if runtime_match:
            return runtime_match
    if rule_profile:
        return get_default_icms_rule(rule_profile, item)
    return None


def has_default_icms_rule(rule_profile: str, item: dict[str, object]) -> bool:
    return get_default_icms_rule(rule_profile, item) is not None


def has_configured_icms_rule(
    rule_profile: str,
    runtime_rules: list[dict[str, object]] | None,
    item: dict[str, object],
) -> bool:
    return get_configured_icms_rule(rule_profile, runtime_rules, item) is not None


def apply_default_icms_rule_actions(item: dict[str, object], rule: dict[str, object]) -> dict[str, object]:
    normalized_item = dict(item)
    sale_value = normalized_item["sale_value"] if isinstance(normalized_item.get("sale_value"), Decimal) else Decimal("0")
    quantity = normalized_item["quantity"] if isinstance(normalized_item.get("quantity"), Decimal) else Decimal("0")
    normalized_item["_configured_rule_signature"] = build_rule_signature(rule)

    new_cst = rule.get("new_cst")
    if isinstance(new_cst, str) and new_cst.strip():
        normalized_item["cst_icms"] = new_cst.strip()

    new_cfop = rule.get("new_cfop")
    if isinstance(new_cfop, str) and new_cfop.strip():
        normalized_item["cfop"] = new_cfop.strip()

    sale_value_factor = rule.get("sale_value_factor")
    if isinstance(sale_value_factor, Decimal):
        factored_sale_value = (sale_value * sale_value_factor).quantize(Decimal("0.01"))
        factored_quantity = (quantity * sale_value_factor).quantize(Decimal("0.01"))
        sale_value = factored_sale_value
        normalized_item["sale_value"] = factored_sale_value
        normalized_item["quantity"] = factored_quantity
        normalized_item["_sale_value_factor"] = sale_value_factor
        if "total_operation_value" in normalized_item:
            normalized_item["total_operation_value"] = factored_sale_value

    set_sale_value = rule.get("set_sale_value")
    if isinstance(set_sale_value, Decimal):
        sale_value = set_sale_value
        normalized_item["sale_value"] = set_sale_value
        if "total_operation_value" in normalized_item:
            normalized_item["total_operation_value"] = set_sale_value

    if rule.get("zero_icms"):
        normalized_item["base_icms"] = Decimal("0")
        normalized_item["icms_rate"] = Decimal("0")
        normalized_item["icms_value"] = Decimal("0")
        return normalized_item

    forced_rate = rule.get("force_rate")
    if isinstance(forced_rate, Decimal):
        normalized_item["icms_rate"] = forced_rate

    set_base_icms = rule.get("set_base_icms")
    if isinstance(set_base_icms, Decimal):
        normalized_item["base_icms"] = set_base_icms
    elif rule.get("use_sale_value_as_base"):
        normalized_item["base_icms"] = sale_value

    base_reduction_percent = rule.get("base_reduction_percent")
    if isinstance(base_reduction_percent, Decimal) and rule.get("recalculate_reduced_base"):
        current_base_icms = sale_value if sale_value > Decimal("0") else normalized_item["base_icms"] if isinstance(normalized_item.get("base_icms"), Decimal) else Decimal("0")
        reduction_factor = (Decimal("100") - base_reduction_percent) / Decimal("100")
        if reduction_factor < Decimal("0"):
            reduction_factor = Decimal("0")
        normalized_item["base_icms"] = (current_base_icms * reduction_factor).quantize(Decimal("0.01"))
        normalized_item["_base_reduction_percent"] = base_reduction_percent
        normalized_item["_document_recalculate_icms"] = True

    icms_per_quantity = rule.get("icms_per_quantity")
    if isinstance(icms_per_quantity, Decimal):
        normalized_item["base_icms"] = Decimal("0")
        normalized_item["icms_rate"] = Decimal("0")
        normalized_item["icms_value"] = (quantity * icms_per_quantity).quantize(Decimal("0.01"))
        if rule.get("preserve_icms_value"):
            normalized_item["_preserve_icms_value"] = True
        return normalized_item

    set_icms_value = rule.get("set_icms_value")
    if isinstance(set_icms_value, Decimal):
        normalized_item["_document_target_icms_total"] = set_icms_value.quantize(Decimal("0.01"))
        normalized_item["icms_value"] = set_icms_value
        normalized_item["_preserve_icms_value"] = True
        return normalized_item

    if rule.get("recalculate_icms_value"):
        normalized_item["_document_recalculate_icms"] = True

    base_icms = normalized_item["base_icms"] if isinstance(normalized_item.get("base_icms"), Decimal) else Decimal("0")
    icms_rate = normalized_item["icms_rate"] if isinstance(normalized_item.get("icms_rate"), Decimal) else Decimal("0")
    normalized_item["icms_value"] = (
        (base_icms * icms_rate / Decimal("100")).quantize(Decimal("0.01"))
        if base_icms > Decimal("0") and icms_rate > Decimal("0")
        else Decimal("0")
    )
    return normalized_item


def expand_default_icms_rule_items(
    rule_profile: str,
    item: dict[str, object],
) -> list[dict[str, object]]:
    rule = get_default_icms_rule(rule_profile, item)
    if not rule:
        return [dict(item)]

    split_rules = rule.get("split_into")
    if isinstance(split_rules, list):
        split_items: list[dict[str, object]] = []
        for split_rule in split_rules:
            if isinstance(split_rule, dict):
                split_items.append(apply_default_icms_rule_actions(item, split_rule))
        if split_items:
            return split_items

    return [apply_default_icms_rule_actions(item, rule)]


def expand_configured_icms_rule_items(
    rule_profile: str,
    runtime_rules: list[dict[str, object]] | None,
    item: dict[str, object],
) -> list[dict[str, object]]:
    rule = get_configured_icms_rule(rule_profile, runtime_rules, item)
    if not rule:
        return [dict(item)]

    split_rules = rule.get("split_into")
    if isinstance(split_rules, list):
        split_items: list[dict[str, object]] = []
        for split_rule in split_rules:
            if isinstance(split_rule, dict):
                split_items.append(apply_default_icms_rule_actions(item, split_rule))
        if split_items:
            return split_items

    return [apply_default_icms_rule_actions(item, rule)]


def apply_default_icms_rules(rule_profile: str, item: dict[str, object]) -> dict[str, object]:
    operation_type = str(item.get("operation_type", "")).strip().lower()
    if operation_type not in {"entrada", "saida"}:
        return dict(item)

    current_item = dict(item)
    for _ in range(5):
        rule = get_default_icms_rule(rule_profile, current_item)
        if not rule:
            break
        rule_signature = build_rule_signature(rule)
        if current_item.get("_last_applied_rule_signature") == rule_signature:
            break
        updated_item = apply_default_icms_rule_actions(current_item, rule)
        updated_item["_last_applied_rule_signature"] = rule_signature
        if updated_item == current_item:
            break
        current_item = updated_item
    return current_item


def apply_configured_icms_rules(
    rule_profile: str,
    runtime_rules: list[dict[str, object]] | None,
    item: dict[str, object],
) -> dict[str, object]:
    operation_type = str(item.get("operation_type", "")).strip().lower()
    if operation_type not in {"entrada", "saida"}:
        return dict(item)
    if item.get("_configured_rules_applied"):
        return dict(item)

    current_item = dict(item)
    for _ in range(5):
        rule = get_configured_icms_rule(rule_profile, runtime_rules, current_item)
        if not rule:
            break
        rule_signature = build_rule_signature(rule)
        if current_item.get("_last_applied_rule_signature") == rule_signature:
            break
        updated_item = apply_default_icms_rule_actions(current_item, rule)
        updated_item["_last_applied_rule_signature"] = rule_signature
        if updated_item == current_item:
            break
        current_item = updated_item
    current_item["_configured_rules_applied"] = True
    return current_item


def apply_sped_icms_consistency_rules(
    item: dict[str, object],
    rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    normalized_item = apply_configured_icms_rules(rule_profile, runtime_rules, item) if (rule_profile or runtime_rules) else dict(item)
    operation_type = str(normalized_item.get("operation_type", "")).strip().lower()
    cst_digits = normalize_cst_icms_for_sped(str(normalized_item.get("cst_icms", "")))
    cst_suffix = cst_digits[-2:]

    base_icms = normalized_item["base_icms"] if isinstance(normalized_item.get("base_icms"), Decimal) else Decimal("0")
    icms_rate = normalized_item["icms_rate"] if isinstance(normalized_item.get("icms_rate"), Decimal) else Decimal("0")
    icms_value = normalized_item["icms_value"] if isinstance(normalized_item.get("icms_value"), Decimal) else Decimal("0")

    if operation_type == "entrada" and cst_suffix in {"02", "15", "53", "61"}:
        normalized_item["base_icms"] = Decimal("0")
        if not normalized_item.get("_preserve_icms_value"):
            normalized_item["icms_value"] = Decimal("0")
        if cst_suffix == "61":
            normalized_item["icms_rate"] = Decimal("0")
    else:
        normalized_item["base_icms"] = base_icms
        normalized_item["icms_rate"] = icms_rate
        normalized_item["icms_value"] = icms_value

    return normalized_item
