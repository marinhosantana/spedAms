from __future__ import annotations

import re
from collections import defaultdict
from decimal import Decimal

from app.models import CompareSpedDocument
from app.parsers.compare_sheet import normalize_compare_document_number
from app.parsers.sped_parser import normalize_document_key, parse_decimal
from app.services.compare_sped_launcher import compare_register

def compare_decimal_value(value: object) -> Decimal:
    cleaned = re.sub(r"[^0-9,.\-]", "", str(value or "").strip())
    if not cleaned:
        return Decimal("0.00")
    return parse_decimal(cleaned).quantize(Decimal("0.01"))

def describe_compare_xml_sped_value_difference(invoice: object, document: object) -> str:
    xml_value = compare_decimal_value(getattr(invoice, "total_value", ""))
    sped_value = compare_decimal_value(getattr(document, "total_value", ""))
    difference = (sped_value - xml_value).quantize(Decimal("0.01"))
    if difference == 0:
        return "Valores iguais."
    xml_st = Decimal(str(getattr(invoice, "total_st", 0) or 0)).quantize(Decimal("0.01"))
    xml_ipi = Decimal(str(getattr(invoice, "total_ipi", 0) or 0)).quantize(Decimal("0.01"))
    if difference > 0 and xml_st == 0:
        return "SPED maior que XML. Conferir ICMS ST informado no SPED e ausente no total do XML."
    if difference < 0 and abs(abs(difference) - xml_ipi) <= Decimal("0.01"):
        return "SPED menor que XML. Diferença bate com IPI do XML que não entrou no VL_DOC do SPED."
    if difference < 0 and xml_ipi:
        return "SPED menor que XML. Conferir IPI/acréscimos do XML contra VL_DOC do SPED."
    return "Valor total do SPED diferente do XML. Conferir composição: produtos, desconto, frete, outros, ST e IPI."

def build_compare_sped_number_indexes(
    documents: dict[str, CompareSpedDocument],
) -> tuple[dict[tuple[str, str], CompareSpedDocument], dict[str, list[CompareSpedDocument]]]:
    documents_by_number_series: dict[tuple[str, str], CompareSpedDocument] = {}
    documents_by_number: dict[str, list[CompareSpedDocument]] = defaultdict(list)
    for document in documents.values():
        normalized_number = normalize_compare_document_number(document.number)
        normalized_series = normalize_compare_document_number(document.series)
        if not normalized_number:
            continue
        documents_by_number_series[(normalized_number, normalized_series)] = document
        documents_by_number[normalized_number].append(document)
    return documents_by_number_series, documents_by_number

def compare_sped_c100_has_generated_nfse_service(lines: list[str], start_index: int) -> bool:
    for child_line in lines[start_index + 1:]:
        register = compare_register(child_line)
        if register == "C100" or not register.startswith("C"):
            break
        fields = child_line.split("|")
        if register == "C170":
            cfop = fields[11].strip() if len(fields) > 11 else ""
            if cfop.endswith("933"):
                return True
        if register == "C190":
            cfop = fields[3].strip() if len(fields) > 3 else ""
            if cfop.endswith("933"):
                return True
        if register == "C195":
            observation_code = fields[2].strip().upper() if len(fields) > 2 else ""
            if observation_code.startswith("SP") and observation_code.endswith("933"):
                return True
    return False

def find_compare_sped_document_by_key_or_number(
    invoice: object,
    documents: dict[str, CompareSpedDocument],
    documents_by_number_series: dict[tuple[str, str], CompareSpedDocument],
    documents_by_number: dict[str, list[CompareSpedDocument]],
) -> tuple[CompareSpedDocument | None, str]:
    key = str(getattr(invoice, "key", "")).strip()
    document = documents.get(key)
    if document is not None:
        return document, "key"

    normalized_number = normalize_compare_document_number(getattr(invoice, "number", ""))
    if not normalized_number:
        return None, ""
    normalized_series = normalize_compare_document_number(getattr(invoice, "series", ""))
    issuer_cnpj = normalize_document_key(str(getattr(invoice, "issuer_cnpj", "") or ""))
    invoice_model = str(getattr(invoice, "model", "") or "").upper()
    if invoice_model.startswith("NFS") and issuer_cnpj:
        nfse_keys = [
            f"C100:NFSE:{issuer_cnpj}:NUM:{normalized_number}:SER:{normalize_compare_document_number('01')}",
            f"C100:NFSE:{issuer_cnpj}:NUM:{normalized_number}:SER:{normalized_series}",
            f"A100:NFSE:{issuer_cnpj}:NUM:{normalized_number}:SER:{normalized_series}",
        ]
        for nfse_key in nfse_keys:
            document = documents.get(nfse_key)
            if document is not None:
                return document, "nfse"
    document = documents_by_number_series.get((normalized_number, normalized_series))
    if document is None:
        candidates = documents_by_number.get(normalized_number, [])
        if issuer_cnpj:
            issuer_candidates = [
                candidate
                for candidate in candidates
                if normalize_document_key(candidate.issuer_cnpj) == issuer_cnpj
            ]
            if len(issuer_candidates) == 1:
                document = issuer_candidates[0]
        if document is None and len(candidates) == 1:
            document = candidates[0]
    if document is None:
        return None, ""
    return document, "number"
