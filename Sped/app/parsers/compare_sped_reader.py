from __future__ import annotations

from pathlib import Path
from typing import Callable

from app.config import COMPARE_KEY_PATTERN
from app.models import CompareSpedDocument
from app.parsers.compare_sheet import normalize_compare_document_number
from app.parsers.sped_parser import normalize_document_key
from app.services.compare_matching import compare_sped_c100_has_generated_nfse_service
from app.services.compare_operations import normalize_compare_operation_scope
from app.services.compare_sped_launcher import compare_register


def extract_company_tax_id_from_sped(sped_file: Path) -> str:
    with sped_file.open("r", encoding="latin-1", errors="ignore") as current_file:
        for raw_line in current_file:
            if not raw_line.startswith("|0000|"):
                continue
            fields = raw_line.rstrip("\r\n").split("|")
            cnpj = normalize_document_key(fields[7] if len(fields) > 7 else "")
            cpf = normalize_document_key(fields[8] if len(fields) > 8 else "")
            return cnpj or cpf
    return ""


def collect_compare_sped_documents(
    sped_file: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict[str, CompareSpedDocument]:
    documents: dict[str, CompareSpedDocument] = {}
    selected_scope = normalize_compare_operation_scope(operation_scope)
    with sped_file.open("r", encoding="latin-1", errors="ignore") as current_file:
        lines = current_file.readlines()
        participant_tax_id_by_code: dict[str, str] = {}
        for raw_line in lines:
            line = raw_line.strip()
            if compare_register(line) != "0150":
                continue
            fields = line.split("|")
            code = fields[2].strip() if len(fields) > 2 else ""
            cnpj = normalize_document_key(fields[5] if len(fields) > 5 else "")
            cpf = normalize_document_key(fields[6] if len(fields) > 6 else "")
            if code:
                participant_tax_id_by_code[code] = cnpj or cpf
        total_lines = max(len(lines), 1)
        if progress_callback:
            progress_callback(0, total_lines, f"Lendo SPED: {sped_file.name}")
        for line_index, raw_line in enumerate(lines, start=1):
            line = raw_line.strip()
            if progress_callback and (line_index == total_lines or line_index % 5000 == 0):
                progress_callback(line_index, total_lines, f"Lendo SPED {line_index}/{total_lines} linhas.")
            register = compare_register(line)
            if register not in {"C100", "A100"}:
                continue
            fields = line.split("|")
            operation_type = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
            if selected_scope != "Ambos" and operation_type != selected_scope:
                continue
            participant_code = fields[4].strip() if register == "C100" and len(fields) > 4 else ""
            participant_tax_id = participant_tax_id_by_code.get(participant_code, "")
            document_model = fields[5].strip() if len(fields) > 5 else ""
            key = fields[9].strip() if len(fields) > 9 else ""
            if register == "C100" and not (len(key) == 44 and key.isdigit()):
                match = COMPARE_KEY_PATTERN.search(line)
                key = match.group(0) if match else ""
            elif register == "A100":
                key = normalize_document_key(key) if normalize_document_key(key) else key
            if not key:
                if register not in {"A100", "C100"}:
                    continue
                number = fields[8].strip() if len(fields) > 8 else ""
                series = fields[6].strip() if register == "A100" and len(fields) > 6 else fields[7].strip() if len(fields) > 7 else ""
                normalized_number = normalize_compare_document_number(number)
                if not normalized_number:
                    continue
                normalized_series = normalize_compare_document_number(series)
                if register == "C100" and document_model == "01" and participant_tax_id:
                    key = f"{register}:NFSE:{participant_tax_id}:NUM:{normalized_number}:SER:{normalized_series}"
                else:
                    key = f"{register}:NUM:{normalized_number}:SER:{normalized_series}"
            if not key:
                continue
            documents[key] = CompareSpedDocument(
                key=key,
                number=fields[8].strip() if len(fields) > 8 else "",
                series=fields[6].strip() if register == "A100" and len(fields) > 6 else fields[7].strip() if len(fields) > 7 else "",
                issue_date=fields[10].strip() if len(fields) > 10 else "",
                issuer_cnpj=participant_tax_id,
                total_value=fields[12].strip() if len(fields) > 12 else "",
                file_path=str(sped_file),
                operation_type=operation_type,
                model=document_model,
                generated_nfse_service=(
                    register == "C100"
                    and document_model == "01"
                    and compare_sped_c100_has_generated_nfse_service(lines, line_index - 1)
                ),
            )
    return documents


def collect_compare_sped_keys(
    sped_file: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> set[str]:
    return set(collect_compare_sped_documents(sped_file, operation_scope, progress_callback).keys())
