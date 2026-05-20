from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable

from app.config import COMPARE_KEY_PATTERN
from app.models import CompareSheetInvoice
from app.parsers.compare_xml import compare_clean
from app.parsers.excel_parser import get_first_xlsx_sheet_name, read_xlsx_sheet_rows
from app.parsers.sped_parser import normalize_document_key
from app.services.compare_operations import normalize_compare_operation_scope


def compare_normalize_header_name(name: object) -> str:
    return re.sub(r"[^A-Z0-9]+", "", compare_clean(name).upper())


def normalize_compare_document_number(value: object) -> str:
    text = compare_clean(value).upper()
    digits = re.sub(r"\D+", "", text)
    if digits:
        return digits.lstrip("0") or "0"
    return re.sub(r"[^A-Z0-9]+", "", text).lstrip("0")


def compare_find_header_index(headers: dict[str, int], *aliases: str) -> int | None:
    for alias in aliases:
        index = headers.get(compare_normalize_header_name(alias))
        if index is not None:
            return index
    return None


def compare_find_key_in_values(values: list[object], preferred_index: int | None = None) -> re.Match[str] | None:
    if preferred_index is not None and preferred_index < len(values):
        match = COMPARE_KEY_PATTERN.search(compare_clean(values[preferred_index]))
        if match:
            return match
    for value in values:
        match = COMPARE_KEY_PATTERN.search(compare_clean(value))
        if match:
            return match
    return None


def convert_compare_xls_to_xlsx(sheet_file: Path) -> Path:
    temp_fd, temp_name = tempfile.mkstemp(prefix="compare_sheet_", suffix=".xlsx")
    os.close(temp_fd)
    temp_path = Path(temp_name)
    input_fd, input_name = tempfile.mkstemp(prefix="compare_sheet_input_", suffix=".xls")
    os.close(input_fd)
    temp_input_path = Path(input_name)
    with sheet_file.open("rb") as src_file, temp_input_path.open("wb") as dst_file:
        shutil.copyfileobj(src_file, dst_file)
    src = str(temp_input_path.resolve()).replace("'", "''")
    dst = str(temp_path.resolve()).replace("'", "''")
    script = (
        "$ErrorActionPreference='Stop'; "
        f"$src = '{src}'; "
        f"$dst = '{dst}'; "
        "$excel = $null; $wb = $null; $pv = $null; "
        "try { "
        "$excel = New-Object -ComObject Excel.Application; "
        "$excel.DisplayAlerts = $false; "
        "$excel.Visible = $false; "
        "$excel.AutomationSecurity = 3; "
        "$openErrors = @(); "
        "foreach ($repairMode in @(1, 2, 3)) { "
        "try { "
        "$wb = $excel.Workbooks.Open($src, 0, $true, 5, '', '', $true, 2, '', $false, $false, 0, $true, $true, $repairMode); "
        "break "
        "} catch { $openErrors += $_.Exception.Message } "
        "} "
        "if ($wb -eq $null) { "
        "try { "
        "$pv = $excel.ProtectedViewWindows.Open($src); "
        "$wb = $pv.Edit() "
        "} catch { "
        "$openErrors += $_.Exception.Message; "
        "throw ($openErrors -join ' | ') "
        "} "
        "} "
        "$wb.SaveAs($dst, 51) "
        "} finally { "
        "if ($wb -ne $null) { $wb.Close($false) }; "
        "if ($pv -ne $null) { $pv.Close() }; "
        "if ($excel -ne $null) { $excel.Quit() } "
        "}"
    )
    try:
        startupinfo = None
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=True,
            capture_output=True,
            text=True,
            startupinfo=startupinfo,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.CalledProcessError as exc:
        temp_path.unlink(missing_ok=True)
        detail = (exc.stderr or exc.stdout or str(exc)).strip()
        raise RuntimeError(
            "Nao foi possivel abrir a planilha .xls automaticamente. "
            "Se o arquivo veio da internet/WhatsApp/e-mail, clique com o botao direito nele, "
            "abra Propriedades e marque Desbloquear, ou abra no Excel e salve como .xlsx.\n\n"
            f"Detalhe: {detail}"
        )
    finally:
        temp_input_path.unlink(missing_ok=True)
    return temp_path


def format_compare_sheet_date(value: object) -> str:
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)) and value > 20000:
        try:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(value))).strftime("%Y-%m-%d")
        except (OverflowError, ValueError):
            pass
    text = compare_clean(value)
    if re.fullmatch(r"\d+(?:[.,]\d+)?", text):
        try:
            serial = float(text.replace(",", "."))
            if serial > 20000:
                return (dt.datetime(1899, 12, 30) + dt.timedelta(days=serial)).strftime("%Y-%m-%d")
        except (OverflowError, ValueError):
            pass
    return text[:-9] if text.endswith(" 00:00:00") else text


def collect_compare_sheet_invoices(
    sheet_file: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareSheetInvoice], int]:
    effective_file = sheet_file
    temp_file: Path | None = None
    sheet_suffix = sheet_file.suffix.lower()
    if sheet_suffix not in {".xlsx", ".xls", ".xlsm", ".xlt", ".xltx", ".xltm"}:
        raise ValueError(
            "Formato de planilha nao suportado. Use .xlsx, .xls, .xlsm, .xlt, .xltx ou .xltm."
        )
    if sheet_suffix in {".xls", ".xlt"}:
        if progress_callback:
            progress_callback(0, 100, "Convertendo planilha antiga para leitura.")
        temp_file = convert_compare_xls_to_xlsx(sheet_file)
        effective_file = temp_file
    try:
        if progress_callback:
            progress_callback(10, 100, f"Abrindo planilha: {sheet_file.name}")
        rows = read_xlsx_sheet_rows(effective_file, get_first_xlsx_sheet_name(effective_file))
    finally:
        if temp_file is not None:
            temp_file.unlink(missing_ok=True)
    if not rows:
        return [], 0

    headers = {compare_normalize_header_name(name): index for index, name in enumerate(rows[0])}
    key_index = compare_find_header_index(headers, "Chave", "CHAVE_ACESSO", "CHAVE_ACESS0", "CHNFE", "CHAVE NFE", "NFE_CHAVE")
    number_index = compare_find_header_index(headers, "Numero", "NUM_NOTA", "NOTA", "NRO_NOTA", "NUMERO")
    series_index = compare_find_header_index(headers, "Serie", "SERIE")
    issue_date_index = compare_find_header_index(headers, "Dt_Emissao", "DT_EMISSAO", "DATA_EMISSAO", "EMISSAO")
    issuer_index = compare_find_header_index(
        headers,
        "CNPJ_CPF_CnpjEmit",
        "CNPJ_EMIT",
        "CNPJ Emitente",
        "CNPJ_CPF_EMIT",
        "CNPJFORNECEDOR",
        "CNPJ_FORNECEDOR",
        "CNPJ",
        "COD_FORNECEDOR",
    )
    value_index = compare_find_header_index(
        headers,
        "Valor",
        "Total Operacao",
        "Valor Operacao",
        "VL_NOTA",
        "VALOR",
        "VLR_NOTA",
        "VR_TOTAL",
        "VRTOTAL",
        "TOTAL",
    )
    operation_index = compare_find_header_index(
        headers,
        "Tipo Movimento",
        "Tipo_Movimento",
        "Entrada Saida",
        "Tipo Operacao",
        "Operacao",
        "Tipo",
    )

    invoices: list[CompareSheetInvoice] = []
    ignored = 0
    data_rows = rows[1:]
    total_rows = max(len(data_rows), 1)
    for row_index, values in enumerate(data_rows, start=1):
        if not any(compare_clean(value) for value in values):
            continue

        def get_value(index: int | None) -> object:
            if index is None or index >= len(values):
                return ""
            return values[index]

        match = compare_find_key_in_values(values, key_index)
        key_value = match.group(0) if match else ""
        if not key_value and key_index is not None:
            candidate_key = compare_clean(get_value(key_index))
            normalized_candidate_key = normalize_document_key(candidate_key)
            if len(normalized_candidate_key) >= 20:
                key_value = normalized_candidate_key
            elif len(candidate_key) >= 20:
                key_value = candidate_key
        if not key_value:
            ignored += 1
            continue
        raw_operation_type = compare_clean(get_value(operation_index))
        operation_type = normalize_compare_operation_scope(raw_operation_type) if raw_operation_type else ""
        issuer_cnpj = re.sub(r"\D+", "", compare_clean(get_value(issuer_index)))
        if not issuer_cnpj and len(key_value) == 44:
            issuer_cnpj = key_value[6:20]
        invoices.append(
            CompareSheetInvoice(
                key=key_value,
                number=compare_clean(get_value(number_index)),
                series=compare_clean(get_value(series_index)),
                issue_date=format_compare_sheet_date(get_value(issue_date_index)),
                issuer_cnpj=issuer_cnpj,
                total_value=compare_clean(get_value(value_index)),
                file_path=str(sheet_file),
                operation_type=operation_type,
            )
        )
        if progress_callback and (row_index == total_rows or row_index % 500 == 0):
            progress_callback(row_index, total_rows, f"Lendo planilha {row_index}/{total_rows} linhas.")
    return invoices, ignored
