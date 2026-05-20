from __future__ import annotations

import datetime as dt
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import tkinter as tk
import csv
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Callable, Iterable
from tkinter import BOTH, END, LEFT, MULTIPLE, RIGHT, BooleanVar, Canvas, Listbox, Menu, StringVar, Text, Tk, Toplevel
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from app.config import (
    APP_DEFAULT_CONFIG,
    COMPARE_KEY_PATTERN,
    COMPARE_MARK_CHECKED,
    COMPARE_MARK_UNCHECKED,
    COMPARE_NS_NFE,
    COMPARE_NS_NFSE,
)
from app.exporters.excel_base import (
    excel_column_name,
    xml_escape,
    build_cell,
    build_sheet_xml,
    sanitize_excel_sheet_name,
    normalize_excel_header_text,
    is_excel_numeric_value,
    decimal_from_excel_value,
    detect_excel_column_kind,
    should_total_excel_column,
    body_style_for_column_kind,
    total_style_for_column_kind,
    build_sheet_rows_with_metadata,
    build_excel_styles_xml,
    EXCEL_STYLE_DEFAULT,
    EXCEL_STYLE_HEADER,
    EXCEL_STYLE_CURRENCY,
    EXCEL_STYLE_PERCENT,
    EXCEL_STYLE_NUMBER,
    EXCEL_STYLE_TOTAL_LABEL,
    EXCEL_STYLE_TOTAL_CURRENCY,
    EXCEL_STYLE_TOTAL_PERCENT,
    EXCEL_STYLE_TOTAL_NUMBER,
    EXCEL_STYLE_HEADER_BLUE,
    EXCEL_STYLE_HEADER_YELLOW,
    EXCEL_STYLE_BODY_BLUE,
    EXCEL_STYLE_BODY_YELLOW,
    EXCEL_STYLE_CURRENCY_BLUE,
    EXCEL_STYLE_CURRENCY_YELLOW,
    EXCEL_STYLE_PERCENT_BLUE,
    EXCEL_STYLE_PERCENT_YELLOW,
    EXCEL_STYLE_NUMBER_BLUE,
    EXCEL_STYLE_NUMBER_YELLOW,
)
from app.exporters.workbook_exporter import (
    write_simple_excel_workbook,
    build_month_color_style_maps,
    write_monthly_colored_excel_workbook,
    write_monthly_colored_excel_workbook_with_sheets,
    write_simple_csv_file,
    serialize_value_for_clipboard,
)
from app.models import (
    CompareInvestigationMatchedRow,
    CompareSheetInvoice,
    CompareSpedDocument,
    CompareXmlInvoice,
    CompareXmlItem,
    ProductRecord,
)
from app.parsers.sped_parser import (
    first_non_empty,
    get_field,
    infer_operation_type_from_cfop,
    normalize_cfop,
    normalize_document_key,
    normalize_sped_line,
    parse_decimal,
    parse_rate,
    read_combined_e110_summary,
    read_sped_e110_summary,
    read_sped_plain_lines,
    read_sped_summary_register_rows,
)
from app.parsers.excel_parser import (
    parse_excel_shared_strings,
    extract_excel_cell_value,
    read_xlsx_sheet_rows,
    read_xlsx_sheet,
    get_first_xlsx_sheet_name,
    read_filter_descriptions_file,
    parse_filter_values,
    has_active_item_filters,
)
from app.parsers.compare_xml import (
    compare_clean,
    compare_sanitize,
    compare_to_float,
    compare_extract_money_values,
    compare_format_float,
    compare_xml_text,
    compare_nfse_text,
    compare_extract_key,
    compare_extract_nfse_key,
    compare_is_nfse_invoice,
    compare_date_to_sped,
    compare_extract_icms,
    compare_extract_pis_cofins,
    compare_extract_ipi,
    parse_compare_nfse_file,
    parse_compare_xml_file,
    describe_compare_ignored_xml,
    extract_xml_cancellation_event,
    collect_xml_cancellation_events,
    collect_compare_ignored_xml_rows,
)
from app.services.product_import import (
    build_import_products_from_xml_sources,
    build_product_origin_candidates_from_xml_sources,
    collect_cest_values,
    collect_xml_files,
    format_cest_values,
    get_xml_local_name,
    local_name_text,
    ncm_matches_filters,
    normalize_ncm,
    parse_nfce_xml_items,
    read_sped_contrib_product_rows,
    read_sped_contrib_detailed_rows,
    summarize_pis_cofins_analysis,
    build_import_products_from_sped_contrib_sources,
    build_product_origin_candidates_from_sped_file,
    build_import_products_from_sped_0200,
    build_import_products_from_sped_fiscal_sources,
    build_import_products_from_consolidated_sources,
)
from app.services.tax_rules import (
    build_pis_cofins_side_values,
    compute_display_icms_rate,
    format_decimal_sped,
    has_icms_reduction,
    merge_product_cst,
    merge_product_rate,
    normalize_cst_icms_for_sped,
    normalize_header,
    normalize_operation_type,
    normalize_tax_code,
    normalize_text,
    resolve_header,
)
from app.services.compare_sped_launcher import (
    compare_register,
    compare_line,
    find_compare_index,
    compare_find_c100_by_key,
    compare_find_existing_invoice,
    compare_next_code,
    compare_ensure_0150,
    compare_ensure_0190,
    compare_ensure_0200,
    build_compare_c100,
    build_compare_nfse_c170,
    build_compare_c170,
    build_compare_c190,
    recalc_compare_block_counts,
    recalc_compare_9900,
    recalc_compare_all,
    build_compare_updated_path,
    find_compare_c100_insert_index,
    add_compare_invoice_to_lines,
    launch_compare_invoices_in_sped,
    launch_compare_invoice_in_sped,
)
from app.repositories.mysql_cadastro import MysqlCadastroRepository


def collect_compare_xml_invoices(
    xml_folder: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareXmlInvoice], int]:
    invoices: list[CompareXmlInvoice] = []
    ignored = 0
    xml_paths = sorted(xml_folder.rglob("*.xml"))
    total_xml = max(len(xml_paths), 1)
    if progress_callback:
        progress_callback(0, total_xml, f"Localizados {len(xml_paths)} XML(s).")
    for index, xml_path in enumerate(xml_paths, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
        else:
            invoices.append(invoice)
        if progress_callback and (index == total_xml or index % 10 == 0):
            progress_callback(index, total_xml, f"Lendo XML {index}/{total_xml}: {xml_path.name}")
    return invoices, ignored


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


def classify_xml_item_operation(invoice: CompareXmlInvoice, item: CompareXmlItem, company_tax_id: str = "") -> str:
    company_key = normalize_document_key(company_tax_id)
    issuer_key = normalize_document_key(invoice.issuer_cnpj)
    recipient_key = normalize_document_key(invoice.recipient_cnpj)
    model = str(invoice.model or "").strip()
    if model == "65":
        return "Saida"
    if model.upper().startswith("NFS"):
        if company_key:
            if issuer_key == company_key:
                return "Saida"
            if recipient_key == company_key:
                return "Entrada"
        return "Entrada"
    cfop = str(item.cfop or "").strip()
    is_outgoing_cfop = cfop.startswith(("5", "6"))

    if company_key:
        if issuer_key == company_key:
            return "Saida" if is_outgoing_cfop else "Entrada"
        if recipient_key == company_key:
            return "Entrada"
    return "Saida" if is_outgoing_cfop else "Entrada"


def xml_contains_company_tax_id(xml_path: Path, company_tax_id: str) -> bool:
    company_key = normalize_document_key(company_tax_id)
    if not company_key:
        return False
    try:
        text = xml_path.read_text(encoding="utf-8", errors="ignore")
    except UnicodeDecodeError:
        text = xml_path.read_text(encoding="latin-1", errors="ignore")
    return company_key in normalize_document_key(text)


def build_xml_cfop_summary_rows(
    xml_source: Path,
    company_tax_id: str = "",
    operation_scope: str = "Todos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[dict[str, object]], dict[str, object]]:
    xml_files = [xml_source] if xml_source.is_file() else sorted(xml_source.rglob("*.xml"))
    total_files = max(len(xml_files), 1)
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    ignored = 0
    skipped_by_company = 0
    ignored_log_rows: list[dict[str, object]] = []
    document_log_rows: list[dict[str, object]] = []
    cancelled_display_rows: list[dict[str, object]] = []
    document_keys: set[str] = set()
    normalized_company_tax_id = normalize_document_key(company_tax_id)
    normalized_operation_scope = normalize_compare_operation_scope(operation_scope)
    cancellation_events = collect_xml_cancellation_events(xml_files)
    cancelled_documents = 0
    cancelled_value_total = Decimal("0")

    if progress_callback:
        progress_callback(0, total_files, f"Localizados {len(xml_files)} XML(s).")

    for index, xml_path in enumerate(xml_files, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
            ignored_log_rows.append(
                {
                    "reason": describe_compare_ignored_xml(xml_path),
                    "company_tax_id": normalized_company_tax_id,
                    "key": "",
                    "number": "",
                    "series": "",
                    "issue_date": "",
                    "issuer_cnpj": "",
                    "issuer_name": "",
                    "recipient_cnpj": "",
                    "recipient_name": "",
                    "total_value": Decimal("0"),
                    "file_path": str(xml_path),
                }
            )
        else:
            issuer_key = normalize_document_key(invoice.issuer_cnpj)
            recipient_key = normalize_document_key(invoice.recipient_cnpj)
            company_found_in_xml = (
                not normalized_company_tax_id
                or normalized_company_tax_id in {issuer_key, recipient_key}
                or xml_contains_company_tax_id(xml_path, normalized_company_tax_id)
            )
            if normalized_company_tax_id and not company_found_in_xml:
                skipped_by_company += 1
                ignored_log_rows.append(
                    {
                        "reason": "Fora da empresa",
                        "company_tax_id": normalized_company_tax_id,
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "issuer_cnpj": issuer_key,
                        "issuer_name": invoice.issuer_name,
                        "recipient_cnpj": recipient_key,
                        "recipient_name": invoice.recipient_name,
                        "total_value": Decimal(str(invoice.total_doc or 0)),
                        "file_path": str(xml_path),
                    }
                )
                if progress_callback and (index == total_files or index % 10 == 0):
                    progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
                continue
            if invoice.key:
                document_keys.add(invoice.key)
            cancellation_event = cancellation_events.get(invoice.key)
            if cancellation_event is not None:
                cancelled_documents += 1
                cancelled_value_total += Decimal(str(invoice.total_doc or 0))
                cancelled_cfops = sorted({str(item.cfop or "").strip() for item in invoice.items if str(item.cfop or "").strip()})
                cancelled_display_rows.append(
                    {
                        "operation_type": "Cancelado",
                        "cfop": " | ".join(cancelled_cfops) or "(sem CFOP)",
                        "document_count": 1,
                        "items": len(invoice.items),
                        "operation_value": Decimal(str(invoice.total_products or invoice.total_doc or 0)),
                        "discount_value": Decimal(str(invoice.total_discount or 0)),
                        "base_icms": Decimal(str(invoice.total_bc_icms or 0)),
                        "icms_value": Decimal(str(invoice.total_icms or 0)),
                        "base_icms_st": Decimal(str(invoice.total_bc_st or 0)),
                        "icms_st_value": Decimal(str(invoice.total_st or 0)),
                        "ipi_value": Decimal(str(invoice.total_ipi or 0)),
                        "base_pis": Decimal("0"),
                        "pis_value": Decimal(str(invoice.total_pis or 0)),
                        "base_cofins": Decimal("0"),
                        "cofins_value": Decimal(str(invoice.total_cofins or 0)),
                        "details": [],
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "file_path": str(xml_path),
                        "cancel_reason": cancellation_event.get("reason", ""),
                    }
                )
                ignored_log_rows.append(
                    {
                        "reason": f"NF-e/NFC-e cancelada: {cancellation_event.get('reason', '')}",
                        "company_tax_id": normalized_company_tax_id,
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "issuer_cnpj": issuer_key,
                        "issuer_name": invoice.issuer_name,
                        "recipient_cnpj": recipient_key,
                        "recipient_name": invoice.recipient_name,
                        "total_value": Decimal(str(invoice.total_doc or 0)),
                        "file_path": str(xml_path),
                    }
                )
                if progress_callback and (index == total_files or index % 10 == 0):
                    progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
                continue
            invoice_operation_types = sorted(
                {
                    classify_xml_item_operation(invoice, item, normalized_company_tax_id)
                    for item in invoice.items
                }
            )
            if not invoice_operation_types:
                invoice_operation_types = [infer_compare_invoice_operation(invoice, normalized_company_tax_id)]
            document_log_rows.append(
                {
                    "operation_type": " | ".join(operation for operation in invoice_operation_types if operation),
                    "key": invoice.key,
                    "number": invoice.number,
                    "series": invoice.series,
                    "issue_date": invoice.issue_date,
                    "issuer_cnpj": issuer_key,
                    "issuer_name": invoice.issuer_name,
                    "recipient_cnpj": recipient_key,
                    "recipient_name": invoice.recipient_name,
                    "total_value": Decimal(str(invoice.total_doc or 0)),
                    "item_count": len(invoice.items),
                    "file_path": str(xml_path),
                }
            )
            for item in invoice.items:
                cfop = str(item.cfop or "").strip() or "(sem CFOP)"
                operation_type = classify_xml_item_operation(invoice, item, normalized_company_tax_id)
                if (
                    normalized_company_tax_id
                    and normalized_company_tax_id not in {issuer_key, recipient_key}
                    and company_found_in_xml
                ):
                    operation_type = "Entrada"
                if normalized_operation_scope != "Ambos" and operation_type != normalized_operation_scope:
                    continue
                key = (operation_type, cfop)
                bucket = grouped_rows.setdefault(
                    key,
                    {
                        "operation_type": operation_type,
                        "cfop": cfop,
                        "documents": set(),
                        "items": 0,
                        "operation_value": Decimal("0"),
                        "discount_value": Decimal("0"),
                        "base_icms": Decimal("0"),
                        "icms_value": Decimal("0"),
                        "base_icms_st": Decimal("0"),
                        "icms_st_value": Decimal("0"),
                        "ipi_value": Decimal("0"),
                        "base_pis": Decimal("0"),
                        "pis_value": Decimal("0"),
                        "base_cofins": Decimal("0"),
                        "cofins_value": Decimal("0"),
                        "details": [],
                    },
                )
                if invoice.key:
                    bucket["documents"].add(invoice.key)
                bucket["items"] += 1
                detail_row = {
                    "operation_type": operation_type,
                    "cfop": cfop,
                    "key": invoice.key,
                    "number": invoice.number,
                    "series": invoice.series,
                    "issue_date": invoice.issue_date,
                    "issuer_cnpj": issuer_key,
                    "issuer_name": invoice.issuer_name,
                    "recipient_cnpj": recipient_key,
                    "recipient_name": invoice.recipient_name,
                    "item_no": item.item_no,
                    "code": item.code,
                    "description": item.description,
                    "operation_value": Decimal(str(item.value or 0)),
                    "discount_value": Decimal(str(item.discount or 0)),
                    "base_icms": Decimal(str(item.vl_bc_icms or 0)),
                    "icms_value": Decimal(str(item.vl_icms or 0)),
                    "base_icms_st": Decimal(str(item.vl_bc_icms_st or 0)),
                    "icms_st_value": Decimal(str(item.vl_icms_st or 0)),
                    "ipi_value": Decimal(str(item.vl_ipi or 0)),
                    "base_pis": Decimal(str(item.vl_bc_pis or 0)),
                    "pis_value": Decimal(str(item.vl_pis or 0)),
                    "base_cofins": Decimal(str(item.vl_bc_cofins or 0)),
                    "cofins_value": Decimal(str(item.vl_cofins or 0)),
                    "file_path": str(xml_path),
                }
                bucket["details"].append(detail_row)
                bucket["operation_value"] += Decimal(str(item.value or 0))
                bucket["discount_value"] += Decimal(str(item.discount or 0))
                bucket["base_icms"] += Decimal(str(item.vl_bc_icms or 0))
                bucket["icms_value"] += Decimal(str(item.vl_icms or 0))
                bucket["base_icms_st"] += Decimal(str(item.vl_bc_icms_st or 0))
                bucket["icms_st_value"] += Decimal(str(item.vl_icms_st or 0))
                bucket["ipi_value"] += Decimal(str(item.vl_ipi or 0))
                bucket["base_pis"] += Decimal(str(item.vl_bc_pis or 0))
                bucket["pis_value"] += Decimal(str(item.vl_pis or 0))
                bucket["base_cofins"] += Decimal(str(item.vl_bc_cofins or 0))
                bucket["cofins_value"] += Decimal(str(item.vl_cofins or 0))
        if progress_callback and (index == total_files or index % 10 == 0):
            progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")

    rows: list[dict[str, object]] = []
    for bucket in grouped_rows.values():
        documents = bucket.pop("documents")
        row = dict(bucket)
        row["document_count"] = len(documents)
        rows.append(row)
    rows.sort(key=lambda row: (str(row.get("operation_type", "")), str(row.get("cfop", ""))))
    stats = {
        "xml_total": len(xml_files),
        "xml_valid": len(document_keys),
        "xml_ignored": ignored,
        "xml_skipped_by_company": skipped_by_company,
        "xml_cancelled": cancelled_documents,
        "xml_cancelled_value_total": cancelled_value_total.quantize(Decimal("0.01")),
        "company_tax_id": normalized_company_tax_id,
        "ignored_log_rows": ignored_log_rows,
        "document_log_rows": document_log_rows,
        "cancellation_events": list(cancellation_events.values()),
        "cancelled_display_rows": cancelled_display_rows,
    }
    return rows, stats


def normalize_compare_operation_scope(value: object) -> str:
    normalized = normalize_text(str(value))
    if normalized in {"E", "ENT", "ENTR"} or normalized.startswith("ENTRADA"):
        return "Entrada"
    if normalized in {"S", "SAI"} or normalized.startswith("SAIDA"):
        return "Saida"
    return "Ambos"


def filter_xml_summary_rows_by_scope(rows: list[dict[str, object]], operation_scope: object) -> list[dict[str, object]]:
    selected_scope = normalize_compare_operation_scope(operation_scope)
    if selected_scope == "Ambos":
        return list(rows)
    return [
        row
        for row in rows
        if normalize_compare_operation_scope(row.get("operation_type", "")) == selected_scope
    ]


def build_xml_entry_credit_rows(
    xml_source: Path,
    company_tax_id: str = "",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    xml_files = [xml_source] if xml_source.is_file() else sorted(xml_source.rglob("*.xml"))
    total_files = max(len(xml_files), 1)
    normalized_company_tax_id = normalize_document_key(company_tax_id)
    invoice_rows: list[dict[str, object]] = []
    grouped_cfop: dict[str, dict[str, object]] = {}
    ignored = 0
    skipped_by_company = 0
    cancellation_events = collect_xml_cancellation_events(xml_files)
    cancelled_documents = 0
    cancelled_value_total = Decimal("0")

    if progress_callback:
        progress_callback(0, total_files, f"Localizados {len(xml_files)} XML(s).")

    for index, xml_path in enumerate(xml_files, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        issuer_key = normalize_document_key(invoice.issuer_cnpj)
        recipient_key = normalize_document_key(invoice.recipient_cnpj)
        company_found_in_xml = (
            not normalized_company_tax_id
            or normalized_company_tax_id in {issuer_key, recipient_key}
            or xml_contains_company_tax_id(xml_path, normalized_company_tax_id)
        )
        if normalized_company_tax_id and not company_found_in_xml:
            skipped_by_company += 1
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue
        if cancellation_events.get(invoice.key) is not None:
            cancelled_documents += 1
            cancelled_value_total += Decimal(str(invoice.total_doc or 0))
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        entry_items: list[CompareXmlItem] = []
        for item in invoice.items:
            operation_type = classify_xml_item_operation(invoice, item, normalized_company_tax_id)
            if (
                normalized_company_tax_id
                and normalized_company_tax_id not in {issuer_key, recipient_key}
                and company_found_in_xml
            ):
                operation_type = "Entrada"
            if operation_type == "Entrada":
                entry_items.append(item)

        if not entry_items:
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        details: list[dict[str, object]] = []
        cfops: set[str] = set()
        total_operation = Decimal("0")
        total_discount = Decimal("0")
        total_base_icms = Decimal("0")
        total_icms = Decimal("0")
        total_base_icms_st = Decimal("0")
        total_icms_st = Decimal("0")
        total_ipi = Decimal("0")

        for item in entry_items:
            cfop = str(item.cfop or "").strip() or "(sem CFOP)"
            cfops.add(cfop)
            item_operation = Decimal(str(item.value or 0))
            item_discount = Decimal(str(item.discount or 0))
            item_base_icms = Decimal(str(item.vl_bc_icms or 0))
            item_icms = Decimal(str(item.vl_icms or 0))
            item_base_icms_st = Decimal(str(item.vl_bc_icms_st or 0))
            item_icms_st = Decimal(str(item.vl_icms_st or 0))
            item_ipi = Decimal(str(item.vl_ipi or 0))
            detail = {
                "key": invoice.key,
                "number": invoice.number,
                "series": invoice.series,
                "issue_date": invoice.issue_date,
                "issuer_cnpj": issuer_key,
                "issuer_name": invoice.issuer_name,
                "recipient_cnpj": recipient_key,
                "recipient_name": invoice.recipient_name,
                "item_no": item.item_no,
                "code": item.code,
                "description": item.description,
                "ncm": item.ncm,
                "cfop": cfop,
                "cst_icms": item.cst_icms,
                "quantity": Decimal(str(item.quantity or 0)),
                "operation_value": item_operation,
                "discount_value": item_discount,
                "base_icms": item_base_icms,
                "icms_rate": Decimal(str(item.aliq_icms or 0)),
                "icms_value": item_icms,
                "base_icms_st": item_base_icms_st,
                "icms_st_value": item_icms_st,
                "ipi_value": item_ipi,
                "file_path": str(xml_path),
            }
            details.append(detail)
            total_operation += item_operation
            total_discount += item_discount
            total_base_icms += item_base_icms
            total_icms += item_icms
            total_base_icms_st += item_base_icms_st
            total_icms_st += item_icms_st
            total_ipi += item_ipi

            bucket = grouped_cfop.setdefault(
                cfop,
                {
                    "cfop": cfop,
                    "documents": set(),
                    "items": 0,
                    "operation_value": Decimal("0"),
                    "discount_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "base_icms_st": Decimal("0"),
                    "icms_st_value": Decimal("0"),
                    "ipi_value": Decimal("0"),
                    "details": [],
                },
            )
            if invoice.key:
                bucket["documents"].add(invoice.key)
            bucket["items"] += 1
            bucket["operation_value"] += item_operation
            bucket["discount_value"] += item_discount
            bucket["base_icms"] += item_base_icms
            bucket["icms_value"] += item_icms
            bucket["base_icms_st"] += item_base_icms_st
            bucket["icms_st_value"] += item_icms_st
            bucket["ipi_value"] += item_ipi
            bucket["details"].append(detail)

        invoice_rows.append(
            {
                "key": invoice.key,
                "number": invoice.number,
                "series": invoice.series,
                "issue_date": invoice.issue_date,
                "issuer_cnpj": issuer_key,
                "issuer_name": invoice.issuer_name,
                "recipient_cnpj": recipient_key,
                "recipient_name": invoice.recipient_name,
                "cfops": " | ".join(sorted(cfops)),
                "items": len(entry_items),
                "operation_value": total_operation,
                "discount_value": total_discount,
                "base_icms": total_base_icms,
                "icms_value": total_icms,
                "base_icms_st": total_base_icms_st,
                "icms_st_value": total_icms_st,
                "ipi_value": total_ipi,
                "file_path": str(xml_path),
                "details": details,
            }
        )

        if progress_callback and (index == total_files or index % 10 == 0):
            progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")

    cfop_rows: list[dict[str, object]] = []
    for bucket in grouped_cfop.values():
        documents = bucket.pop("documents")
        row = dict(bucket)
        row["document_count"] = len(documents)
        cfop_rows.append(row)
    invoice_rows.sort(key=lambda row: (str(row.get("issue_date", "")), str(row.get("number", "")), str(row.get("series", ""))))
    cfop_rows.sort(key=lambda row: str(row.get("cfop", "")))
    stats = {
        "xml_total": len(xml_files),
        "xml_valid_entries": len(invoice_rows),
        "xml_ignored": ignored,
        "xml_skipped_by_company": skipped_by_company,
        "xml_cancelled": cancelled_documents,
        "xml_cancelled_value_total": cancelled_value_total.quantize(Decimal("0.01")),
        "company_tax_id": normalized_company_tax_id,
        "cancellation_events": list(cancellation_events.values()),
    }
    return invoice_rows, cfop_rows, stats


def infer_compare_invoice_operation(invoice: CompareXmlInvoice, company_tax_id: str = "") -> str:
    operations = {
        classify_xml_item_operation(invoice, item, company_tax_id)
        for item in invoice.items
    }
    operations.discard("")
    if len(operations) == 1:
        return next(iter(operations))
    if len(operations) > 1:
        return "Entrada/Saida"
    return classify_xml_item_operation(
        invoice,
        CompareXmlItem(
            item_no="",
            code="",
            description="",
            unit="",
            ncm="",
            cfop="",
            quantity=0.0,
            value=0.0,
            discount=0.0,
            cst_icms="",
            vl_bc_icms=0.0,
            aliq_icms=0.0,
            vl_icms=0.0,
            vl_bc_icms_st=0.0,
            aliq_icms_st=0.0,
            vl_icms_st=0.0,
            vl_ipi=0.0,
            cst_pis="",
            vl_bc_pis=0.0,
            aliq_pis=0.0,
            vl_pis=0.0,
            cst_cofins="",
            vl_bc_cofins=0.0,
            aliq_cofins=0.0,
            vl_cofins=0.0,
        ),
        company_tax_id,
    )


def build_compare_invoice_note_snapshot(invoice: CompareXmlInvoice) -> dict[str, object]:
    items: list[dict[str, object]] = []
    for item in invoice.items:
        sale_value = Decimal(str(item.value or 0)).quantize(Decimal("0.01"))
        ipi_value = Decimal(str(item.vl_ipi or 0)).quantize(Decimal("0.01"))
        icms_st_value = Decimal(str(item.vl_icms_st or 0)).quantize(Decimal("0.01"))
        total_operation_value = (sale_value + ipi_value + icms_st_value).quantize(Decimal("0.01"))
        items.append(
            {
                "item_number": item.item_no,
                "code": item.code,
                "description": item.description,
                "ncm": item.ncm,
                "cest": "",
                "cfop": item.cfop,
                "cst_icms": item.cst_icms,
                "icms_rate": Decimal(str(item.aliq_icms or 0)).quantize(Decimal("0.01")),
                "quantity": Decimal(str(item.quantity or 0)),
                "sale_value": sale_value,
                "total_operation_value": total_operation_value,
                "discount_value": Decimal(str(item.discount or 0)).quantize(Decimal("0.01")),
                "ipi_value": ipi_value,
                "base_icms": Decimal(str(item.vl_bc_icms or 0)).quantize(Decimal("0.01")),
                "icms_value": Decimal(str(item.vl_icms or 0)).quantize(Decimal("0.01")),
                "icms_st_value": icms_st_value,
                "source_register": "XML",
            }
        )
    return {
        "period": "",
        "document_date": compare_date_to_sped(invoice.issue_date) or invoice.issue_date,
        "document_number": invoice.number,
        "document_series": invoice.series,
        "document_model": invoice.model,
        "document_key": invoice.key,
        "participant_code": "",
        "participant_name": invoice.issuer_name,
        "participant_tax_id": normalize_document_key(invoice.issuer_cnpj),
        "items": items,
    }


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


def collect_compare_fiscal_basis_totals(sped_file: Path, operation_scope: str = "Ambos") -> dict[str, object]:
    selected_scope = normalize_compare_operation_scope(operation_scope)
    totals: dict[str, Decimal] = {
        "c100_total": Decimal("0.00"),
        "a100_total": Decimal("0.00"),
        "c190_total": Decimal("0.00"),
        "c590_total": Decimal("0.00"),
        "d590_total": Decimal("0.00"),
    }
    counts: dict[str, int] = {
        "c100_count": 0,
        "a100_count": 0,
        "c190_count": 0,
        "c590_count": 0,
        "d590_count": 0,
    }
    current_c_operation = ""
    current_d_operation = ""
    current_c100: dict[str, object] | None = None
    current_c500: dict[str, object] | None = None
    current_d500: dict[str, object] | None = None
    c100_c190_rows: list[list[object]] = []
    c500_c590_rows: list[list[object]] = []
    d500_d590_rows: list[list[object]] = []

    def finish_c100_document() -> None:
        nonlocal current_c100
        if not current_c100:
            return
        c100_value = Decimal(current_c100.get("document_value", Decimal("0.00")))
        c190_value = Decimal(current_c100.get("summary_value", Decimal("0.00")))
        difference = (c190_value - c100_value).quantize(Decimal("0.01"))
        if difference:
            c100_c190_rows.append(
                [
                    current_c100.get("operation_type", ""),
                    current_c100.get("model", ""),
                    current_c100.get("key", ""),
                    current_c100.get("number", ""),
                    current_c100.get("series", ""),
                    current_c100.get("issue_date", ""),
                    c100_value,
                    c190_value,
                    difference,
                    current_c100.get("line_number", ""),
                    "; ".join(str(item) for item in current_c100.get("summary_details", [])),
                    Decimal(current_c100.get("vl_merc", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_desc", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_frt", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_seg", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_out", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_st", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_ipi", Decimal("0.00"))),
                ]
            )
        current_c100 = None

    def finish_c500_document() -> None:
        nonlocal current_c500
        if not current_c500:
            return
        summary_value = Decimal(current_c500.get("summary_value", Decimal("0.00"))).quantize(Decimal("0.01"))
        if summary_value:
            c500_c590_rows.append(
                [
                    current_c500.get("operation_type", ""),
                    current_c500.get("model", ""),
                    current_c500.get("number", ""),
                    current_c500.get("series", ""),
                    current_c500.get("issue_date", ""),
                    summary_value,
                    current_c500.get("line_number", ""),
                    "; ".join(str(item) for item in current_c500.get("summary_details", [])),
                ]
            )
        current_c500 = None

    def finish_d500_document() -> None:
        nonlocal current_d500
        if not current_d500:
            return
        summary_value = Decimal(current_d500.get("summary_value", Decimal("0.00"))).quantize(Decimal("0.01"))
        if summary_value:
            d500_d590_rows.append(
                [
                    current_d500.get("operation_type", ""),
                    current_d500.get("model", ""),
                    current_d500.get("number", ""),
                    current_d500.get("series", ""),
                    current_d500.get("issue_date", ""),
                    summary_value,
                    current_d500.get("line_number", ""),
                    "; ".join(str(item) for item in current_d500.get("summary_details", [])),
                ]
            )
        current_d500 = None

    with sped_file.open("r", encoding="latin-1", errors="ignore") as current_file:
        for line_number, raw_line in enumerate(current_file, start=1):
            line = raw_line.strip()
            register = compare_register(line)
            fields = line.split("|")
            if register in {"C100", "C500"}:
                finish_c100_document()
                finish_c500_document()
                current_c_operation = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
            elif register.startswith("C") and register not in {"C190", "C590"} and register in {"C001", "C990"}:
                finish_c100_document()
                finish_c500_document()
                current_c_operation = ""
            if register == "D500":
                finish_d500_document()
                current_d_operation = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
            elif register.startswith("D") and register not in {"D590"} and register in {"D001", "D990"}:
                finish_d500_document()
                current_d_operation = ""

            if register == "C100":
                operation_type = current_c_operation
                if selected_scope == "Ambos" or operation_type == selected_scope:
                    document_value = compare_decimal_value(fields[12] if len(fields) > 12 else "")
                    totals["c100_total"] += document_value
                    counts["c100_count"] += 1
                    current_c100 = {
                        "operation_type": operation_type,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[8].strip() if len(fields) > 8 else "",
                        "key": fields[9].strip() if len(fields) > 9 else "",
                        "issue_date": fields[10].strip() if len(fields) > 10 else "",
                        "document_value": document_value,
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                        "vl_merc": compare_decimal_value(fields[16] if len(fields) > 16 else ""),
                        "vl_desc": compare_decimal_value(fields[14] if len(fields) > 14 else ""),
                        "vl_frt": compare_decimal_value(fields[18] if len(fields) > 18 else ""),
                        "vl_seg": compare_decimal_value(fields[19] if len(fields) > 19 else ""),
                        "vl_out": compare_decimal_value(fields[20] if len(fields) > 20 else ""),
                        "vl_st": compare_decimal_value(fields[24] if len(fields) > 24 else ""),
                        "vl_ipi": compare_decimal_value(fields[25] if len(fields) > 25 else ""),
                    }
                else:
                    current_c100 = None
            elif register == "A100":
                operation_type = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
                if selected_scope == "Ambos" or operation_type == selected_scope:
                    totals["a100_total"] += compare_decimal_value(fields[12] if len(fields) > 12 else "")
                    counts["a100_count"] += 1
            elif register == "C500":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    current_c500 = {
                        "operation_type": current_c_operation,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[10].strip() if len(fields) > 10 else "",
                        "issue_date": fields[11].strip() if len(fields) > 11 else "",
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                    }
                else:
                    current_c500 = None
            elif register == "D500":
                if selected_scope == "Ambos" or current_d_operation == selected_scope:
                    current_d500 = {
                        "operation_type": current_d_operation,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[8].strip() if len(fields) > 8 else "",
                        "issue_date": fields[10].strip() if len(fields) > 10 else "",
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                    }
                else:
                    current_d500 = None
            elif register == "C190":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["c190_total"] += summary_value
                    counts["c190_count"] += 1
                    if current_c100 is not None:
                        current_c100["summary_value"] = Decimal(current_c100.get("summary_value", Decimal("0.00"))) + summary_value
                        current_c100.setdefault("summary_details", []).append(
                            f"L{line_number} C190 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
            elif register == "C590":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["c590_total"] += summary_value
                    counts["c590_count"] += 1
                    if current_c500 is not None:
                        current_c500["summary_value"] = Decimal(current_c500.get("summary_value", Decimal("0.00"))) + summary_value
                        current_c500.setdefault("summary_details", []).append(
                            f"L{line_number} C590 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
            elif register == "D590":
                if selected_scope == "Ambos" or current_d_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["d590_total"] += summary_value
                    counts["d590_count"] += 1
                    if current_d500 is not None:
                        current_d500["summary_value"] = Decimal(current_d500.get("summary_value", Decimal("0.00"))) + summary_value
                        current_d500.setdefault("summary_details", []).append(
                            f"L{line_number} D590 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
    finish_c100_document()
    finish_c500_document()
    finish_d500_document()
    document_total = (totals["c100_total"] + totals["a100_total"]).quantize(Decimal("0.01"))
    fiscal_total = (totals["c190_total"] + totals["c590_total"] + totals["d590_total"]).quantize(Decimal("0.01"))
    result: dict[str, object] = {key: value.quantize(Decimal("0.01")) for key, value in totals.items()}
    result.update(counts)
    result["document_total"] = document_total
    result["fiscal_total"] = fiscal_total
    result["fiscal_document_difference"] = (fiscal_total - document_total).quantize(Decimal("0.01"))
    result["c100_c190_rows"] = c100_c190_rows
    result["c500_c590_rows"] = c500_c590_rows
    result["d500_d590_rows"] = d500_d590_rows
    return result


def collect_compare_sped_keys(
    sped_file: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> set[str]:
    return set(collect_compare_sped_documents(sped_file, operation_scope, progress_callback).keys())


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


def compare_sped_with_xml_folder(
    sped_file: Path,
    xml_folder: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareXmlInvoice], list[CompareXmlInvoice], list[CompareSpedDocument], dict[str, object]]:
    def report_range(start: int, end: int, current: int, total: int, message: str) -> None:
        if progress_callback:
            span = max(end - start, 1)
            percent = start + int(span * current / max(total, 1))
            progress_callback(percent, 100, message)

    sped_documents = collect_compare_sped_documents(
        sped_file,
        operation_scope,
        lambda current, total, message: report_range(0, 35, current, total, message),
    )
    sped_keys = set(sped_documents.keys())
    sped_documents_by_number_series, sped_documents_by_number = build_compare_sped_number_indexes(sped_documents)
    company_tax_id = extract_company_tax_id_from_sped(sped_file)
    selected_scope = normalize_compare_operation_scope(operation_scope)
    xml_invoices, ignored_xml = collect_compare_xml_invoices(
        xml_folder,
        lambda current, total, message: report_range(35, 85, current, total, message),
    )
    cancellation_events = collect_xml_cancellation_events(sorted(xml_folder.rglob("*.xml")))
    xml_operation_by_key = {
        invoice.key: infer_compare_invoice_operation(invoice, company_tax_id)
        for invoice in xml_invoices
    }
    filtered_xml_invoices = [
        invoice
        for invoice in xml_invoices
        if selected_scope == "Ambos" or xml_operation_by_key.get(invoice.key) == selected_scope
    ]
    cancelled_xml_invoices = [
        invoice
        for invoice in filtered_xml_invoices
        if invoice.key in cancellation_events
    ]
    active_xml_invoices = [
        invoice
        for invoice in filtered_xml_invoices
        if invoice.key not in cancellation_events
    ]
    if progress_callback:
        progress_callback(90, 100, "Comparando chaves entre SPED e XMLs.")
    xml_keys = {invoice.key for invoice in active_xml_invoices}
    xml_keys_with_cancelled = xml_keys | {invoice.key for invoice in cancelled_xml_invoices}
    xml_matches_by_key: dict[str, CompareSpedDocument] = {}
    xml_matched_by_number_keys: set[str] = set()
    matched_rows: list[CompareInvestigationMatchedRow] = []
    present: list[CompareXmlInvoice] = []
    missing: list[CompareXmlInvoice] = []
    for invoice in active_xml_invoices:
        document, match_type = find_compare_sped_document_by_key_or_number(
            invoice,
            sped_documents,
            sped_documents_by_number_series,
            sped_documents_by_number,
        )
        if document is None:
            missing.append(invoice)
            continue
        present.append(invoice)
        xml_matches_by_key[invoice.key] = document
        xml_value = compare_decimal_value(invoice.total_value)
        sped_value = compare_decimal_value(document.total_value)
        matched_rows.append(
            CompareInvestigationMatchedRow(
                key=invoice.key,
                number=invoice.number,
                series=invoice.series,
                model=getattr(invoice, "model", ""),
                issuer_cnpj=invoice.issuer_cnpj,
                issue_date=invoice.issue_date,
                xml_value=format_decimal_sped(xml_value),
                sped_value=format_decimal_sped(sped_value),
                difference=format_decimal_sped((sped_value - xml_value).quantize(Decimal("0.01"))),
                match_type=match_type,
                xml_file=invoice.file_path,
                sped_key=document.key,
                reason=describe_compare_xml_sped_value_difference(invoice, document),
            )
        )
        if match_type == "number":
            xml_matched_by_number_keys.add(document.key)

    matched_sped_keys = {document.key for document in xml_matches_by_key.values()}
    sped_missing_xml = [
        document
        for key, document in sped_documents.items()
        if key not in xml_keys_with_cancelled and key not in matched_sped_keys
    ]
    present.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    missing.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    cancelled_xml_invoices.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    sped_missing_xml.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    operation_by_key = {key: document.operation_type for key, document in sped_documents.items()}
    operation_by_key.update({key: document.operation_type for key, document in xml_matches_by_key.items()})
    fiscal_basis = collect_compare_fiscal_basis_totals(sped_file, operation_scope)
    present_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in present), Decimal("0.00")).quantize(Decimal("0.01"))
    present_sped_value_total = sum((compare_decimal_value(row.sped_value) for row in matched_rows), Decimal("0.00")).quantize(Decimal("0.01"))
    missing_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in missing), Decimal("0.00")).quantize(Decimal("0.01"))
    sped_missing_xml_value_total = sum((compare_decimal_value(document.total_value) for document in sped_missing_xml), Decimal("0.00")).quantize(Decimal("0.01"))
    cancelled_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in cancelled_xml_invoices), Decimal("0.00")).quantize(Decimal("0.01"))
    return present, missing, sped_missing_xml, {
        "sped_keys": len(sped_keys),
        "xml_total": len(active_xml_invoices),
        "xml_total_lidos": len(xml_invoices),
        "xml_ignored": ignored_xml,
        "xml_cancelled": len(cancelled_xml_invoices),
        "xml_cancelled_value_total": cancelled_xml_value_total,
        "cancelled_xml_invoices": cancelled_xml_invoices,
        "cancellation_events": list(cancellation_events.values()),
        "present_total": len(matched_sped_keys),
        "source_present_total": len(present),
        "missing_total": len(missing),
        "sped_missing_xml_total": len(sped_missing_xml),
        "match_by_number_total": len(xml_matched_by_number_keys),
        "source_match_by_number_total": sum(1 for invoice in present if xml_matches_by_key.get(invoice.key) and xml_matches_by_key[invoice.key].key in xml_matched_by_number_keys),
        "operation_by_key": operation_by_key,
        "source_operation_by_key": xml_operation_by_key,
        "matched_rows": matched_rows,
        "fiscal_basis": fiscal_basis,
        "present_xml_value_total": present_xml_value_total,
        "present_sped_value_total": present_sped_value_total,
        "missing_xml_value_total": missing_xml_value_total,
        "sped_missing_xml_value_total": sped_missing_xml_value_total,
    }


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


def compare_sped_with_sheet(
    sped_file: Path,
    sheet_file: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareSheetInvoice], list[CompareSheetInvoice], dict[str, object]]:
    def report_range(start: int, end: int, current: int, total: int, message: str) -> None:
        if progress_callback:
            span = max(end - start, 1)
            percent = start + int(span * current / max(total, 1))
            progress_callback(percent, 100, message)

    selected_scope = normalize_compare_operation_scope(operation_scope)
    all_sped_documents = collect_compare_sped_documents(sped_file, "Ambos")
    sped_documents = collect_compare_sped_documents(
        sped_file,
        operation_scope,
        lambda current, total, message: report_range(0, 45, current, total, message),
    )
    sped_keys = set(sped_documents.keys())
    sheet_invoices, ignored_rows = collect_compare_sheet_invoices(
        sheet_file,
        lambda current, total, message: report_range(45, 85, current, total, message),
    )
    sped_documents_by_number_series, sped_documents_by_number = build_compare_sped_number_indexes(sped_documents)
    sheet_matches_by_key: dict[str, CompareSpedDocument] = {}
    sheet_matched_by_number_keys: set[str] = set()
    sheet_source_matched_by_number = 0

    def find_sped_document_for_sheet_invoice(invoice: CompareSheetInvoice) -> CompareSpedDocument | None:
        nonlocal sheet_source_matched_by_number
        document, match_type = find_compare_sped_document_by_key_or_number(
            invoice,
            sped_documents,
            sped_documents_by_number_series,
            sped_documents_by_number,
        )
        if document is not None:
            sheet_matches_by_key[invoice.key] = document
            if match_type == "number":
                sheet_matched_by_number_keys.add(document.key)
                sheet_source_matched_by_number += 1
        return document

    filtered_sheet_invoices = [
        invoice
        for invoice in sheet_invoices
        if (
            selected_scope == "Ambos"
            or invoice.operation_type == selected_scope
            or (
                not invoice.operation_type
                and (
                    invoice.key not in all_sped_documents
                    or all_sped_documents[invoice.key].operation_type == selected_scope
                )
            )
        )
    ]
    if progress_callback:
        progress_callback(90, 100, "Comparando chaves entre SPED e planilha.")
    present = [row for row in filtered_sheet_invoices if find_sped_document_for_sheet_invoice(row) is not None]
    missing = [row for row in filtered_sheet_invoices if row.key not in {invoice.key for invoice in present}]
    present.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    missing.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    matched_sped_keys = {document.key for document in sheet_matches_by_key.values()}
    operation_by_key = {key: document.operation_type for key, document in all_sped_documents.items()}
    operation_by_key.update({key: document.operation_type for key, document in sheet_matches_by_key.items()})
    return present, missing, {
        "sped_keys": len(sped_keys),
        "sheet_total": len(filtered_sheet_invoices),
        "sheet_total_lidas": len(sheet_invoices),
        "sheet_ignored": ignored_rows,
        "present_total": len(matched_sped_keys),
        "source_present_total": len(present),
        "missing_total": len(missing),
        "match_by_number_total": len(sheet_matched_by_number_keys),
        "source_match_by_number_total": sheet_source_matched_by_number,
        "operation_by_key": operation_by_key,
        "source_operation_by_key": {invoice.key: invoice.operation_type for invoice in sheet_invoices if invoice.operation_type},
    }


def set_field(fields: list[str], index: int, value: str) -> None:
    while len(fields) <= index:
        fields.append("")
    fields[index] = value


def display_text(value: object, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


def map_xml_cfop_to_entry_cfops(cfop: str) -> set[str]:
    normalized_cfop = normalize_cfop(cfop)
    if normalized_cfop in {"1401", "1403"}:
        return {normalized_cfop}
    if normalized_cfop in {"5401", "5403", "5405"}:
        # Na entrada o SPED pode estar em 1401 ou 1403, mesmo quando o XML veio
        # com CFOP de saida 5401/5403/5405. O casamento final ainda exige chave + item.
        return {"1401", "1403"}
    return set()


def parse_nfe_xml_st_items(file_path: Path) -> tuple[str, dict[tuple[str, str], list[dict[str, object]]]]:
    namespace = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
    root = ET.parse(file_path).getroot()

    document_key = normalize_document_key(
        root.findtext(".//nfe:protNFe/nfe:infProt/nfe:chNFe", default="", namespaces=namespace)
    )
    if not document_key:
        inf_nfe = root.find(".//nfe:infNFe", namespace)
        if inf_nfe is not None:
            inf_nfe_id = inf_nfe.attrib.get("Id", "")
            if inf_nfe_id.startswith("NFe"):
                inf_nfe_id = inf_nfe_id[3:]
            document_key = normalize_document_key(inf_nfe_id)

    items: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    if not document_key:
        return "", items

    for det in root.findall(".//nfe:det", namespace):
        xml_cfop = det.findtext("nfe:prod/nfe:CFOP", default="", namespaces=namespace)
        target_cfops = map_xml_cfop_to_entry_cfops(xml_cfop)
        if not target_cfops:
            continue

        icms_group = det.find("nfe:imposto/nfe:ICMS", namespace)
        if icms_group is None:
            continue

        base_icms_st = Decimal("0")
        icms_st_rate = Decimal("0")
        icms_st_value = Decimal("0")
        for element in icms_group.iter():
            local_name = get_xml_local_name(element.tag)
            text = (element.text or "").strip()
            if not text:
                continue
            if local_name in {"vBCST", "vBCSTRet"}:
                base_icms_st = parse_decimal(text)
            elif local_name == "pICMSST":
                icms_st_rate = parse_decimal(text)
            elif local_name in {"vICMSST", "vICMSSTRet"}:
                icms_st_value = parse_decimal(text)

        item_number = (det.attrib.get("nItem", "") or "").strip()
        code = first_non_empty(
            det.findtext("nfe:prod/nfe:cEANTrib", default="", namespaces=namespace),
            det.findtext("nfe:prod/nfe:cProd", default="", namespaces=namespace),
        )
        xml_codes = {
            normalize_document_key(det.findtext("nfe:prod/nfe:cEANTrib", default="", namespaces=namespace)),
            normalize_document_key(det.findtext("nfe:prod/nfe:cEAN", default="", namespaces=namespace)),
            normalize_document_key(det.findtext("nfe:prod/nfe:cProd", default="", namespaces=namespace)),
        }
        xml_codes.discard("")
        if not item_number and not code:
            continue

        for target_cfop in target_cfops:
            items[(item_number, target_cfop)].append(
                {
                    "code": str(code).strip(),
                    "xml_codes": xml_codes,
                    "xml_cfop": normalize_cfop(xml_cfop),
                    "cfop": target_cfop,
                    "base_icms_st": base_icms_st,
                    "icms_st_rate": icms_st_rate,
                    "icms_st_value": icms_st_value,
                }
            )

    return document_key, items


def collect_xml_candidate_document_keys(detailed_sales: list[dict[str, object]]) -> set[str]:
    candidate_keys: set[str] = set()
    for item in detailed_sales:
        if str(item.get("operation_type", "")).strip() != "Entrada":
            continue
        original_cfop = str(item.get("original_cfop", item.get("cfop", ""))).strip()
        if original_cfop not in {"1401", "1403"}:
            continue
        document_key = normalize_document_key(str(item.get("document_key", "")))
        if document_key:
            candidate_keys.add(document_key)
    return candidate_keys


def read_xml_document_key(file_path: Path) -> str:
    try:
        for _, element in ET.iterparse(file_path, events=("end",)):
            local_name = get_xml_local_name(element.tag)
            if local_name == "chNFe":
                text = (element.text or "").strip()
                if text:
                    return normalize_document_key(text)
            if local_name == "infNFe":
                inf_nfe_id = element.attrib.get("Id", "")
                if inf_nfe_id.startswith("NFe"):
                    inf_nfe_id = inf_nfe_id[3:]
                normalized = normalize_document_key(inf_nfe_id)
                if normalized:
                    return normalized
        return ""
    except ET.ParseError:
        return ""


def parse_selected_paths(raw_value: str) -> list[Path]:
    selected_paths: list[Path] = []
    for chunk in str(raw_value or "").split(";"):
        text = chunk.strip().strip('"')
        if not text:
            continue
        path = Path(text)
        if path not in selected_paths:
            selected_paths.append(path)
    return selected_paths


def format_selected_paths(paths: list[Path]) -> str:
    return "; ".join(str(path) for path in paths)


XML_SELECTION_COLLAPSE_THRESHOLD = 250


def collapse_xml_selection_paths(selected_files: list[str]) -> tuple[list[Path], bool]:
    normalized_paths: list[Path] = []
    for selected in selected_files:
        selected_path = Path(selected)
        if selected_path not in normalized_paths:
            normalized_paths.append(selected_path)
    if len(normalized_paths) < XML_SELECTION_COLLAPSE_THRESHOLD:
        return normalized_paths, False

    parent_folders = {path.parent.resolve() for path in normalized_paths}
    if len(parent_folders) == 1:
        return [next(iter(parent_folders))], True
    return normalized_paths, False


def get_xml_worker_count(total_files: int) -> int:
    if total_files <= 1:
        return 1
    cpu_total = os.cpu_count() or 1
    return max(1, min(8, cpu_total, total_files))


def build_pis_cofins_period_comparison_rows(
    sped_paths: list[Path],
    operation_type: str,
    xml_sources: list[Path] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[str], list[dict[str, object]]]:
    if not sped_paths:
        raise ValueError("Selecione ao menos um arquivo SPED Contribuicoes para processar.")
    if len(sped_paths) > 12:
        raise ValueError("A consulta aceita no maximo 12 arquivos SPED por vez.")

    normalized_operation = normalize_operation_type(operation_type)
    normalized_xml_sources = [path for path in dict.fromkeys(xml_sources or []) if path.exists()]
    period_labels: list[str] = []
    comparison_rows: list[dict[str, object]] = []
    used_labels: set[str] = set()

    total_files = len(sped_paths)
    for file_index, sped_path in enumerate(sped_paths, start=1):
        if not sped_path.exists():
            raise ValueError(f"O arquivo SPED nao foi encontrado: {sped_path}")
        if progress_callback:
            progress_callback(file_index - 1, total_files, f"Lendo {normalized_operation.lower()}: {sped_path.name}")

        detailed_rows = [
            item
            for item in read_sped_contrib_detailed_rows(sped_path)
            if normalize_operation_type(str(item.get("operation_type", "")).strip()) == normalized_operation
        ]
        document_rows = scan_sped_contrib_c100_documents(sped_path, normalized_operation)
        if normalized_xml_sources:
            rebuilt_rows = rebuild_sped_contrib_items_from_xml(document_rows, detailed_rows, normalized_operation, normalized_xml_sources)
            if rebuilt_rows:
                detailed_rows.extend(rebuilt_rows)
        label_base = infer_sped_period_label(sped_path, detailed_rows)
        label = label_base
        suffix = 2
        while label in used_labels:
            label = f"{label_base} ({suffix})"
            suffix += 1
        used_labels.add(label)
        period_labels.append(label)

        grouped_items: dict[str, list[dict[str, object]]] = defaultdict(list)
        for item in detailed_rows:
            code = str(item.get("code", "")).strip() or str(item.get("description", "")).strip()
            if code:
                grouped_items[code].append(item)

        for code, items in grouped_items.items():
            cfops = sorted({str(item.get("cfop", "")).strip() for item in items if str(item.get("cfop", "")).strip()})
            csts_pis = sorted({normalize_tax_code(item.get("cst_pis", ""), 2) for item in items if str(item.get("cst_pis", "")).strip()})
            csts_cofins = sorted({normalize_tax_code(item.get("cst_cofins", ""), 2) for item in items if str(item.get("cst_cofins", "")).strip()})
            descriptions = sorted({str(item.get("description", "")).strip() for item in items if str(item.get("description", "")).strip()})
            ncms = sorted({str(item.get("ncm", "")).strip() for item in items if str(item.get("ncm", "")).strip()})
            cests = collect_cest_values(items)
            document_keys = {
                normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
                for item in items
                if normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
            }
            launch_details = [dict(item) for item in sorted(items, key=lambda current: (str(current.get("document_date", "")), str(current.get("document_number", "")), str(current.get("item_number", ""))))]
            summary = summarize_pis_cofins_analysis(launch_details, normalized_operation)
            pis_base = sum((Decimal(item.get("base_pis", Decimal("0"))) for item in items), Decimal("0"))
            cofins_base = sum((Decimal(item.get("base_cofins", Decimal("0"))) for item in items), Decimal("0"))
            comparison_rows.append(
                {
                    "period": label,
                    "file_name": sped_path.name,
                    "curve_abc": "C",
                    "code": code,
                    "description": " | ".join(descriptions),
                    "ncm": " | ".join(ncms),
                    "cest": " | ".join(cests),
                    "cfop": " | ".join(cfops),
                    "cst_pis": " | ".join(csts_pis),
                    "cst_cofins": " | ".join(csts_cofins),
                    "suppliers": " | ".join(sorted({str(item.get("participant_name", "")).strip() for item in items if str(item.get("participant_name", "")).strip()})),
                    "supplier_count": len({(str(item.get("participant_name", "")).strip(), str(item.get("participant_tax_id", "")).strip()) for item in items if str(item.get("participant_name", "")).strip() or str(item.get("participant_tax_id", "")).strip()}),
                    "quantity": sum((Decimal(item.get("quantity", Decimal("0"))) for item in items), Decimal("0")),
                    "sale_value": sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                    "base_pis": pis_base,
                    "base_cofins": cofins_base,
                    "pis_value": Decimal(summary["pis_value"]),
                    "cofins_value": Decimal(summary["cofins_value"]),
                    "aliquota_pis": next((Decimal(item.get("aliquota_pis", Decimal("0"))) for item in launch_details if Decimal(item.get("aliquota_pis", Decimal("0"))) > 0), Decimal("0")),
                    "aliquota_cofins": next((Decimal(item.get("aliquota_cofins", Decimal("0"))) for item in launch_details if Decimal(item.get("aliquota_cofins", Decimal("0"))) > 0), Decimal("0")),
                    "document_count": len(document_keys),
                    "launch_count": len(items),
                    "status": str(summary["status"]),
                    "launch_details": launch_details,
                }
            )

        if progress_callback:
            progress_callback(file_index, total_files, f"{normalized_operation}s processadas: {sped_path.name}")

    curve_labels = calculate_abc_curve_labels(
        {
            str(row.get("code", "")).strip(): Decimal(row.get("sale_value", Decimal("0")))
            for row in comparison_rows
            if str(row.get("code", "")).strip()
        }
    )
    for row in comparison_rows:
        row["curve_abc"] = curve_labels.get(str(row.get("code", "")).strip(), "C")

    comparison_rows.sort(
        key=lambda item: (
            str(item.get("code", "")),
            str(item.get("period", "")),
            str(item.get("cfop", "")),
            str(item.get("cst_pis", "")),
            str(item.get("cst_cofins", "")),
        )
    )
    return period_labels, comparison_rows


def load_nfce_xml_items_for_index(
    file_path: Path,
    candidate_document_keys: set[str] | None = None,
    ncm_filters: set[str] | None = None,
) -> list[dict[str, object]]:
    if candidate_document_keys:
        document_key_hint = read_xml_document_key(file_path)
        if not document_key_hint or document_key_hint not in candidate_document_keys:
            return []

    items = parse_nfce_xml_items(file_path)
    if not items:
        return []

    filtered_items: list[dict[str, object]] = []
    for item in items:
        document_key = normalize_document_key(str(item.get("chave_acesso", "")))
        if not document_key:
            continue
        if candidate_document_keys and document_key not in candidate_document_keys:
            continue
        if ncm_filters and not ncm_matches_filters(str(item.get("ncm", "")), ncm_filters):
            continue
        filtered_items.append(item)
    return filtered_items


def export_nfce_items_by_ncm(
    xml_sources: list[Path],
    ncm_filters: set[str],
    output_path: Path,
) -> tuple[int, int, int]:
    xml_files = collect_xml_files(xml_sources)
    if not xml_files:
        raise ValueError("Nenhum arquivo XML foi encontrado nos caminhos selecionados.")

    filtered_items: list[dict[str, object]] = []
    skipped_files: list[list[object]] = []

    worker_count = get_xml_worker_count(len(xml_files))
    if worker_count == 1:
        file_entries = []
        for file_path in xml_files:
            try:
                items = load_nfce_xml_items_for_index(file_path, ncm_filters=ncm_filters)
            except ET.ParseError as exc:
                skipped_files.append([str(file_path), f"XML invalido: {exc}"])
                continue
            except Exception as exc:
                skipped_files.append([str(file_path), str(exc)])
                continue
            file_entries.append((file_path, items))
    else:
        file_entries = []
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_to_path = {
                executor.submit(load_nfce_xml_items_for_index, file_path, None, ncm_filters): file_path
                for file_path in xml_files
            }
            for future in as_completed(future_to_path):
                file_path = future_to_path[future]
                try:
                    items = future.result()
                except ET.ParseError as exc:
                    skipped_files.append([str(file_path), f"XML invalido: {exc}"])
                    continue
                except Exception as exc:
                    skipped_files.append([str(file_path), str(exc)])
                    continue
                file_entries.append((file_path, items))

    for _, items in sorted(file_entries, key=lambda current: str(current[0])):
        filtered_items.extend(items)

    if not filtered_items:
        raise ValueError("Nenhum item de NFC-e foi encontrado para o(s) NCM(s) informado(s).")

    headers = [
        "Arquivo XML",
        "Chave de Acesso",
        "Numero NFC-e",
        "Serie",
        "Data Emissao",
        "CNPJ Emitente",
        "Emitente",
        "Item",
        "Codigo",
        "EAN",
        "Descricao",
        "NCM",
        "CFOP",
        "Unidade",
        "Quantidade",
        "Valor Unitario",
        "Valor Produto",
        "Valor Desconto",
        "Valor Frete",
        "Valor Outros",
        "Valor Total Tributos",
        "Orig ICMS",
        "CST ICMS",
        "CSOSN ICMS",
        "Base ICMS",
        "Aliquota ICMS",
        "Valor ICMS",
        "Base ICMS ST",
        "Aliquota ICMS ST",
        "Valor ICMS ST",
        "CST PIS",
        "Base PIS",
        "Aliquota PIS",
        "Valor PIS",
        "CST COFINS",
        "Base COFINS",
        "Aliquota COFINS",
        "Valor COFINS",
        "CST IPI",
        "Base IPI",
        "Aliquota IPI",
        "Valor IPI",
    ]
    item_rows = [
        [
            item["arquivo_xml"],
            item["chave_acesso"],
            item["numero_nfce"],
            item["serie"],
            item["data_emissao"],
            item["cnpj_emitente"],
            item["emitente"],
            item["item"],
            item["codigo"],
            item["ean"],
            item["descricao"],
            item["ncm"],
            item["cfop"],
            item["unidade"],
            item["quantidade"],
            item["valor_unitario"],
            item["valor_produto"],
            item["valor_desconto"],
            item["valor_frete"],
            item["valor_outros"],
            item["valor_total_tributos"],
            item["orig_icms"],
            item["cst_icms"],
            item["csosn_icms"],
            item["base_icms"],
            item["aliquota_icms"],
            item["valor_icms"],
            item["base_icms_st"],
            item["aliquota_icms_st"],
            item["valor_icms_st"],
            item["cst_pis"],
            item["base_pis"],
            item["aliquota_pis"],
            item["valor_pis"],
            item["cst_cofins"],
            item["base_cofins"],
            item["aliquota_cofins"],
            item["valor_cofins"],
            item["cst_ipi"],
            item["base_ipi"],
            item["aliquota_ipi"],
            item["valor_ipi"],
        ]
        for item in filtered_items
    ]

    resumo_por_ncm: dict[str, dict[str, object]] = {}
    for item in filtered_items:
        ncm = str(item["ncm"])
        bucket = resumo_por_ncm.setdefault(
            ncm,
            {
                "ncm": ncm,
                "quantidade_itens": 0,
                "valor_produto": Decimal("0"),
                "valor_icms": Decimal("0"),
                "valor_icms_st": Decimal("0"),
                "valor_pis": Decimal("0"),
                "valor_cofins": Decimal("0"),
                "valor_ipi": Decimal("0"),
                "valor_total_tributos": Decimal("0"),
            },
        )
        bucket["quantidade_itens"] = int(bucket["quantidade_itens"]) + 1
        for field_name in ("valor_produto", "valor_icms", "valor_icms_st", "valor_pis", "valor_cofins", "valor_ipi", "valor_total_tributos"):
            bucket[field_name] = bucket[field_name] + item[field_name]

    resumo_headers = [
        "NCM",
        "Quantidade Itens",
        "Valor Produto",
        "Valor ICMS",
        "Valor ICMS ST",
        "Valor PIS",
        "Valor COFINS",
        "Valor IPI",
        "Valor Total Tributos",
    ]
    resumo_rows = [
        [
            row["ncm"],
            row["quantidade_itens"],
            row["valor_produto"],
            row["valor_icms"],
            row["valor_icms_st"],
            row["valor_pis"],
            row["valor_cofins"],
            row["valor_ipi"],
            row["valor_total_tributos"],
        ]
        for row in sorted(resumo_por_ncm.values(), key=lambda current: str(current["ncm"]))
    ]

    sheets = [("Itens NFCe", headers, item_rows), ("Resumo NCM", resumo_headers, resumo_rows)]
    if skipped_files:
        sheets.append(("XMLs Ignorados", ["Arquivo XML", "Motivo"], skipped_files))

    write_simple_excel_workbook(output_path, sheets)
    return len(xml_files), len(filtered_items), len(skipped_files)


def compose_xml_icms_cst_for_sped(xml_item: dict[str, object], fallback: str = "") -> str:
    orig = "".join(char for char in str(xml_item.get("orig_icms", "")).strip() if char.isdigit())
    cst = "".join(char for char in str(xml_item.get("cst_icms", "")).strip() if char.isdigit())
    csosn = "".join(char for char in str(xml_item.get("csosn_icms", "")).strip() if char.isdigit())
    suffix = csosn or cst
    if suffix:
        return normalize_cst_icms_for_sped(f"{orig}{suffix}")
    return normalize_cst_icms_for_sped(fallback)


def build_xml_fiscal_item_index(
    xml_sources: list[Path],
    candidate_document_keys: set[str] | None = None,
    ncm_filters: set[str] | None = None,
) -> dict[str, list[dict[str, object]]]:
    xml_files = collect_xml_files(xml_sources)
    if not xml_files:
        return {}

    xml_index: dict[str, list[dict[str, object]]] = defaultdict(list)
    worker_count = get_xml_worker_count(len(xml_files))
    if worker_count == 1:
        indexed_entries = []
        for file_path in xml_files:
            try:
                items = load_nfce_xml_items_for_index(file_path, candidate_document_keys, ncm_filters)
            except (ET.ParseError, UnicodeDecodeError):
                continue
            indexed_entries.append((file_path, items))
    else:
        indexed_entries = []
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_to_path = {
                executor.submit(load_nfce_xml_items_for_index, file_path, candidate_document_keys, ncm_filters): file_path
                for file_path in xml_files
            }
            for future in as_completed(future_to_path):
                file_path = future_to_path[future]
                try:
                    items = future.result()
                except (ET.ParseError, UnicodeDecodeError):
                    continue
                indexed_entries.append((file_path, items))

    for _, items in sorted(indexed_entries, key=lambda current: str(current[0])):
        for item in items:
            document_key = normalize_document_key(str(item.get("chave_acesso", "")))
            if document_key:
                xml_index[document_key].append(item)
    return xml_index


def build_xml_fiscal_identity_index(xml_index: dict[str, list[dict[str, object]]]) -> dict[tuple[str, str, str], list[dict[str, object]]]:
    xml_index_by_identity: dict[tuple[str, str, str], list[dict[str, object]]] = {}
    for items in xml_index.values():
        if not items:
            continue
        first_item = items[0]
        identity_key = (
            str(first_item.get("modelo", "")).strip(),
            str(first_item.get("numero_nfce", "")).strip(),
            str(first_item.get("serie", "")).strip(),
        )
        if identity_key[1]:
            xml_index_by_identity[identity_key] = items
    return xml_index_by_identity


def scan_sped_contrib_c100_documents(
    sped_path: Path,
    operation_scope: str,
) -> dict[tuple[str, str], dict[str, object]]:
    documents: dict[tuple[str, str], dict[str, object]] = {}
    participants: dict[str, dict[str, str]] = {}
    scoped_operations = {
        str(item["operation_type"])
        for item in filter_details_by_operation_scope(
            [{"operation_type": "Entrada"}, {"operation_type": "Saida"}],
            operation_scope,
        )
    }

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|"):
            continue
        fields = normalize_sped_line(line)
        register = get_field(fields, 1)
        if register == "0150":
            participant_code = get_field(fields, 2)
            participants[participant_code] = {
                "name": get_field(fields, 3),
                "tax_id": first_non_empty(get_field(fields, 5), get_field(fields, 6)),
            }
            continue
        if register != "C100":
            continue
        operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
        if operation_type not in scoped_operations:
            continue
        document_number = get_field(fields, 8)
        document_key = normalize_document_key(get_field(fields, 9))
        if not document_key:
            continue
        participant_code = get_field(fields, 4)
        participant = participants.get(participant_code, {})
        documents[(operation_type, document_key)] = {
            "operation_type": operation_type,
            "document_number": document_number,
            "document_key": document_key,
            "document_date": first_non_empty(get_field(fields, 10), get_field(fields, 11)),
            "document_series": get_field(fields, 7),
            "participant_code": participant_code,
            "participant_name": participant.get("name", ""),
            "participant_tax_id": participant.get("tax_id", ""),
            "cod_mod": get_field(fields, 5),
            "cod_sit": get_field(fields, 6),
        }
    return documents


def rebuild_sped_contrib_items_from_xml(
    document_rows: dict[tuple[str, str], dict[str, object]],
    detailed_rows: list[dict[str, object]],
    operation_type: str,
    xml_sources: list[Path],
) -> list[dict[str, object]]:
    document_keys_with_items = {
        normalize_document_key(str(item.get("document_key", "")))
        for item in detailed_rows
        if normalize_document_key(str(item.get("document_key", "")))
    }
    candidate_document_keys = {
        document_key
        for (row_operation, document_key), document_meta in document_rows.items()
        if row_operation == operation_type and str(document_meta.get("cod_mod", "")).strip() in {"55", "65"} and document_key not in document_keys_with_items
    }
    if not candidate_document_keys:
        return []

    xml_index = build_xml_fiscal_item_index(xml_sources, candidate_document_keys=candidate_document_keys)
    if not xml_index:
        return []
    xml_index_by_identity = build_xml_fiscal_identity_index(xml_index)

    rebuilt_rows: list[dict[str, object]] = []
    for (row_operation, document_key), document_meta in document_rows.items():
        if row_operation != operation_type:
            continue
        if str(document_meta.get("cod_mod", "")).strip() not in {"55", "65"}:
            continue
        if document_key in document_keys_with_items:
            continue
        xml_items = xml_index.get(document_key, [])
        if not xml_items:
            xml_items = xml_index_by_identity.get(
                (
                    str(document_meta.get("cod_mod", "")).strip(),
                    str(document_meta.get("document_number", "")).strip(),
                    str(document_meta.get("document_series", "")).strip(),
                ),
                [],
            )
        if not xml_items:
            continue
        for xml_item in xml_items:
            rebuilt_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(document_meta.get("document_number", "")).strip() or str(xml_item.get("numero_nfce", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(document_meta.get("document_date", "")).strip() or str(xml_item.get("data_emissao", "")).strip(),
                    "document_series": str(document_meta.get("document_series", "")).strip() or str(xml_item.get("serie", "")).strip(),
                    "document_model": str(document_meta.get("cod_mod", "")).strip() or str(xml_item.get("modelo", "")).strip() or "55",
                    "participant_code": str(document_meta.get("participant_code", "")).strip(),
                    "participant_name": str(document_meta.get("participant_name", "")).strip() or str(xml_item.get("emitente", "")).strip(),
                    "participant_tax_id": str(document_meta.get("participant_tax_id", "")).strip() or str(xml_item.get("cnpj_emitente", "")).strip(),
                    "item_number": str(xml_item.get("item", "")).strip(),
                    "code": str(xml_item.get("codigo", "")).strip() or str(xml_item.get("descricao", "")).strip(),
                    "description": str(xml_item.get("descricao", "")).strip(),
                    "ncm": str(xml_item.get("ncm", "")).strip(),
                    "cest": str(xml_item.get("cest", "")).strip(),
                    "cfop": str(xml_item.get("cfop", "")).strip(),
                    "quantity": Decimal(xml_item.get("quantidade", Decimal("0"))),
                    "sale_value": Decimal(xml_item.get("valor_operacao", xml_item.get("valor_produto", Decimal("0")))),
                    "cst_pis": normalize_tax_code(xml_item.get("cst_pis", ""), 2),
                    "base_pis": Decimal(xml_item.get("base_pis", Decimal("0"))),
                    "aliquota_pis": Decimal(xml_item.get("aliquota_pis", Decimal("0"))),
                    "pis_value": Decimal(xml_item.get("valor_pis", Decimal("0"))),
                    "cst_cofins": normalize_tax_code(xml_item.get("cst_cofins", ""), 2),
                    "base_cofins": Decimal(xml_item.get("base_cofins", Decimal("0"))),
                    "aliquota_cofins": Decimal(xml_item.get("aliquota_cofins", Decimal("0"))),
                    "cofins_value": Decimal(xml_item.get("valor_cofins", Decimal("0"))),
                    "xml_rebuilt": True,
                }
            )
    return rebuilt_rows


def scan_sped_c100_documents(
    sped_path: Path,
    operation_scope: str,
) -> dict[tuple[str, str], dict[str, object]]:
    documents: dict[tuple[str, str], dict[str, object]] = {}
    scoped_operations = {
        str(item["operation_type"])
        for item in filter_details_by_operation_scope(
            [{"operation_type": "Entrada"}, {"operation_type": "Saida"}],
            operation_scope,
        )
    }

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|C100|"):
            continue
        fields = normalize_sped_line(line)
        operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
        if operation_type not in scoped_operations:
            continue
        document_number = get_field(fields, 8)
        document_key = normalize_document_key(get_field(fields, 9))
        if not document_key:
            continue
        documents[(operation_type, document_key)] = {
            "operation_type": operation_type,
            "document_number": document_number,
            "document_key": document_key,
            "document_date": get_field(fields, 10),
            "document_series": get_field(fields, 7),
            "participant_code": get_field(fields, 4),
            "cod_mod": get_field(fields, 5),
            "cod_sit": get_field(fields, 6),
        }
    return documents


def read_c100_c190_fallback_rows(sped_path: Path, operation_type: str, cod_mod: str = "65") -> list[dict[str, object]]:
    fallback_rows: list[dict[str, object]] = []
    current_document: dict[str, object] | None = None

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|"):
            continue
        fields = normalize_sped_line(line)
        register = get_field(fields, 1)

        if register == "C100":
            line_operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
            document_key = normalize_document_key(get_field(fields, 9))
            current_document = None
            if line_operation_type != operation_type:
                continue
            if str(get_field(fields, 5)).strip() != str(cod_mod).strip():
                continue
            if not document_key:
                continue
            current_document = {
                "operation_type": line_operation_type,
                "document_number": get_field(fields, 8),
                "document_key": document_key,
                "document_date": get_field(fields, 10),
                "document_series": get_field(fields, 7),
                "document_model": get_field(fields, 5),
                "participant_code": get_field(fields, 4),
                "participant_name": "",
                "participant_tax_id": "",
            }
            continue

        if register == "C190" and current_document:
            document_model = str(current_document.get("document_model", "")).strip()
            document_number = str(current_document.get("document_number", "")).strip()
            document_key = str(current_document.get("document_key", "")).strip()
            fallback_rows.append(
                {
                    **current_document,
                    "item_number": "",
                    "code": f"DOC-{document_number or document_key[:12]}",
                    "description": f"Documento modelo {document_model or cod_mod} sem itens C170 - fallback C190",
                    "ncm": "",
                    "cest": "",
                    "cst_icms": get_field(fields, 2),
                    "cfop": get_field(fields, 3),
                    "quantity": Decimal("0"),
                    "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
                    "sale_value": parse_decimal(get_field(fields, 5)),
                    "base_icms": parse_decimal(get_field(fields, 6)),
                    "icms_value": parse_decimal(get_field(fields, 7)),
                    "base_icms_st": parse_decimal(get_field(fields, 8)),
                    "icms_st_rate": Decimal("0"),
                    "icms_st_value": parse_decimal(get_field(fields, 9)),
                    "ipi_value": parse_decimal(get_field(fields, 11)),
                    "fallback_source": "C100/C190",
                }
            )
            continue

    return fallback_rows


def allocate_decimal_proportionally(
    weights: list[Decimal],
    target_total: Decimal,
) -> list[Decimal]:
    if not weights:
        return []
    quantized_target = target_total.quantize(Decimal("0.01"))
    positive_weights = [max(weight, Decimal("0")) for weight in weights]
    total_weight = sum(positive_weights, Decimal("0"))
    if total_weight <= 0:
        allocations = [Decimal("0.00")] * len(weights)
        allocations[-1] = quantized_target
        return allocations

    allocations: list[Decimal] = []
    allocated_total = Decimal("0.00")
    for index, weight in enumerate(positive_weights):
        if index == len(positive_weights) - 1:
            value = (quantized_target - allocated_total).quantize(Decimal("0.01"))
        else:
            value = (quantized_target * weight / total_weight).quantize(Decimal("0.01"))
            allocated_total += value
        allocations.append(value)
    return allocations


def normalize_xml_rebuilt_items_with_fallback(
    rebuilt_items: list[dict[str, object]],
    fallback_rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    if not rebuilt_items or not fallback_rows:
        return rebuilt_items

    fallback_groups: dict[tuple[str, str, Decimal], dict[str, Decimal]] = {}
    fallback_groups_by_cfop_cst: dict[tuple[str, str], list[tuple[tuple[str, str, Decimal], dict[str, Decimal]]]] = defaultdict(list)
    for fallback_row in fallback_rows:
        key = (
            normalize_cst_icms_for_sped(str(fallback_row.get("cst_icms", ""))),
            str(fallback_row.get("cfop", "")).strip(),
            Decimal(fallback_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01")),
        )
        bucket = fallback_groups.setdefault(
            key,
            {
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "ipi_value": Decimal("0"),
            },
        )
        bucket["base_icms"] += Decimal(fallback_row.get("base_icms", Decimal("0")))
        bucket["icms_value"] += Decimal(fallback_row.get("icms_value", Decimal("0")))
        bucket["base_icms_st"] += Decimal(fallback_row.get("base_icms_st", Decimal("0")))
        bucket["icms_st_value"] += Decimal(fallback_row.get("icms_st_value", Decimal("0")))
        bucket["ipi_value"] += Decimal(fallback_row.get("ipi_value", Decimal("0")))
    for key, totals in fallback_groups.items():
        fallback_groups_by_cfop_cst[(key[0], key[1])].append((key, totals))

    rebuilt_groups: dict[tuple[str, str, Decimal], list[dict[str, object]]] = defaultdict(list)
    for rebuilt_item in rebuilt_items:
        key = (
            normalize_cst_icms_for_sped(str(rebuilt_item.get("cst_icms", ""))),
            str(rebuilt_item.get("cfop", "")).strip(),
            Decimal(rebuilt_item.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01")),
        )
        rebuilt_groups[key].append(rebuilt_item)

    for key, grouped_items in rebuilt_groups.items():
        fallback_group_key = key
        fallback_totals = fallback_groups.get(fallback_group_key)
        if not fallback_totals:
            fallback_candidates = fallback_groups_by_cfop_cst.get((key[0], key[1]), [])
            if len(fallback_candidates) == 1:
                fallback_group_key, fallback_totals = fallback_candidates[0]
        if not fallback_totals:
            continue
        weights = [Decimal(item.get("sale_value", Decimal("0"))) for item in grouped_items]
        for field_name in ("base_icms", "icms_value", "base_icms_st", "icms_st_value", "ipi_value"):
            allocated_values = allocate_decimal_proportionally(weights, fallback_totals[field_name])
            for item, value in zip(grouped_items, allocated_values):
                item[field_name] = value
                item["icms_rate"] = fallback_group_key[2]
    return rebuilt_items


def find_matching_xml_fiscal_item(
    detail: dict[str, object],
    xml_items: list[dict[str, object]],
    used_indexes: set[int],
) -> tuple[int, dict[str, object]] | tuple[None, None]:
    item_number = str(detail.get("item_number", "")).strip()
    code = normalize_document_key(str(detail.get("code", "")))
    description = normalize_text(str(detail.get("description", "")))
    ncm = normalize_ncm(str(detail.get("ncm", "")))

    if item_number:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            if str(xml_item.get("item", "")).strip() == item_number:
                return index, xml_item

    if code:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            xml_codes = {
                normalize_document_key(str(xml_item.get("codigo", ""))),
                normalize_document_key(str(xml_item.get("ean", ""))),
            }
            xml_codes.discard("")
            if code in xml_codes:
                return index, xml_item

    if description and ncm:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            if (
                normalize_text(str(xml_item.get("descricao", ""))) == description
                and normalize_ncm(str(xml_item.get("ncm", ""))) == ncm
            ):
                return index, xml_item

    return None, None


def apply_xml_fiscal_adjustments_to_details(
    detailed_sales: list[dict[str, object]],
    sped_documents: dict[tuple[str, str], dict[str, object]],
    xml_sources: list[Path],
    ncm_filters: set[str],
) -> tuple[list[dict[str, object]], set[tuple[str, str]], list[dict[str, object]]]:
    if not xml_sources or not ncm_filters or not sped_documents:
        return list(detailed_sales), set(), []

    candidate_document_keys = {document_key for _, document_key in sped_documents}
    xml_index = build_xml_fiscal_item_index(xml_sources, candidate_document_keys, ncm_filters)
    if not xml_index:
        return list(detailed_sales), set(), []

    affected_documents: set[tuple[str, str]] = set()
    adjustment_rows: list[dict[str, object]] = []
    xml_rebuilt_details: list[dict[str, object]] = []

    for document_identity, document_meta in sped_documents.items():
        xml_items = xml_index.get(str(document_meta.get("document_key", "")), [])
        if not xml_items:
            continue
        affected_documents.add(document_identity)
        for xml_item in xml_items:
            rebuilt_item = {
                "operation_type": str(document_meta.get("operation_type", "")),
                "document_number": str(document_meta.get("document_number", "")),
                "document_key": str(document_meta.get("document_key", "")),
                "document_date": str(document_meta.get("document_date", "")),
                "item_number": str(xml_item.get("item", "")),
                "code": str(xml_item.get("codigo", "")),
                "description": str(xml_item.get("descricao", "")),
                "ncm": str(xml_item.get("ncm", "")),
                "cst_icms": compose_xml_icms_cst_for_sped(xml_item),
                "cfop": str(xml_item.get("cfop", "")),
                "quantity": xml_item.get("quantidade", Decimal("0")),
                "icms_rate": xml_item.get("aliquota_icms", Decimal("0")),
                "sale_value": xml_item.get("valor_produto", Decimal("0")),
                "base_icms": xml_item.get("base_icms", Decimal("0")),
                "icms_value": xml_item.get("valor_icms", Decimal("0")),
                "base_icms_st": xml_item.get("base_icms_st", Decimal("0")),
                "icms_st_rate": xml_item.get("aliquota_icms_st", Decimal("0")),
                "icms_st_value": xml_item.get("valor_icms_st", Decimal("0")),
                "ipi_value": xml_item.get("valor_ipi", Decimal("0")),
            }
            xml_rebuilt_details.append(rebuilt_item)
            adjustment_rows.append(
                {
                    "operation_type": rebuilt_item["operation_type"],
                    "document_number": rebuilt_item["document_number"],
                    "document_key": rebuilt_item["document_key"],
                    "item_number": rebuilt_item["item_number"],
                    "code": rebuilt_item["code"],
                    "ncm_original": "",
                    "ncm_xml": rebuilt_item["ncm"],
                    "cfop_original": "",
                    "cfop_xml": rebuilt_item["cfop"],
                    "cst_original": "",
                    "cst_xml": rebuilt_item["cst_icms"],
                    "vl_item_original": Decimal("0"),
                    "vl_item_xml": rebuilt_item["sale_value"],
                    "vl_bc_original": Decimal("0"),
                    "vl_bc_xml": rebuilt_item["base_icms"],
                    "aliq_original": Decimal("0"),
                    "aliq_xml": rebuilt_item["icms_rate"],
                    "vl_icms_original": Decimal("0"),
                    "vl_icms_xml": rebuilt_item["icms_value"],
                    "vl_bc_st_original": Decimal("0"),
                    "vl_bc_st_xml": rebuilt_item["base_icms_st"],
                    "vl_icms_st_original": Decimal("0"),
                    "vl_icms_st_xml": rebuilt_item["icms_st_value"],
                    "vl_ipi_original": Decimal("0"),
                    "vl_ipi_xml": rebuilt_item["ipi_value"],
                }
            )

    preserved_sales = [
        dict(item)
        for item in detailed_sales
        if (
            str(item.get("operation_type", "")),
            normalize_document_key(str(item.get("document_key", ""))),
        ) not in affected_documents
    ]
    preserved_sales.extend(xml_rebuilt_details)
    return preserved_sales, affected_documents, adjustment_rows


def build_xml_st_index(
    xml_sources: list[Path],
    candidate_document_keys: set[str] | None = None,
) -> dict[str, dict[tuple[str, str], list[dict[str, object]]]]:
    if not xml_sources:
        return {}

    xml_files: list[Path] = []
    for xml_source in xml_sources:
        if not xml_source.exists():
            raise ValueError(f"O caminho de XML informado nao existe: {xml_source}")
        if xml_source.is_file():
            xml_files.append(xml_source)
        else:
            xml_files.extend(sorted(xml_source.rglob("*.xml")))

    xml_index: dict[str, dict[tuple[str, str], list[dict[str, object]]]] = {}
    for file_path in xml_files:
        if file_path.suffix.lower() != ".xml":
            continue
        if candidate_document_keys:
            document_key_hint = read_xml_document_key(file_path)
            if not document_key_hint or document_key_hint not in candidate_document_keys:
                continue
        try:
            document_key, items = parse_nfe_xml_st_items(file_path)
        except (ET.ParseError, UnicodeDecodeError):
            continue
        if document_key and items:
            document_bucket = xml_index.setdefault(document_key, {})
            for item_key, item_values in items.items():
                document_bucket.setdefault(item_key, []).extend(item_values)
    return xml_index


def apply_xml_st_adjustments_to_details(
    detailed_sales: list[dict[str, object]],
    xml_sources: list[Path],
) -> tuple[list[dict[str, object]], set[tuple[str, str]], list[dict[str, object]], list[dict[str, object]]]:
    candidate_document_keys = collect_xml_candidate_document_keys(detailed_sales)
    xml_index = build_xml_st_index(xml_sources, candidate_document_keys)
    if not xml_index:
        missing_rows = [
            {
                "operation_type": str(item.get("operation_type", "")).strip(),
                "document_number": str(item.get("document_number", "")).strip(),
                "document_key": normalize_document_key(str(item.get("document_key", ""))),
                "document_date": str(item.get("document_date", "")).strip(),
                "item_number": str(item.get("item_number", "")).strip(),
                "code": str(item.get("code", "")).strip(),
                "description": str(item.get("description", "")).strip(),
                "cfop_original": str(item.get("original_cfop", item.get("cfop", ""))).strip(),
                "cfop_ajustado": str(item.get("cfop", "")).strip(),
                "missing_reason": "Nenhum XML selecionado com chave correspondente.",
            }
            for item in detailed_sales
            if str(item.get("operation_type", "")).strip() == "Entrada"
            and str(item.get("original_cfop", item.get("cfop", ""))).strip() in {"1401", "1403"}
        ]
        return list(detailed_sales), set(), [], missing_rows

    adjusted_sales: list[dict[str, object]] = []
    affected_documents: set[tuple[str, str]] = set()
    adjustment_rows: list[dict[str, object]] = []
    missing_rows: list[dict[str, object]] = []
    fallback_usage: dict[str, set[tuple[str, str, str]]] = defaultdict(set)

    for item in detailed_sales:
        rebuilt_item = dict(item)
        operation_type = str(rebuilt_item.get("operation_type", "")).strip()
        document_key = normalize_document_key(str(rebuilt_item.get("document_key", "")))
        cfop = str(rebuilt_item.get("cfop", "")).strip()
        original_cfop = str(rebuilt_item.get("original_cfop", cfop)).strip()
        item_number = str(rebuilt_item.get("item_number", "")).strip()
        target_cfop = original_cfop if original_cfop in {"1401", "1403"} else cfop
        if operation_type != "Entrada" or target_cfop not in {"1401", "1403"} or not document_key:
            adjusted_sales.append(rebuilt_item)
            continue

        document_items = xml_index.get(document_key)
        if not document_items:
            missing_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(rebuilt_item.get("document_number", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(rebuilt_item.get("document_date", "")).strip(),
                    "item_number": item_number,
                    "code": str(rebuilt_item.get("code", "")).strip(),
                    "description": str(rebuilt_item.get("description", "")).strip(),
                    "cfop_original": target_cfop,
                    "cfop_ajustado": cfop,
                    "missing_reason": "Chave do SPED nao encontrada nos XMLs selecionados.",
                }
            )
            adjusted_sales.append(rebuilt_item)
            continue

        code = normalize_document_key(str(rebuilt_item.get("code", "")).strip())
        xml_candidates = document_items.get((item_number, target_cfop), [])
        xml_item = None

        if xml_candidates:
            if code:
                xml_item = next(
                    (
                        candidate
                        for candidate in xml_candidates
                        if code in candidate.get("xml_codes", set())
                    ),
                    None,
                )
            if xml_item is None and len(xml_candidates) == 1:
                xml_item = xml_candidates[0]
            if xml_item is not None:
                fallback_usage[document_key].add((item_number, target_cfop, str(xml_item.get("code", "")).strip()))

        if xml_item is None and code:
            used_keys = fallback_usage[document_key]
            for candidate_key, candidate_items in document_items.items():
                if candidate_key[1] != target_cfop:
                    continue
                for candidate_item in candidate_items:
                    usage_key = (candidate_key[0], candidate_key[1], str(candidate_item.get("code", "")).strip())
                    if usage_key in used_keys:
                        continue
                    if code in candidate_item.get("xml_codes", set()):
                        xml_item = candidate_item
                        used_keys.add(usage_key)
                        break
                if xml_item is not None:
                    break

        if xml_item is None:
            missing_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(rebuilt_item.get("document_number", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(rebuilt_item.get("document_date", "")).strip(),
                    "item_number": item_number,
                    "code": str(rebuilt_item.get("code", "")).strip(),
                    "description": str(rebuilt_item.get("description", "")).strip(),
                    "cfop_original": target_cfop,
                    "cfop_ajustado": cfop,
                    "missing_reason": "Chave encontrada, mas o item/CFOP 1401/1403 nao foi localizado no XML.",
                }
            )
            adjusted_sales.append(rebuilt_item)
            continue

        original_base_icms_st = rebuilt_item.get("base_icms_st", Decimal("0"))
        original_icms_st_rate = rebuilt_item.get("icms_st_rate", Decimal("0"))
        original_icms_st_value = rebuilt_item.get("icms_st_value", Decimal("0"))
        rebuilt_item["base_icms_st"] = xml_item["base_icms_st"]
        rebuilt_item["icms_st_rate"] = xml_item["icms_st_rate"]
        rebuilt_item["icms_st_value"] = xml_item["icms_st_value"]
        adjustment_rows.append(
            {
                "operation_type": operation_type,
                "document_number": str(rebuilt_item.get("document_number", "")),
                "document_key": document_key,
                "item_number": item_number,
                "code": str(rebuilt_item.get("code", "")).strip(),
                "cfop_sped": target_cfop,
                "cfop_ajustado": cfop,
                "cfop_xml": str(xml_item.get("xml_cfop", "")),
                "base_icms_st_original": original_base_icms_st,
                "icms_st_rate_original": original_icms_st_rate,
                "icms_st_value_original": original_icms_st_value,
                "base_icms_st_ajustado": xml_item["base_icms_st"],
                "icms_st_rate_ajustado": xml_item["icms_st_rate"],
                "icms_st_value_ajustado": xml_item["icms_st_value"],
            }
        )
        adjusted_sales.append(rebuilt_item)
        affected_documents.add((operation_type, document_key or str(rebuilt_item.get("document_number", ""))))

    return adjusted_sales, affected_documents, adjustment_rows, missing_rows


def build_operation_summary_rows(
    detailed_sales: list[dict[str, object]],
    operation_type: str,
    rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> tuple[list[dict[str, object]], dict[str, Decimal]]:
    grouped_rows: dict[tuple[str, str, Decimal], dict[str, object]] = {}
    normalized_operation = normalize_operation_type(operation_type)

    for source_item in detailed_sales:
        if normalize_operation_type(str(source_item.get("operation_type", ""))) != normalized_operation:
            continue
        item = apply_sped_icms_consistency_rules(source_item, rule_profile, runtime_rules)
        cst_icms = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
        cfop = str(item.get("cfop", "")).strip()
        icms_rate = item.get("icms_rate", Decimal("0"))
        icms_rate = icms_rate if isinstance(icms_rate, Decimal) else Decimal("0")
        rate_key = icms_rate.quantize(Decimal("0.01"))
        key = (cst_icms, cfop, rate_key)

        sale_value = item.get("sale_value", Decimal("0"))
        base_icms = item.get("base_icms", Decimal("0"))
        icms_value = item.get("icms_value", Decimal("0"))
        base_icms_st = item.get("base_icms_st", Decimal("0"))
        icms_st_value = item.get("icms_st_value", Decimal("0"))
        ipi_value = item.get("ipi_value", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_icms = base_icms if isinstance(base_icms, Decimal) else Decimal("0")
        icms_value = icms_value if isinstance(icms_value, Decimal) else Decimal("0")
        base_icms_st = base_icms_st if isinstance(base_icms_st, Decimal) else Decimal("0")
        icms_st_value = icms_st_value if isinstance(icms_st_value, Decimal) else Decimal("0")
        ipi_value = ipi_value if isinstance(ipi_value, Decimal) else Decimal("0")
        operation_value = (sale_value + ipi_value + icms_st_value).quantize(Decimal("0.01"))
        reduction_value = sale_value - base_icms if normalize_cst_icms_for_sped(cst_icms)[-2:] in {"20", "70"} and sale_value > base_icms else Decimal("0")

        if key not in grouped_rows:
            grouped_rows[key] = {
                "operation_type": normalized_operation,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": rate_key,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "ipi_value": Decimal("0"),
                "reduction_value": Decimal("0"),
                "document_keys": set(),
                "launch_count": 0,
            }

        grouped_rows[key]["total_operation_value"] += operation_value
        grouped_rows[key]["base_icms"] += base_icms
        grouped_rows[key]["icms_value"] += icms_value
        grouped_rows[key]["base_icms_st"] += base_icms_st
        grouped_rows[key]["icms_st_value"] += icms_st_value
        grouped_rows[key]["ipi_value"] += ipi_value
        grouped_rows[key]["reduction_value"] += reduction_value
        grouped_rows[key]["launch_count"] += 1
        document_identity = normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
        if document_identity:
            grouped_rows[key]["document_keys"].add(document_identity)

    rows = sorted(
        grouped_rows.values(),
        key=lambda row: (
            str(row["cst_icms"]),
            str(row["cfop"]),
            Decimal(row["icms_rate"]),
        ),
    )
    totals = {
        "total_operation_value": sum((Decimal(row["total_operation_value"]) for row in rows), Decimal("0")),
        "base_icms": sum((Decimal(row["base_icms"]) for row in rows), Decimal("0")),
        "icms_value": sum((Decimal(row["icms_value"]) for row in rows), Decimal("0")),
        "base_icms_st": sum((Decimal(row["base_icms_st"]) for row in rows), Decimal("0")),
        "icms_st_value": sum((Decimal(row["icms_st_value"]) for row in rows), Decimal("0")),
        "ipi_value": sum((Decimal(row["ipi_value"]) for row in rows), Decimal("0")),
        "reduction_value": sum((Decimal(row["reduction_value"]) for row in rows), Decimal("0")),
    }
    return rows, totals


def filter_c190_rows(
    c190_rows: list[dict[str, object]],
    cst_values: set[str],
    cfop_values: set[str],
) -> list[dict[str, object]]:
    normalized_csts = {normalize_cst_icms_for_sped(str(value)) for value in cst_values if str(value).strip()}
    normalized_cfops = {str(value).strip() for value in cfop_values if str(value).strip()}
    filtered_rows: list[dict[str, object]] = []
    for row in c190_rows:
        cst_value = normalize_cst_icms_for_sped(str(row.get("cst_icms", "")))
        cfop_value = str(row.get("cfop", "")).strip()
        if normalized_csts and cst_value not in normalized_csts:
            continue
        if normalized_cfops and cfop_value not in normalized_cfops:
            continue
        filtered_rows.append(row)
    return filtered_rows


def build_operation_summary_rows_from_c190(
    c190_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[dict[str, object]], dict[str, Decimal]]:
    grouped_rows: dict[tuple[str, str, Decimal], dict[str, object]] = {}
    normalized_operation = normalize_operation_type(operation_type)

    for source_row in c190_rows:
        if normalize_operation_type(str(source_row.get("operation_type", ""))) != normalized_operation:
            continue

        cst_icms = normalize_cst_icms_for_sped(str(source_row.get("cst_icms", "")))
        cfop = str(source_row.get("cfop", "")).strip()
        icms_rate = source_row.get("icms_rate", Decimal("0"))
        icms_rate = icms_rate if isinstance(icms_rate, Decimal) else Decimal("0")
        rate_key = icms_rate.quantize(Decimal("0.01"))
        key = (cst_icms, cfop, rate_key)

        total_operation_value = source_row.get("total_operation_value", Decimal("0"))
        base_icms = source_row.get("base_icms", Decimal("0"))
        icms_value = source_row.get("icms_value", Decimal("0"))
        base_icms_st = source_row.get("base_icms_st", Decimal("0"))
        icms_st_value = source_row.get("icms_st_value", Decimal("0"))
        ipi_value = source_row.get("ipi_value", Decimal("0"))
        reduction_value = source_row.get("reduction_value", Decimal("0"))
        total_operation_value = total_operation_value if isinstance(total_operation_value, Decimal) else Decimal("0")
        base_icms = base_icms if isinstance(base_icms, Decimal) else Decimal("0")
        icms_value = icms_value if isinstance(icms_value, Decimal) else Decimal("0")
        base_icms_st = base_icms_st if isinstance(base_icms_st, Decimal) else Decimal("0")
        icms_st_value = icms_st_value if isinstance(icms_st_value, Decimal) else Decimal("0")
        ipi_value = ipi_value if isinstance(ipi_value, Decimal) else Decimal("0")
        reduction_value = reduction_value if isinstance(reduction_value, Decimal) else Decimal("0")

        if key not in grouped_rows:
            grouped_rows[key] = {
                "operation_type": normalized_operation,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": rate_key,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "ipi_value": Decimal("0"),
                "reduction_value": Decimal("0"),
                "document_keys": set(),
                "launch_count": 0,
            }

        grouped_rows[key]["total_operation_value"] += total_operation_value
        grouped_rows[key]["base_icms"] += base_icms
        grouped_rows[key]["icms_value"] += icms_value
        grouped_rows[key]["base_icms_st"] += base_icms_st
        grouped_rows[key]["icms_st_value"] += icms_st_value
        grouped_rows[key]["ipi_value"] += ipi_value
        grouped_rows[key]["reduction_value"] += reduction_value
        grouped_rows[key]["launch_count"] += 1
        document_identity = normalize_document_key(str(source_row.get("document_key", ""))) or str(source_row.get("document_number", "")).strip()
        if document_identity:
            grouped_rows[key]["document_keys"].add(document_identity)

    rows = sorted(
        grouped_rows.values(),
        key=lambda row: (
            str(row["cst_icms"]),
            str(row["cfop"]),
            Decimal(row["icms_rate"]),
        ),
    )
    totals = {
        "total_operation_value": sum((Decimal(row["total_operation_value"]) for row in rows), Decimal("0")),
        "base_icms": sum((Decimal(row["base_icms"]) for row in rows), Decimal("0")),
        "icms_value": sum((Decimal(row["icms_value"]) for row in rows), Decimal("0")),
        "base_icms_st": sum((Decimal(row["base_icms_st"]) for row in rows), Decimal("0")),
        "icms_st_value": sum((Decimal(row["icms_st_value"]) for row in rows), Decimal("0")),
        "ipi_value": sum((Decimal(row["ipi_value"]) for row in rows), Decimal("0")),
        "reduction_value": sum((Decimal(row["reduction_value"]) for row in rows), Decimal("0")),
    }
    return rows, totals


def build_synthetic_launch_details_from_c190(
    c190_rows: list[dict[str, object]],
    cst_icms: str,
    cfop: str,
    icms_rate: Decimal,
    operation_type: str,
) -> list[dict[str, object]]:
    normalized_operation = normalize_operation_type(operation_type)
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)
    normalized_cfop = str(cfop).strip()
    rate_key = icms_rate.quantize(Decimal("0.01"))
    synthetic_details: list[dict[str, object]] = []

    for row in c190_rows:
        if normalize_operation_type(str(row.get("operation_type", ""))) != normalized_operation:
            continue
        row_rate = row.get("icms_rate", Decimal("0"))
        row_rate = row_rate if isinstance(row_rate, Decimal) else Decimal("0")
        if normalize_cst_icms_for_sped(str(row.get("cst_icms", ""))) != normalized_cst:
            continue
        if str(row.get("cfop", "")).strip() != normalized_cfop:
            continue
        if row_rate.quantize(Decimal("0.01")) != rate_key:
            continue

        detail = {
            "operation_type": row.get("operation_type", operation_type),
            "document_number": row.get("document_number", ""),
            "document_key": row.get("document_key", ""),
            "document_date": row.get("document_date", ""),
            "document_series": row.get("document_series", ""),
            "document_model": row.get("document_model", ""),
            "participant_code": row.get("participant_code", ""),
            "participant_name": row.get("participant_name", ""),
            "participant_tax_id": row.get("participant_tax_id", ""),
            "item_number": "C190",
            "code": "C190",
            "description": "Lancamento sintetico do registro C190",
            "ncm": "",
            "cest": "",
            "cst_icms": normalized_cst,
            "cfop": normalized_cfop,
            "quantity": Decimal("0"),
            "icms_rate": rate_key,
            "sale_value": row.get("total_operation_value", Decimal("0")),
            "total_operation_value": row.get("total_operation_value", Decimal("0")),
            "discount_value": Decimal("0"),
            "base_icms": row.get("base_icms", Decimal("0")),
            "icms_value": row.get("icms_value", Decimal("0")),
            "base_icms_st": row.get("base_icms_st", Decimal("0")),
            "icms_st_value": row.get("icms_st_value", Decimal("0")),
            "ipi_value": row.get("ipi_value", Decimal("0")),
            "source_register": row.get("source_register", "C190"),
            "source_line_number": row.get("line_number", ""),
            "raw_line": row.get("raw_line", ""),
        }
        detail["note_snapshot"] = {
            "document_number": detail["document_number"],
            "document_key": detail["document_key"],
            "document_date": detail["document_date"],
            "document_series": detail["document_series"],
            "document_model": detail["document_model"],
            "participant_code": detail["participant_code"],
            "participant_name": detail["participant_name"],
            "participant_tax_id": detail["participant_tax_id"],
            "items": [dict(detail)],
        }
        synthetic_details.append(detail)

    return synthetic_details


def get_note_item_display_sale_value(item: dict[str, object]) -> Decimal:
    sale_value = Decimal(item.get("sale_value", Decimal("0")))
    description = str(item.get("description", "")).strip()
    code = str(item.get("code", "")).strip()
    description_key = normalize_text(description)
    is_c190_fallback = (
        str(item.get("fallback_source", "")).strip() == "C100/C190"
        or str(item.get("source_register", "")).strip() == "C190"
        or str(item.get("item_number", "")).strip() == "C190"
        or ("SEM ITENS C170" in description_key and "FALLBACK C190" in description_key)
        or (code.startswith("DOC-") and "FALLBACK C190" in description_key)
    )
    if not is_c190_fallback:
        return sale_value

    tax_additions = (
        Decimal(item.get("icms_st_value", Decimal("0")))
        + Decimal(item.get("ipi_value", Decimal("0")))
    )
    adjusted_value = (sale_value - tax_additions).quantize(Decimal("0.01"))
    return adjusted_value if adjusted_value >= 0 else sale_value


def get_launch_total_operation_value(item: dict[str, object]) -> Decimal:
    total_operation_value = item.get("total_operation_value")
    if isinstance(total_operation_value, Decimal):
        return total_operation_value
    sale_value = Decimal(item.get("sale_value", Decimal("0")))
    if str(item.get("source_register", "")).strip() == "C190" or str(item.get("item_number", "")).strip() == "C190":
        return sale_value
    return (
        sale_value
        + Decimal(item.get("ipi_value", Decimal("0")))
        + Decimal(item.get("icms_st_value", Decimal("0")))
    ).quantize(Decimal("0.01"))


def get_operation_base_difference(item: dict[str, object]) -> Decimal:
    return (
        get_launch_total_operation_value(item)
        - Decimal(item.get("base_icms", Decimal("0")))
    ).quantize(Decimal("0.01"))


def describe_operation_base_difference(
    operation_value: Decimal,
    base_icms: Decimal,
    ipi_value: Decimal = Decimal("0"),
    icms_st_value: Decimal = Decimal("0"),
    discount_value: Decimal = Decimal("0"),
) -> str:
    difference = (operation_value - base_icms).quantize(Decimal("0.01"))
    if difference == Decimal("0.00"):
        return ""
    absolute_difference = abs(difference)
    tax_additions = (ipi_value + icms_st_value).quantize(Decimal("0.01"))
    if ipi_value > 0 and absolute_difference == ipi_value.quantize(Decimal("0.01")):
        return "IPI"
    if icms_st_value > 0 and absolute_difference == icms_st_value.quantize(Decimal("0.01")):
        return "ICMS ST"
    if tax_additions > 0 and absolute_difference == tax_additions:
        return "IPI + ICMS ST"
    if discount_value > 0 and (operation_value - discount_value).quantize(Decimal("0.01")) == base_icms.quantize(Decimal("0.01")):
        return "Desconto"
    return "Operacao > Base" if difference > 0 else "Base > Operacao"


def build_operation_launch_details_map(
    detailed_sales: list[dict[str, object]],
    operation_type: str,
    rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> dict[tuple[str, str, Decimal], list[dict[str, object]]]:
    grouped_details: dict[tuple[str, str, Decimal], list[dict[str, object]]] = defaultdict(list)
    normalized_operation = normalize_operation_type(operation_type)
    documents_by_identity: dict[tuple[str, str], dict[str, object]] = {}

    for item in detailed_sales:
        if normalize_operation_type(str(item.get("operation_type", ""))) != normalized_operation:
            continue
        document_key = normalize_document_key(str(item.get("document_key", "")))
        document_number = str(item.get("document_number", "")).strip()
        identity = (document_key or document_number, str(item.get("document_date", "")).strip())
        if not identity[0]:
            continue
        bucket = documents_by_identity.setdefault(
            identity,
            {
                "document_number": document_number,
                "document_key": document_key,
                "document_date": str(item.get("document_date", "")).strip(),
                "document_series": str(item.get("document_series", "")).strip(),
                "document_model": str(item.get("document_model", "")).strip(),
                "participant_code": str(item.get("participant_code", "")).strip(),
                "participant_name": str(item.get("participant_name", "")).strip(),
                "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                "items": [],
            },
        )
        bucket["items"].append(dict(item))

    for source_item in detailed_sales:
        if normalize_operation_type(str(source_item.get("operation_type", ""))) != normalized_operation:
            continue
        item = apply_sped_icms_consistency_rules(source_item, rule_profile, runtime_rules)
        cst_icms = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
        cfop = str(item.get("cfop", "")).strip()
        icms_rate = item.get("icms_rate", Decimal("0"))
        icms_rate = icms_rate if isinstance(icms_rate, Decimal) else Decimal("0")
        enriched_item = dict(item)
        enriched_item["note_snapshot"] = documents_by_identity.get(
            (
                normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip(),
                str(item.get("document_date", "")).strip(),
            )
        )
        grouped_details[(cst_icms, cfop, icms_rate.quantize(Decimal("0.01")))].append(enriched_item)

    return grouped_details


def enrich_details_with_note_snapshots(
    detailed_sales: list[dict[str, object]],
    operation_type: str = "",
) -> list[dict[str, object]]:
    if not detailed_sales:
        return []

    normalized_operation = normalize_operation_type(operation_type) if str(operation_type).strip() else ""
    documents_by_identity: dict[tuple[str, str, str], dict[str, object]] = {}

    for item in detailed_sales:
        if normalized_operation and normalize_operation_type(str(item.get("operation_type", ""))) != normalized_operation:
            continue
        document_key = normalize_document_key(str(item.get("document_key", "")))
        document_number = str(item.get("document_number", "")).strip()
        document_date = str(item.get("document_date", "")).strip()
        period = str(item.get("period", "")).strip()
        identity = (document_key or document_number, document_date, period)
        if not identity[0]:
            continue
        bucket = documents_by_identity.setdefault(
            identity,
            {
                "document_number": document_number,
                "document_key": document_key,
                "document_date": document_date,
                "document_series": str(item.get("document_series", "")).strip(),
                "document_model": str(item.get("document_model", "")).strip(),
                "participant_code": str(item.get("participant_code", "")).strip(),
                "participant_name": str(item.get("participant_name", "")).strip(),
                "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                "period": period,
                "items": [],
            },
        )
        bucket["items"].append(dict(item))

    enriched_details: list[dict[str, object]] = []
    for item in detailed_sales:
        enriched_item = dict(item)
        if normalized_operation and normalize_operation_type(str(item.get("operation_type", ""))) != normalized_operation:
            enriched_details.append(enriched_item)
            continue
        document_key = normalize_document_key(str(item.get("document_key", "")))
        document_number = str(item.get("document_number", "")).strip()
        document_date = str(item.get("document_date", "")).strip()
        period = str(item.get("period", "")).strip()
        note_snapshot = documents_by_identity.get((document_key or document_number, document_date, period))
        if isinstance(note_snapshot, dict):
            enriched_item["note_snapshot"] = note_snapshot
        enriched_details.append(enriched_item)

    return enriched_details


def build_reduction_launch_rows(
    detailed_sales: list[dict[str, object]],
    rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> list[dict[str, object]]:
    reduction_rows: list[dict[str, object]] = []

    for item in detailed_sales:
        adjusted_item = apply_sped_icms_consistency_rules(item, rule_profile, runtime_rules)
        sale_value = adjusted_item.get("sale_value", Decimal("0"))
        base_icms = adjusted_item.get("base_icms", Decimal("0"))
        icms_value = adjusted_item.get("icms_value", Decimal("0"))
        icms_rate = adjusted_item.get("icms_rate", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_icms = base_icms if isinstance(base_icms, Decimal) else Decimal("0")
        icms_value = icms_value if isinstance(icms_value, Decimal) else Decimal("0")
        icms_rate = icms_rate if isinstance(icms_rate, Decimal) else Decimal("0")
        if not has_icms_reduction(icms_rate, sale_value, base_icms, icms_value):
            continue

        reduction_rows.append(
            {
                "operation_type": str(adjusted_item.get("operation_type", "")).strip(),
                "document_date": str(adjusted_item.get("document_date", "")).strip(),
                "document_number": str(adjusted_item.get("document_number", "")).strip(),
                "participant_name": str(adjusted_item.get("participant_name", "")).strip(),
                "document_key": str(adjusted_item.get("document_key", "")).strip(),
                "item_number": str(adjusted_item.get("item_number", "")).strip(),
                "code": str(adjusted_item.get("code", "")).strip(),
                "description": str(adjusted_item.get("description", "")).strip(),
                "cfop": str(adjusted_item.get("cfop", "")).strip(),
                "cst_icms": str(adjusted_item.get("cst_icms", "")).strip(),
                "icms_rate": icms_rate,
                "sale_value": sale_value,
                "base_icms": base_icms,
                "icms_value": icms_value,
                "reduction_value": max(Decimal("0"), sale_value - base_icms),
                "note_snapshot": adjusted_item.get("note_snapshot"),
            }
        )

    reduction_rows.sort(
        key=lambda row: (
            str(row["operation_type"]),
            str(row["document_date"]),
            str(row["document_number"]),
            str(row["item_number"]),
            str(row["code"]),
        )
    )
    return reduction_rows


def build_reduction_rows_from_c190(c190_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    reduction_rows: list[dict[str, object]] = []
    for row in c190_rows:
        reduction_value = row.get("reduction_value", Decimal("0"))
        reduction_value = reduction_value if isinstance(reduction_value, Decimal) else Decimal("0")
        if reduction_value <= Decimal("0"):
            continue
        total_operation_value = row.get("total_operation_value", Decimal("0"))
        base_icms = row.get("base_icms", Decimal("0"))
        icms_value = row.get("icms_value", Decimal("0"))
        icms_rate = row.get("icms_rate", Decimal("0"))
        reduction_rows.append(
            {
                "operation_type": str(row.get("operation_type", "")).strip(),
                "document_date": str(row.get("document_date", "")).strip(),
                "document_number": str(row.get("document_number", "")).strip(),
                "participant_name": "",
                "document_key": str(row.get("document_key", "")).strip(),
                "item_number": "",
                "code": "",
                "description": "Reducao de base identificada pelo resumo fiscal C190",
                "cfop": str(row.get("cfop", "")).strip(),
                "cst_icms": str(row.get("cst_icms", "")).strip(),
                "icms_rate": icms_rate if isinstance(icms_rate, Decimal) else Decimal("0"),
                "sale_value": total_operation_value if isinstance(total_operation_value, Decimal) else Decimal("0"),
                "base_icms": base_icms if isinstance(base_icms, Decimal) else Decimal("0"),
                "icms_value": icms_value if isinstance(icms_value, Decimal) else Decimal("0"),
                "reduction_value": reduction_value,
                "note_snapshot": None,
            }
        )
    reduction_rows.sort(
        key=lambda row: (
            str(row["operation_type"]),
            str(row["document_date"]),
            str(row["document_number"]),
            str(row["cfop"]),
            str(row["cst_icms"]),
        )
    )
    return reduction_rows


def build_filtered_apuracao_rows(
    entry_totals: dict[str, Decimal],
    exit_totals: dict[str, Decimal],
    original_summary: dict[str, Decimal],
    preserve_original_adjustments: bool,
) -> list[tuple[str, Decimal]]:
    ajustes_debito_doc = original_summary["ajustes_debito_doc"] if preserve_original_adjustments else Decimal("0")
    ajustes_debito = original_summary["ajustes_debito"] if preserve_original_adjustments else Decimal("0")
    estornos_credito = original_summary["estornos_credito"] if preserve_original_adjustments else Decimal("0")
    ajustes_credito_doc = original_summary["ajustes_credito_doc"] if preserve_original_adjustments else Decimal("0")
    ajustes_credito = original_summary["ajustes_credito"] if preserve_original_adjustments else Decimal("0")
    estornos_debito = original_summary["estornos_debito"] if preserve_original_adjustments else Decimal("0")
    saldo_credor_anterior = original_summary["saldo_credor_anterior"] if preserve_original_adjustments else Decimal("0")
    deducoes = original_summary["deducoes"] if preserve_original_adjustments else Decimal("0")
    extra_apuracao = original_summary["extra_apuracao"] if preserve_original_adjustments else Decimal("0")

    total_debitos = Decimal(exit_totals.get("icms_value", Decimal("0")))
    total_creditos = Decimal(entry_totals.get("icms_value", Decimal("0")))
    saldo_bruto = (
        total_debitos
        + ajustes_debito_doc
        + ajustes_debito
        + estornos_credito
        - total_creditos
        - ajustes_credito_doc
        - ajustes_credito
        - estornos_debito
        - saldo_credor_anterior
    )
    saldo_devedor = saldo_bruto if saldo_bruto > Decimal("0") else Decimal("0")
    icms_recolher = saldo_devedor - deducoes
    if icms_recolher < Decimal("0"):
        icms_recolher = Decimal("0")
    saldo_credor_transportar = abs(saldo_bruto) if saldo_bruto < Decimal("0") else Decimal("0")

    return [
        ("SAIDAS E PRESTACOES COM DEBITO DO IMPOSTO", total_debitos),
        ("VALOR TOTAL DOS AJUSTES A DEBITO (decorrentes do documento fiscal)", ajustes_debito_doc),
        ("VALOR TOTAL DOS AJUSTES A DEBITO DO IMPOSTO", ajustes_debito),
        ("VALOR TOTAL DOS ESTORNOS DE CREDITOS", estornos_credito),
        ("VALOR TOTAL DOS CREDITOS POR ENTRADAS E AQUISICOES COM CREDITO DO IMPOSTO", total_creditos),
        ("VALOR TOTAL DOS AJUSTES A CREDITO (decorrentes do documento fiscal)", ajustes_credito_doc),
        ("VALOR TOTAL DOS AJUSTES A CREDITO DO IMPOSTO", ajustes_credito),
        ("VALOR TOTAL DOS ESTORNOS DE DEBITOS", estornos_debito),
        ("VALOR TOTAL DO SALDO CREDOR DO PERIODO ANTERIOR", saldo_credor_anterior),
        ("VALOR DO SALDO DEVEDOR", saldo_devedor),
        ("VALOR TOTAL DAS DEDUCOES", deducoes),
        ("VALOR TOTAL DO ICMS A RECOLHER", icms_recolher),
        ("VALOR TOTAL DO SALDO CREDOR A TRANSPORTAR PARA O PERIODO SEGUINTE", saldo_credor_transportar),
        ("VALORES RECOLHIDOS OU A RECOLHER, EXTRA-APURACAO", extra_apuracao),
    ]


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


def build_sales_rows_from_details(
    detailed_sales: list[dict[str, object]],
) -> tuple[list[ProductRecord], list[dict[str, object]]]:
    products: dict[str, ProductRecord] = {}
    grouped_sales: dict[tuple[str, str, Decimal], dict[str, object]] = {}

    for item in detailed_sales:
        code = str(item["code"])
        description = str(item["description"])
        ncm = str(item["ncm"])
        cst_icms = str(item["cst_icms"])
        icms_rate = item["icms_rate"]
        if code and code not in products:
            products[code] = ProductRecord(
                code=code,
                description=description,
                ncm=ncm,
                cst_icms=cst_icms,
                icms_rate=icms_rate if isinstance(icms_rate, Decimal) else None,
                cest=str(item.get("cest", "")).strip(),
            )

        key = (code, cst_icms, icms_rate if isinstance(icms_rate, Decimal) else Decimal("0"))
        if key not in grouped_sales:
            grouped_sales[key] = {
                "code": code,
                "description": description,
                "ncm": ncm,
                "cest": str(item.get("cest", "")).strip(),
                "cst_icms": cst_icms,
                "icms_rate": icms_rate if isinstance(icms_rate, Decimal) else Decimal("0"),
                "quantity": Decimal("0"),
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            }

        grouped_sales[key]["quantity"] += item["quantity"]
        grouped_sales[key]["sale_value"] += item["sale_value"]
        grouped_sales[key]["base_icms"] += item["base_icms"]
        grouped_sales[key]["icms_value"] += item["icms_value"]

    return (
        sorted(products.values(), key=lambda item: item.code),
        sorted(grouped_sales.values(), key=lambda item: (str(item["code"]), str(item["cst_icms"]), item["icms_rate"])),
    )


def build_c190_rows_from_details(
    detailed_sales: list[dict[str, object]],
    rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    # Reagrupa os itens detalhados no mesmo criterio do C190:
    # operacao + documento + chave + CST + CFOP + aliquota.
    c190_rows_map: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    c190_product_map: dict[tuple[str, str, str, str, str, str, Decimal], dict[str, object]] = {}

    for item in detailed_sales:
        item = apply_sped_icms_consistency_rules(item, rule_profile, runtime_rules)
        operation_type = str(item["operation_type"])
        document_number = str(item["document_number"])
        document_key = str(item["document_key"])
        document_date = str(item["document_date"])
        cst_icms = str(item["cst_icms"])
        cfop = str(item["cfop"])
        icms_rate = item["icms_rate"] if isinstance(item["icms_rate"], Decimal) else Decimal("0")
        sale_value = item["sale_value"] if isinstance(item["sale_value"], Decimal) else Decimal("0")
        base_icms = item["base_icms"] if isinstance(item["base_icms"], Decimal) else Decimal("0")
        cst_suffix = normalize_cst_icms_for_sped(cst_icms)[-2:]
        reduction_value = sale_value - base_icms if cst_suffix in {"20", "70"} and sale_value > base_icms else Decimal("0")

        c190_key = (operation_type, document_number, document_key, cst_icms, cfop, icms_rate)
        if c190_key not in c190_rows_map:
            c190_rows_map[c190_key] = {
                "operation_type": operation_type,
                "document_number": document_number,
                "document_key": document_key,
                "document_date": document_date,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": icms_rate,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "reduction_value": Decimal("0"),
                "ipi_value": Decimal("0"),
            }
        c190_rows_map[c190_key]["total_operation_value"] += sale_value
        c190_rows_map[c190_key]["base_icms"] += base_icms
        c190_rows_map[c190_key]["icms_value"] += item["icms_value"]
        c190_rows_map[c190_key]["base_icms_st"] += (
            item["base_icms_st"] if isinstance(item.get("base_icms_st"), Decimal) else Decimal("0")
        )
        c190_rows_map[c190_key]["icms_st_value"] += (
            item["icms_st_value"] if isinstance(item.get("icms_st_value"), Decimal) else Decimal("0")
        )
        c190_rows_map[c190_key]["reduction_value"] += reduction_value
        c190_rows_map[c190_key]["ipi_value"] += (
            item["ipi_value"] if isinstance(item.get("ipi_value"), Decimal) else Decimal("0")
        )

        product_key = (
            operation_type,
            document_number,
            document_key,
            str(item["code"]),
            cst_icms,
            cfop,
            icms_rate,
        )
        if product_key not in c190_product_map:
            c190_product_map[product_key] = {
                "operation_type": operation_type,
                "document_number": document_number,
                "document_key": document_key,
                "document_date": document_date,
                "code": str(item["code"]),
                "description": str(item["description"]),
                "ncm": str(item["ncm"]),
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": icms_rate,
                "quantity": Decimal("0"),
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "ipi_value": Decimal("0"),
            }
        c190_product_map[product_key]["quantity"] += item["quantity"]
        c190_product_map[product_key]["sale_value"] += sale_value
        c190_product_map[product_key]["base_icms"] += base_icms
        c190_product_map[product_key]["icms_value"] += item["icms_value"]
        c190_product_map[product_key]["ipi_value"] += (
            item["ipi_value"] if isinstance(item.get("ipi_value"), Decimal) else Decimal("0")
        )

    c190_rows = sorted(
        c190_rows_map.values(),
        key=lambda row: (
            str(row["operation_type"]),
            str(row["document_number"]),
            str(row["document_key"]),
            str(row["cst_icms"]),
            str(row["cfop"]),
            row["icms_rate"],
        ),
    )
    c190_product_rows = []
    for item in sorted(
        c190_product_map.values(),
        key=lambda row: (
            str(row["operation_type"]),
            str(row["document_number"]),
            str(row["document_key"]),
            str(row["code"]),
            str(row["cst_icms"]),
            str(row["cfop"]),
            row["icms_rate"],
        ),
    ):
        c190_match = c190_rows_map.get(
            (
                str(item["operation_type"]),
                str(item["document_number"]),
                str(item["document_key"]),
                str(item["cst_icms"]),
                str(item["cfop"]),
                item["icms_rate"],
            )
        )
        c190_product_rows.append(
            {
                **item,
                "c190_total_operation_value": c190_match["total_operation_value"] if c190_match else Decimal("0"),
                "c190_base_icms": c190_match["base_icms"] if c190_match else Decimal("0"),
                "c190_icms_value": c190_match["icms_value"] if c190_match else Decimal("0"),
                "c190_base_icms_st": c190_match["base_icms_st"] if c190_match else Decimal("0"),
                "c190_icms_st_value": c190_match["icms_st_value"] if c190_match else Decimal("0"),
                "c190_reduction_value": c190_match["reduction_value"] if c190_match else Decimal("0"),
                "c190_ipi_value": c190_match["ipi_value"] if c190_match else Decimal("0"),
            }
        )

    return c190_rows, c190_product_rows


def attach_c190_totals_to_products(
    c190_rows: list[dict[str, object]],
    c190_product_rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    c190_lookup: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    for item in c190_rows:
        c190_lookup[
            (
                str(item["operation_type"]),
                str(item["document_number"]),
                str(item["document_key"]),
                str(item["cst_icms"]),
                str(item["cfop"]),
                item["icms_rate"] if isinstance(item["icms_rate"], Decimal) else Decimal("0"),
            )
        ] = item

    result: list[dict[str, object]] = []
    for item in c190_product_rows:
        c190_match = c190_lookup.get(
            (
                str(item["operation_type"]),
                str(item["document_number"]),
                str(item["document_key"]),
                str(item["cst_icms"]),
                str(item["cfop"]),
                item["icms_rate"] if isinstance(item["icms_rate"], Decimal) else Decimal("0"),
            )
        )
        result.append(
            {
                **item,
                "c190_total_operation_value": c190_match["total_operation_value"] if c190_match else Decimal("0"),
                "c190_base_icms": c190_match["base_icms"] if c190_match else Decimal("0"),
                "c190_icms_value": c190_match["icms_value"] if c190_match else Decimal("0"),
                "c190_base_icms_st": c190_match["base_icms_st"] if c190_match else Decimal("0"),
                "c190_icms_st_value": c190_match["icms_st_value"] if c190_match else Decimal("0"),
                "c190_reduction_value": c190_match["reduction_value"] if c190_match else Decimal("0"),
                "c190_ipi_value": c190_match["ipi_value"] if c190_match else Decimal("0"),
            }
        )
    return result


def combine_imported_data(
    sped_data: tuple[list[ProductRecord], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]] | None,
    excel_data: tuple[list[ProductRecord], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]] | None,
) -> tuple[list[ProductRecord], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    # Quando SPED e planilha sao informados juntos, o SPED continua sendo a base
    # fiscal e a planilha entra para enriquecer/filtrar os produtos por chave NF-e.
    if sped_data and excel_data:
        sped_product_rows, _, sped_detailed_sales, sped_c190_rows, _ = sped_data
        excel_product_rows, _, excel_detailed_sales, _, excel_c190_product_rows = excel_data

        combined_products: dict[str, ProductRecord] = {item.code: item for item in sped_product_rows}
        for item in excel_product_rows:
            if item.code not in combined_products:
                combined_products[item.code] = item
                continue
            current = combined_products[item.code]
            combined_products[item.code] = ProductRecord(
                code=current.code,
                description=current.description or item.description,
                ncm=current.ncm or item.ncm,
                cst_icms=current.cst_icms or item.cst_icms,
                icms_rate=current.icms_rate if current.icms_rate is not None else item.icms_rate,
                cest=current.cest or item.cest,
            )

        all_detailed_sales = list(sped_detailed_sales)
        all_detailed_sales.extend(excel_detailed_sales)
        _, combined_sales_rows = build_sales_rows_from_details(all_detailed_sales)

        matched_c190_product_rows = attach_c190_totals_to_products(sped_c190_rows, excel_c190_product_rows)
        return (
            sorted(combined_products.values(), key=lambda item: item.code),
            combined_sales_rows,
            all_detailed_sales,
            sped_c190_rows,
            matched_c190_product_rows,
        )

    if sped_data:
        return sped_data
    if excel_data:
        return excel_data

    return [], [], [], [], []


def read_detailed_product_excel(
    file_path: Path,
) -> tuple[list[ProductRecord], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    # Importa a aba de detalhamento por produto e converte o layout da planilha
    # para a mesma estrutura interna usada pela leitura do SPED.
    rows: list[dict[str, str]] | None = None
    selected_sheet_name = ""
    for sheet_name in ("Relatorio Detalhado por Produto",):
        try:
            rows = read_xlsx_sheet(file_path, sheet_name)
            selected_sheet_name = sheet_name
            break
        except ValueError:
            continue

    if rows is None:
        selected_sheet_name = get_first_xlsx_sheet_name(file_path)
        rows = read_xlsx_sheet(file_path, selected_sheet_name)

    if not rows:
        raise ValueError(f'A aba "{selected_sheet_name}" esta vazia.')

    detailed_sales: list[dict[str, object]] = []
    for row in rows:
        description = resolve_header(row, "DESCRICAO", "DESCRICAOPRODUTO", "PRODUTO")
        code = first_non_empty(
            resolve_header(
                row,
                "CODPROD",
                "CODPRODMAPEADO",
                "CODIGOPRODUTO",
                "CODIGO",
                "CODPRODUTO",
                "PRODUTOCODIGO",
            ),
            description,
        )
        ncm = resolve_header(row, "NCM")
        cst_icms = resolve_header(row, "ICMSCST", "CST", "CSON", "CSTICMS", "CSOSN")
        cfop = resolve_header(row, "CFOP")
        operation_type = normalize_operation_type(
            first_non_empty(resolve_header(row, "TIPODEOPERACAO", "TIPOOPERACAO", "OPERACAO", "TPNF"), "Saida")
        )
        document_number = resolve_header(row, "NNF", "NUMERODOCUMENTO", "DOCUMENTO", "NUMDOC", "NUMERO")
        document_key = normalize_document_key(resolve_header(row, "CHAVEDEACESSONOTA", "CHAVEDOCUMENTO", "CHAVE", "CHAVENFE"))
        document_date = resolve_header(row, "DATA", "DATADOCUMENTO", "EMISSAO", "DTEMISSAO")
        quantity = parse_decimal(resolve_header(row, "QUANTIDADEPRODUTO", "QUANTIDADETOTAL", "QUANTIDADE"))
        sale_value = parse_decimal(resolve_header(row, "VALORDOPRODUTO", "VALORPRODUTO", "VALORVENDA", "VALORTOTAL", "VALORCONTABIL", "VALORPRODUTO"))
        base_icms = parse_decimal(
            resolve_header(
                row,
                "ICMSBASECALCULO",
                "BASEDEICMS",
                "BASEICMSPRODUTO",
                "BASECALCULOICMSPRODUTO",
                "BASECALCULOICMS",
                "BASEICMS",
                "BASEDOICMS",
                "BCICMS",
                "VLBASEICMS",
                "VALORBASEICMS",
                "VALORBASEDOICMS",
                "VLBCICMS",
                "VBC",
            )
        )
        icms_rate = parse_rate(
            resolve_header(
                row,
                "ICMSPERCENTUAL",
                "ALIQUOTADEICMS",
                "ALIQUOTAICMS",
                "ALIQUOTADEICMSSIMPLESNACIONAL",
                "ALIQUOTA",
                "ALIQICMS",
                "PERCENTUALICMS",
                "PERCENTUALDOICMS",
                "PICMS",
            )
        ) or Decimal("0")
        icms_value = parse_decimal(
            resolve_header(
                row,
                "VALORICMS",
                "VALORDOICMS",
                "VALORDOICMSSIMPLESNACIONAL",
                "VALORICMSPRODUTO",
                "VALORDOICMSPROPRIO",
                "VALORICMSOPERACAO",
                "VLDOICMS",
                "VLRICMS",
                "VICMS",
                "ICMS",
            )
        )
        if base_icms == Decimal("0") and sale_value > 0 and icms_rate > 0:
            base_icms = sale_value
        if icms_rate == Decimal("0") and base_icms > 0 and icms_value > 0:
            icms_rate = (icms_value * Decimal("100") / base_icms).quantize(Decimal("0.01"))
        if icms_value == Decimal("0") and base_icms > 0 and icms_rate > 0:
            icms_value = (base_icms * icms_rate / Decimal("100")).quantize(Decimal("0.01"))

        if not any([code, description, ncm, cst_icms, cfop, str(quantity), str(sale_value), str(base_icms), str(icms_value)]):
            continue

        detailed_sales.append(
            {
                "operation_type": operation_type,
                "document_number": document_number,
                "document_key": document_key,
                "document_date": document_date,
                "item_number": "",
                "code": code,
                "description": description,
                "ncm": ncm,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "quantity": quantity,
                "icms_rate": icms_rate,
                "sale_value": sale_value,
                "base_icms": base_icms,
                "icms_value": icms_value,
            }
        )

    if not detailed_sales:
        raise ValueError(f'Nenhuma linha valida foi encontrada na aba "{selected_sheet_name}".')

    product_rows, sales_rows = build_sales_rows_from_details(detailed_sales)
    c190_rows, c190_product_rows = build_c190_rows_from_details(detailed_sales)
    return product_rows, sales_rows, detailed_sales, c190_rows, c190_product_rows
def read_sped_file(
    file_path: Path,
) -> tuple[list[ProductRecord], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    # Extrai do SPED apenas os registros necessarios para montar cadastro,
    # resumo de vendas, detalhamento por item e relacionamento com o C190.
    products: dict[str, ProductRecord] = {}
    participants: dict[str, dict[str, str]] = {}
    product_rates: dict[str, set[Decimal]] = defaultdict(set)
    product_csts: dict[str, set[str]] = defaultdict(set)
    grouped_sales: dict[tuple[str, str, Decimal], dict[str, object]] = {}
    c190_product_view: dict[tuple[str, str, str, str, str, str, Decimal], dict[str, object]] = {}
    detailed_sales: list[dict[str, object]] = []
    c190_rows: list[dict[str, object]] = []
    current_operation = ""
    current_document = ""
    current_document_key = ""
    current_document_date = ""
    current_document_series = ""
    current_document_model = ""
    current_document_tax_id = ""
    current_participant_code = ""
    current_participant_name = ""
    current_participant_tax_id = ""

    with file_path.open("r", encoding="latin-1") as sped_file:
        for raw_line in sped_file:
            if not raw_line.startswith("|"):
                continue

            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)

            if register == "0200":
                code = get_field(fields, 2)
                description = get_field(fields, 3)
                ncm = get_field(fields, 8)
                cest = get_field(fields, 13)
                icms_rate = parse_rate(get_field(fields, 12))
                products[code] = ProductRecord(
                    code=code,
                    description=description,
                    ncm=ncm,
                    cst_icms="",
                    icms_rate=icms_rate,
                    cest=cest,
                )
                merge_product_rate(product_rates, code, icms_rate)
                continue

            if register == "0150":
                participant_code = get_field(fields, 2)
                participants[participant_code] = {
                    "name": get_field(fields, 3),
                    "tax_id": first_non_empty(get_field(fields, 5), get_field(fields, 4)),
                }
                continue

            if register == "C100":
                ind_oper = get_field(fields, 2)
                current_operation = "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else ""
                current_participant_code = get_field(fields, 4)
                current_document_model = get_field(fields, 5)
                current_document_series = get_field(fields, 7)
                current_document = get_field(fields, 8)
                current_document_key = get_field(fields, 9)
                current_document_date = get_field(fields, 10)
                current_document_tax_id = extract_tax_id_from_document_key(current_document_key)
                participant_data = participants.get(current_participant_code, {})
                current_participant_name = participant_data.get("name", "")
                current_participant_tax_id = participant_data.get("tax_id", "")
                continue

            if register == "C190":
                c190_rows.append(
                    {
                        "operation_type": current_operation,
                        "document_number": current_document,
                        "document_key": current_document_key,
                        "document_date": current_document_date,
                        "cst_icms": get_field(fields, 2),
                        "cfop": get_field(fields, 3),
                        "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
                        "total_operation_value": parse_decimal(get_field(fields, 5)),
                        "base_icms": parse_decimal(get_field(fields, 6)),
                        "icms_value": parse_decimal(get_field(fields, 7)),
                        "base_icms_st": parse_decimal(get_field(fields, 8)),
                        "icms_st_value": parse_decimal(get_field(fields, 9)),
                        "reduction_value": parse_decimal(get_field(fields, 10)),
                        "ipi_value": parse_decimal(get_field(fields, 11)),
                    }
                )
                continue

            if register != "C170":
                continue

            code = get_field(fields, 3)
            cst_icms = get_field(fields, 10)
            cfop = get_field(fields, 11)
            quantity = parse_decimal(get_field(fields, 5))
            sale_value = parse_decimal(get_field(fields, 7))
            discount_value = parse_decimal(get_field(fields, 8))
            base_icms = parse_decimal(get_field(fields, 13))
            icms_rate = parse_rate(get_field(fields, 14)) or Decimal("0")
            icms_value = parse_decimal(get_field(fields, 15))
            base_icms_st = parse_decimal(get_field(fields, 16))
            icms_st_rate = parse_rate(get_field(fields, 17)) or Decimal("0")
            icms_st_value = parse_decimal(get_field(fields, 18))
            base_ipi = parse_decimal(get_field(fields, 22))
            ipi_rate = parse_rate(get_field(fields, 23)) or Decimal("0")
            ipi_value = parse_decimal(get_field(fields, 24))

            merge_product_rate(product_rates, code, icms_rate)
            merge_product_cst(product_csts, code, cst_icms)
            product = products.get(code)
            description = product.description if product else ""
            ncm = product.ncm if product else ""
            cest = product.cest if product else ""

            key = (code, cst_icms, icms_rate)
            if key not in grouped_sales:
                grouped_sales[key] = {
                    "code": code,
                    "description": description,
                    "ncm": ncm,
                    "cest": cest,
                    "cst_icms": cst_icms,
                    "icms_rate": icms_rate,
                    "quantity": Decimal("0"),
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "base_icms_st": Decimal("0"),
                    "icms_st_value": Decimal("0"),
                    "base_ipi": Decimal("0"),
                    "ipi_rate": Decimal("0"),
                    "ipi_value": Decimal("0"),
                }

            grouped_sales[key]["quantity"] += quantity
            grouped_sales[key]["sale_value"] += sale_value
            grouped_sales[key]["base_icms"] += base_icms
            grouped_sales[key]["icms_value"] += icms_value
            grouped_sales[key]["base_icms_st"] += base_icms_st
            grouped_sales[key]["icms_st_value"] += icms_st_value
            grouped_sales[key]["base_ipi"] += base_ipi
            if base_ipi > 0:
                grouped_sales[key]["ipi_rate"] = ipi_rate
            grouped_sales[key]["ipi_value"] += ipi_value

            detailed_sales.append(
                {
                    "operation_type": current_operation,
                    "document_number": current_document,
                    "document_key": current_document_key,
                    "document_date": current_document_date,
                    "document_series": current_document_series,
                    "document_model": current_document_model,
                    "document_tax_id": current_document_tax_id or current_participant_tax_id,
                    "participant_code": current_participant_code,
                    "participant_name": current_participant_name,
                    "participant_tax_id": current_participant_tax_id,
                    "item_number": get_field(fields, 2),
                    "code": code,
                    "description": description,
                    "ncm": ncm,
                    "cest": cest,
                    "cst_icms": cst_icms,
                    "cfop": cfop,
                    "quantity": quantity,
                    "icms_rate": icms_rate,
                    "sale_value": sale_value,
                    "discount_value": discount_value,
                    "base_icms": base_icms,
                    "icms_value": icms_value,
                    "base_icms_st": base_icms_st,
                    "icms_st_rate": icms_st_rate,
                    "icms_st_value": icms_st_value,
                    "base_ipi": base_ipi,
                    "ipi_rate": ipi_rate,
                    "ipi_value": ipi_value,
                }
            )

            product_key = (
                current_operation,
                current_document,
                current_document_key,
                code,
                cst_icms,
                cfop,
                icms_rate,
            )
            if product_key not in c190_product_view:
                c190_product_view[product_key] = {
                    "operation_type": current_operation,
                    "document_number": current_document,
                    "document_key": current_document_key,
                    "document_date": current_document_date,
                    "code": code,
                    "description": description,
                    "ncm": ncm,
                    "cest": cest,
                    "cst_icms": cst_icms,
                    "cfop": cfop,
                    "icms_rate": icms_rate,
                    "quantity": Decimal("0"),
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "base_icms_st": Decimal("0"),
                    "icms_st_value": Decimal("0"),
                    "base_ipi": Decimal("0"),
                    "ipi_rate": Decimal("0"),
                    "ipi_value": Decimal("0"),
                }

            c190_product_view[product_key]["quantity"] += quantity
            c190_product_view[product_key]["sale_value"] += sale_value
            c190_product_view[product_key]["base_icms"] += base_icms
            c190_product_view[product_key]["icms_value"] += icms_value
            c190_product_view[product_key]["base_icms_st"] += base_icms_st
            c190_product_view[product_key]["icms_st_value"] += icms_st_value
            c190_product_view[product_key]["base_ipi"] += base_ipi
            if base_ipi > 0:
                c190_product_view[product_key]["ipi_rate"] = ipi_rate
            c190_product_view[product_key]["ipi_value"] += ipi_value

    for code, product in list(products.items()):
        resolved_rate = product.icms_rate
        resolved_cst = product.cst_icms
        if resolved_rate is None and len(product_rates[code]) == 1:
            resolved_rate = next(iter(product_rates[code]))
        if not resolved_cst and len(product_csts[code]) == 1:
            resolved_cst = next(iter(product_csts[code]))
        if resolved_rate is not product.icms_rate or resolved_cst != product.cst_icms:
            products[code] = ProductRecord(
                code=product.code,
                description=product.description,
                ncm=product.ncm,
                cst_icms=resolved_cst,
                icms_rate=resolved_rate,
                cest=product.cest,
            )

    for item in grouped_sales.values():
        if not item["description"] and item["code"] in products:
            item["description"] = products[item["code"]].description
            item["ncm"] = products[item["code"]].ncm
        if not item["cst_icms"] and item["code"] in products:
            item["cst_icms"] = products[item["code"]].cst_icms

    product_rows = sorted(products.values(), key=lambda item: item.code)
    sales_rows = sorted(
        grouped_sales.values(),
        key=lambda item: (str(item["code"]), str(item["cst_icms"]), item["icms_rate"]),
    )
    c190_lookup: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    for item in c190_rows:
        c190_lookup[
            (
                str(item["operation_type"]),
                str(item["document_number"]),
                str(item["cst_icms"]),
                str(item["cfop"]),
                str(item["document_key"]),
                item["icms_rate"],
            )
        ] = item
    c190_product_rows: list[dict[str, object]] = []
    for item in sorted(
        c190_product_view.values(),
        key=lambda row: (
            str(row["operation_type"]),
            str(row["document_number"]),
            str(row["document_key"]),
            str(row["code"]),
            str(row["cst_icms"]),
            str(row["cfop"]),
            row["icms_rate"],
        ),
    ):
        c190_match = c190_lookup.get(
            (
                str(item["operation_type"]),
                str(item["document_number"]),
                str(item["cst_icms"]),
                str(item["cfop"]),
                str(item["document_key"]),
                item["icms_rate"],
            )
        )
        c190_product_rows.append(
            {
                **item,
                "c190_total_operation_value": c190_match["total_operation_value"] if c190_match else Decimal("0"),
                "c190_base_icms": c190_match["base_icms"] if c190_match else Decimal("0"),
                "c190_icms_value": c190_match["icms_value"] if c190_match else Decimal("0"),
                "c190_base_icms_st": c190_match["base_icms_st"] if c190_match else Decimal("0"),
                "c190_icms_st_value": c190_match["icms_st_value"] if c190_match else Decimal("0"),
                "c190_reduction_value": c190_match["reduction_value"] if c190_match else Decimal("0"),
                "c190_ipi_value": c190_match["ipi_value"] if c190_match else Decimal("0"),
            }
        )

    return product_rows, sales_rows, detailed_sales, c190_rows, c190_product_rows


def classify_entry_exit_marker(operation_type: object, cst_icms: object, cfop_value: object) -> str:
    operation = normalize_operation_type(operation_type)
    cst = normalize_cst_icms_for_sped(str(cst_icms or ""))
    cst_suffix = cst[-2:]
    cfop = re.sub(r"\D", "", str(cfop_value or "").strip())
    if not cfop:
        return ""

    if operation == "Saida":
        if cfop in {"5101", "5102", "5401", "5405", "6101", "6102", "6401", "6404", "6405"}:
            return "V"
        return ""

    if operation != "Entrada":
        return ""

    if cfop == "1253":
        return "E"
    if cst_suffix in {"40", "41", "50", "51", "90"}:
        return ""
    if cfop in {"1101", "1102", "1401", "1403", "2101", "2102", "2401"}:
        return "C"
    if cfop.startswith(("12", "14", "16", "19", "29")):
        return "O"
    return ""


def classify_entry_exit_category(operation_type: object, marker: str, cst_icms: object, cfop_value: object) -> str:
    operation = normalize_operation_type(operation_type)
    cst = normalize_cst_icms_for_sped(str(cst_icms or ""))
    cst_suffix = cst[-2:]
    cfop = re.sub(r"\D", "", str(cfop_value or "").strip())
    if operation == "Saida":
        if cfop in {"5405", "6405"} or cst_suffix == "60":
            return "Venda ST"
        if marker == "V":
            return "Venda tributada"
        return "Outros debitos de ICMS"
    if operation == "Entrada":
        if marker == "E" or cfop == "1253":
            return "Energia eletrica (previsao)"
        if cfop in {"1401", "1403", "2401"} or cst_suffix == "61":
            return "Compra ST c/ apropriacao ICMS"
        if marker == "C":
            return "Compra tributada"
        if marker == "O":
            return "Outros creditos de ICMS"
    return ""


def build_entry_exit_analysis_rows(sped_paths: list[Path]) -> tuple[list[dict[str, object]], list[list[object]], dict[str, Decimal]]:
    grouped: dict[tuple[str, str, str, Decimal], dict[str, object]] = {}
    source_files: set[str] = set()
    for sped_path in sped_paths:
        _, _, detailed_sales, c190_rows, _ = read_sped_file(sped_path)
        source_files.add(str(sped_path))
        c170_ipi_map: dict[tuple[str, str, str, Decimal], dict[str, Decimal]] = defaultdict(
            lambda: {"base_ipi": Decimal("0"), "ipi_value": Decimal("0")}
        )
        for item in detailed_sales:
            operation = normalize_operation_type(item.get("operation_type", ""))
            if operation not in {"Entrada", "Saida"}:
                continue
            cst = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")).strip())
            cfop = str(item.get("cfop", "")).strip()
            rate = Decimal(item.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
            key = (operation, cst, cfop, rate)
            c170_ipi_map[key]["base_ipi"] += Decimal(item.get("base_ipi", Decimal("0")))
            c170_ipi_map[key]["ipi_value"] += Decimal(item.get("ipi_value", Decimal("0")))

        for row in c190_rows:
            operation = normalize_operation_type(row.get("operation_type", ""))
            if operation not in {"Entrada", "Saida"}:
                continue
            cst = normalize_cst_icms_for_sped(str(row.get("cst_icms", "")).strip())
            cfop = str(row.get("cfop", "")).strip()
            rate = Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
            key = (operation, cst, cfop, rate)
            bucket = grouped.setdefault(
                key,
                {
                    "operation_type": operation,
                    "marker": classify_entry_exit_marker(operation, cst, cfop),
                    "cst_icms": cst,
                    "cfop": cfop,
                    "icms_rate": rate,
                    "effective_rate": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "total_operation_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "base_icms_st": Decimal("0"),
                    "icms_st_value": Decimal("0"),
                    "base_ipi": Decimal("0"),
                    "ipi_rate": Decimal("0"),
                    "ipi_value": Decimal("0"),
                    "source_files": set(),
                    "row_count": 0,
                },
            )
            bucket["icms_value"] += Decimal(row.get("icms_value", Decimal("0")))
            bucket["total_operation_value"] += Decimal(row.get("total_operation_value", Decimal("0")))
            bucket["base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
            bucket["base_icms_st"] += Decimal(row.get("base_icms_st", Decimal("0")))
            bucket["icms_st_value"] += Decimal(row.get("icms_st_value", Decimal("0")))
            bucket["ipi_value"] += Decimal(row.get("ipi_value", Decimal("0")))
            bucket["source_files"].add(str(sped_path))
            bucket["row_count"] = int(bucket["row_count"]) + 1

        for key, ipi_totals in c170_ipi_map.items():
            bucket = grouped.get(key)
            if not bucket:
                continue
            bucket["base_ipi"] += ipi_totals["base_ipi"]
            if Decimal(bucket.get("ipi_value", Decimal("0"))) == Decimal("0"):
                bucket["ipi_value"] += ipi_totals["ipi_value"]

    detail_rows: list[dict[str, object]] = []
    for row in grouped.values():
        operation_total = Decimal(row["total_operation_value"])
        icms_value = Decimal(row["icms_value"])
        row["effective_rate"] = (
            (icms_value * Decimal("100") / operation_total).quantize(Decimal("0.01"))
            if operation_total > 0
            else Decimal("0.00")
        )
        base_ipi = Decimal(row.get("base_ipi", Decimal("0")))
        ipi_value = Decimal(row.get("ipi_value", Decimal("0")))
        row["ipi_rate"] = (
            (ipi_value * Decimal("100") / base_ipi).quantize(Decimal("0.01"))
            if base_ipi > 0
            else Decimal("0.00")
        )
        row["category"] = classify_entry_exit_category(row["operation_type"], str(row["marker"]), row["cst_icms"], row["cfop"])
        row["source_file_count"] = len(row["source_files"])
        detail_rows.append(row)

    detail_rows.sort(
        key=lambda row: (
            0 if row["operation_type"] == "Saida" else 1,
            str(row["cst_icms"]),
            str(row["cfop"]),
            Decimal(row["icms_rate"]),
        )
    )

    def sum_rows(rows: list[dict[str, object]]) -> dict[str, Decimal]:
        return {
            "icms_value": sum((Decimal(row["icms_value"]) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "total_operation_value": sum((Decimal(row["total_operation_value"]) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "base_icms": sum((Decimal(row["base_icms"]) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "base_icms_st": sum((Decimal(row.get("base_icms_st", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "icms_st_value": sum((Decimal(row.get("icms_st_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "base_ipi": sum((Decimal(row.get("base_ipi", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
            "ipi_value": sum((Decimal(row.get("ipi_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        }

    def excel_detail(row: dict[str, object]) -> list[object]:
        return [
            "",
            row["marker"],
            row["cst_icms"],
            row["cfop"],
            Decimal(row["icms_rate"]),
            Decimal(row["effective_rate"]),
            Decimal(row["icms_value"]).quantize(Decimal("0.01")),
            Decimal(row["total_operation_value"]).quantize(Decimal("0.01")),
            Decimal(row["base_icms"]).quantize(Decimal("0.01")),
            Decimal(row.get("base_icms_st", Decimal("0"))).quantize(Decimal("0.01")),
            Decimal(row.get("icms_st_value", Decimal("0"))).quantize(Decimal("0.01")),
            Decimal(row.get("base_ipi", Decimal("0"))).quantize(Decimal("0.01")),
            Decimal(row.get("ipi_rate", Decimal("0"))).quantize(Decimal("0.01")),
            Decimal(row.get("ipi_value", Decimal("0"))).quantize(Decimal("0.01")),
        ]

    def excel_total(label: str, totals: dict[str, Decimal]) -> list[object]:
        return [
            "",
            "",
            label,
            "",
            "",
            "",
            totals["icms_value"],
            totals["total_operation_value"],
            totals["base_icms"],
            totals["base_icms_st"],
            totals["icms_st_value"],
            totals["base_ipi"],
            "",
            totals["ipi_value"],
        ]

    excel_rows: list[list[object]] = []
    headers = [
        "",
        "",
        "CST",
        "CFOP",
        "Aliq ICMS",
        "Aliq Efetiva",
        "Valor ICMS",
        "Total Operacao",
        "Base ICMS",
        "Base ICMS ST",
        "Valor ICMS ST",
        "Base IPI",
        "Aliq IPI",
        "Valor IPI",
    ]
    saida_rows = [row for row in detail_rows if row["operation_type"] == "Saida"]
    entrada_rows = [row for row in detail_rows if row["operation_type"] == "Entrada"]
    excel_rows.extend(excel_detail(row) for row in saida_rows)
    saida_total = sum_rows(saida_rows)
    excel_rows.append(excel_total("Total", saida_total))
    for label in ("Venda tributada", "Venda ST", "Outros debitos de ICMS"):
        excel_rows.append(excel_total(label, sum_rows([row for row in saida_rows if row["category"] == label])))
    saida_classified_total = sum_rows([row for row in saida_rows if row["category"]])
    excel_rows.append(excel_total("Total saidas classificado", saida_classified_total))
    excel_rows.append([""] * len(headers))
    excel_rows.append(headers)
    excel_rows.extend(excel_detail(row) for row in entrada_rows)
    entrada_total = sum_rows(entrada_rows)
    excel_rows.append(excel_total("Total", entrada_total))
    for label in (
        "Compra tributada",
        "Compra ST c/ apropriacao ICMS",
        "Outros creditos de ICMS",
        "Energia eletrica (previsao)",
    ):
        excel_rows.append(excel_total(label, sum_rows([row for row in entrada_rows if row["category"] == label])))
    entrada_classified_total = sum_rows([row for row in entrada_rows if row["category"]])
    excel_rows.append(excel_total("Total entradas classificado", entrada_classified_total))
    recolher = (saida_classified_total["icms_value"] - entrada_classified_total["icms_value"]).quantize(Decimal("0.01"))
    venda_base = sum_rows([row for row in saida_rows if row["category"] in {"Venda tributada", "Venda ST"}])["total_operation_value"]
    percent_sale = (recolher * Decimal("100") / venda_base).quantize(Decimal("0.01")) if venda_base > 0 else Decimal("0.00")
    excel_rows.append(["", "", "", "", "", "A recolher", recolher, "", "", "", "", "", "", ""])
    excel_rows.append(["", "", "", "", "", "% sobre venda", f"{format_decimal_sped(percent_sale)}%", "", "", "", "", "", "", ""])

    totals = {
        "saida_icms": saida_total["icms_value"],
        "saida_operation": saida_total["total_operation_value"],
        "entrada_icms": entrada_total["icms_value"],
        "entrada_operation": entrada_total["total_operation_value"],
        "recolher": recolher,
        "percent_sale": percent_sale,
        "source_files": Decimal(len(source_files)),
    }
    return detail_rows, excel_rows, totals


def write_entry_exit_analysis_excel(output_path: Path, excel_rows: list[list[object]]) -> None:
    headers = [
        "",
        "",
        "CST",
        "CFOP",
        "Aliq ICMS",
        "Aliq Efetiva",
        "Valor ICMS",
        "Total Operacao",
        "Base ICMS",
        "Base ICMS ST",
        "Valor ICMS ST",
        "Base IPI",
        "Aliq IPI",
        "Valor IPI",
    ]
    write_simple_excel_workbook(
        output_path,
        [
            (
                "Resumo Saidas",
                headers,
                excel_rows,
                {"include_total": False},
            )
        ],
    )


def sum_entry_exit_analysis_rows(rows: list[dict[str, object]]) -> dict[str, Decimal]:
    return {
        "icms_value": sum((Decimal(row.get("icms_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "total_operation_value": sum((Decimal(row.get("total_operation_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "base_icms": sum((Decimal(row.get("base_icms", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "base_icms_st": sum((Decimal(row.get("base_icms_st", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "icms_st_value": sum((Decimal(row.get("icms_st_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "base_ipi": sum((Decimal(row.get("base_ipi", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
        "ipi_value": sum((Decimal(row.get("ipi_value", Decimal("0"))) for row in rows), Decimal("0")).quantize(Decimal("0.01")),
    }


def build_entry_exit_footer_rows(rows: list[dict[str, object]], operation_type: str) -> list[dict[str, object]]:
    operation = normalize_operation_type(operation_type)
    total = sum_entry_exit_analysis_rows(rows)
    footer_rows: list[dict[str, object]] = [
        {"label": "Total", **total},
    ]
    labels = (
        ("Venda tributada", "Venda ST", "Outros debitos de ICMS")
        if operation == "Saida"
        else (
            "Compra tributada",
            "Compra ST c/ apropriacao ICMS",
            "Outros creditos de ICMS",
            "Energia eletrica (previsao)",
        )
    )
    for label in labels:
        footer_rows.append(
            {
                "label": label,
                **sum_entry_exit_analysis_rows([row for row in rows if str(row.get("category", "")) == label]),
            }
        )
    footer_rows.append(
        {
            "label": f"Total {operation.lower()}s classificado" if operation else "Total classificado",
            **sum_entry_exit_analysis_rows([row for row in rows if str(row.get("category", "")).strip()]),
        }
    )
    return footer_rows


def format_grouped_line_numbers(grouped_rows: list[dict[str, object]], max_items: int = 8) -> str:
    line_numbers = [str(row.get("line_number", "")).strip() for row in grouped_rows if str(row.get("line_number", "")).strip()]
    if not line_numbers:
        return ""
    preview = ", ".join(line_numbers[:max_items])
    if len(line_numbers) > max_items:
        preview += ", ..."
    return preview


def build_credit_diagnostic_product_sheet_payload(
    detail_rows: list[dict[str, object]],
) -> tuple[list[list[object]], list[dict[str, object]]]:
    product_rows: list[list[object]] = []
    footer_rows: list[dict[str, object]] = []

    for detail_row in detail_rows:
        period = str(detail_row.get("period", ""))
        reason = str(detail_row.get("reason", ""))
        source_register = str(detail_row.get("source_register", ""))
        cst_icms = str(detail_row.get("cst_icms", ""))
        cfop = str(detail_row.get("cfop", ""))
        nominal_rate = Decimal(detail_row.get("icms_rate", Decimal("0")))

        detail_operation_total = Decimal(detail_row.get("total_operation_value", Decimal("0")))
        detail_base_total = Decimal(detail_row.get("base_icms", Decimal("0")))
        detail_icms_total = Decimal(detail_row.get("icms_value", Decimal("0")))
        detail_effective_rate = compute_display_icms_rate(
            nominal_rate,
            detail_operation_total,
            detail_base_total,
            detail_icms_total,
        )

        product_operation_total = Decimal("0")
        product_base_total = Decimal("0")
        product_icms_total = Decimal("0")

        for product_row in detail_row.get("product_rows", []):
            product_sale = Decimal(product_row.get("sale_value", Decimal("0")))
            product_base = Decimal(product_row.get("base_icms", Decimal("0")))
            product_icms = Decimal(product_row.get("icms_value", Decimal("0")))
            product_ratio = (
                (product_base * Decimal("100") / product_sale).quantize(Decimal("0.01"))
                if product_sale > 0
                else Decimal("0.00")
            )
            product_effective_rate = compute_display_icms_rate(
                Decimal(product_row.get("icms_rate", Decimal("0"))),
                product_sale,
                product_base,
                product_icms,
            )
            product_rows.append(
                [
                    product_row.get("period", period),
                    reason,
                    source_register,
                    cst_icms,
                    cfop,
                    product_row.get("code", ""),
                    product_row.get("description", ""),
                    product_row.get("ncm", ""),
                    Decimal(product_row.get("icms_rate", nominal_rate)),
                    product_effective_rate,
                    product_row.get("row_count", 0),
                    product_sale,
                    product_base,
                    product_ratio,
                    product_icms,
                ]
            )
            product_operation_total += product_sale
            product_base_total += product_base
            product_icms_total += product_icms

        operation_diff = detail_operation_total - product_operation_total
        base_diff = detail_base_total - product_base_total
        icms_diff = detail_icms_total - product_icms_total
        if not any(value != 0 for value in (operation_diff, base_diff, icms_diff)):
            continue

        grouped_rows = list(detail_row.get("grouped_rows", []))
        line_numbers = format_grouped_line_numbers(grouped_rows)
        adjustment_label = "SEM PRODUTO IDENTIFICADO" if not detail_row.get("product_rows") else "AJUSTE RESUMO X ITENS"
        product_rows.append(
            [
                period,
                reason,
                source_register,
                cst_icms,
                cfop,
                "",
                adjustment_label,
                line_numbers,
                nominal_rate,
                detail_effective_rate,
                0,
                operation_diff,
                base_diff,
                "",
                icms_diff,
            ]
        )

    return product_rows, footer_rows


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


def summarize_entry_analysis(entry_items: list[dict[str, object]]) -> dict[str, object]:
    cfops = sorted({str(item.get("cfop", "")).strip() for item in entry_items if str(item.get("cfop", "")).strip()})
    csts = sorted({str(item.get("cst_icms", "")).strip() for item in entry_items if str(item.get("cst_icms", "")).strip()})
    total_credit = sum((Decimal(item.get("icms_value", Decimal("0"))) for item in entry_items), Decimal("0"))

    status_parts: list[str] = []
    if not entry_items:
        status_parts.append("Sem entrada")
    else:
        if not cfops:
            status_parts.append("Erro: CFOP vazio")
        elif len(cfops) > 1:
            status_parts.append("Erro: multiplos CFOPs")
        if not csts:
            status_parts.append("Erro: CST vazio")
        elif len(csts) > 1:
            status_parts.append("Erro: multiplos CSTs")
        if total_credit < Decimal("0"):
            status_parts.append("Erro: credito negativo")
        elif total_credit == Decimal("0"):
            status_parts.append("Sem credito")
        if not status_parts:
            status_parts.append("Ok")

    return {
        "cfop": " | ".join(cfops),
        "cst": " | ".join(csts),
        "credit": total_credit,
        "has_entry": bool(entry_items),
        "status": " | ".join(status_parts),
    }


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


def build_product_monthly_summary_rows(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> list[dict[str, object]]:
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    totals_by_code: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))

    for row in filtered_rows:
        code = str(row.get("code", "")).strip()
        period = str(row.get("period", "")).strip()
        if not code or not period:
            continue

        key = (code, period)
        bucket = grouped_rows.setdefault(
            key,
            {
                "period": period,
                "operation_type": operation_type,
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "document_keys": set(),
                "launch_count": 0,
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )

        if not bucket["description"] and str(row.get("description", "")).strip():
            bucket["description"] = str(row.get("description", "")).strip()
        if not bucket["ncm"] and str(row.get("ncm", "")).strip():
            bucket["ncm"] = str(row.get("ncm", "")).strip()

        for cfop in str(row.get("cfop", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                bucket["cfops"].add(cfop_text)
        for cst in str(row.get("cst_icms", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                bucket["csts"].add(cst_text)

        bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        bucket["base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        bucket["icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

        launch_details = list(row.get("launch_details", []))
        bucket["launch_count"] += len(launch_details)
        for detail in launch_details:
            document_identity = (
                normalize_document_key(str(detail.get("document_key", "")))
                or str(detail.get("document_number", "")).strip()
            )
            if document_identity:
                bucket["document_keys"].add(document_identity)

        totals_by_code[code] += Decimal(row.get("sale_value", Decimal("0")))

    curve_labels = calculate_abc_curve_labels(totals_by_code)
    monthly_rows: list[dict[str, object]] = []
    for (_, _), item in grouped_rows.items():
        base_icms = Decimal(item["base_icms"])
        icms_value = Decimal(item["icms_value"])
        monthly_rows.append(
            {
                "period": item["period"],
                "operation_type": item["operation_type"],
                "curve_abc": curve_labels.get(str(item["code"]).strip(), "C"),
                "code": item["code"],
                "description": item["description"],
                "ncm": item["ncm"],
                "cfops": " | ".join(sorted(item["cfops"])),
                "csts": " | ".join(sorted(item["csts"])),
                "document_count": len(item["document_keys"]),
                "launch_count": int(item["launch_count"]),
                "sale_value": Decimal(item["sale_value"]),
                "base_icms": base_icms,
                "icms_value": icms_value,
                "calculated_icms_rate": (
                    (icms_value * Decimal("100") / base_icms).quantize(Decimal("0.01"))
                    if base_icms > 0
                    else Decimal("0.00")
                ),
            }
        )

    monthly_rows.sort(
        key=lambda item: (
            normalize_text(str(item.get("period", ""))),
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("code", ""))),
        )
    )
    return monthly_rows


def period_label_sort_key(value: object) -> tuple[int, int, str]:
    text = str(value or "").strip()
    match = re.fullmatch(r"(\d{2})/(\d{4})", text)
    if match:
        return (int(match.group(2)), int(match.group(1)), text)
    return (9999, 99, normalize_text(text))


def build_product_monthly_linear_dataset(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]]]:
    monthly_rows = build_product_monthly_summary_rows(filtered_rows, operation_type)
    periods = sorted({str(row.get("period", "")).strip() for row in monthly_rows if str(row.get("period", "")).strip()}, key=period_label_sort_key)
    grouped_products: dict[str, dict[str, object]] = {}

    for row in monthly_rows:
        code = str(row.get("code", "")).strip()
        if not code:
            continue
        product_bucket = grouped_products.setdefault(
            code,
            {
                "curve_abc": str(row.get("curve_abc", "C")).strip() or "C",
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "periods": {},
                "total_sale_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        if not product_bucket["description"] and str(row.get("description", "")).strip():
            product_bucket["description"] = str(row.get("description", "")).strip()
        if not product_bucket["ncm"] and str(row.get("ncm", "")).strip():
            product_bucket["ncm"] = str(row.get("ncm", "")).strip()

        for cfop in str(row.get("cfops", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                product_bucket["cfops"].add(cfop_text)
        for cst in str(row.get("csts", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                product_bucket["csts"].add(cst_text)

        period = str(row.get("period", "")).strip()
        product_bucket["periods"][period] = {
            "sale_value": Decimal(row.get("sale_value", Decimal("0"))),
            "base_icms": Decimal(row.get("base_icms", Decimal("0"))),
            "icms_value": Decimal(row.get("icms_value", Decimal("0"))),
            "calculated_icms_rate": Decimal(row.get("calculated_icms_rate", Decimal("0"))),
        }
        product_bucket["total_sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        product_bucket["total_base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        product_bucket["total_icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

    fixed_headers = ["Curva ABC", "Codigo", "Descricao", "NCM", "CFOPs", "CSTs"]
    dynamic_headers: list[str] = []
    for period in periods:
        dynamic_headers.extend(
            [
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Valor ICMS",
                f"{period} Aliquota",
            ]
        )
    export_headers = [
        *fixed_headers,
        *dynamic_headers,
        "Total Valor Operacao",
        "Total Base ICMS",
        "Total Valor ICMS",
        "Aliquota Total",
    ]

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for product in sorted(
        grouped_products.values(),
        key=lambda item: (
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        ),
    ):
        total_base_icms = Decimal(product["total_base_icms"])
        total_icms_value = Decimal(product["total_icms_value"])
        total_rate = (total_icms_value * Decimal("100") / total_base_icms).quantize(Decimal("0.01")) if total_base_icms > 0 else Decimal("0.00")
        values: list[object] = [
            product["curve_abc"],
            product["code"],
            product["description"],
            product["ncm"],
            " | ".join(sorted(product["cfops"])),
            " | ".join(sorted(product["csts"])),
        ]
        for period in periods:
            period_data = product["periods"].get(period)
            if period_data is None:
                values.extend([Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    Decimal(period_data["sale_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["icms_value"]),
                    Decimal(period_data["calculated_icms_rate"]),
                ]
            )
        values.extend(
            [
                Decimal(product["total_sale_value"]),
                total_base_icms,
                total_icms_value,
                total_rate,
            ]
        )
        display_rows.append({"curve_abc": product["curve_abc"], "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows


def build_contrib_operation_summary_rows(
    detailed_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[dict[str, object]], dict[str, Decimal]]:
    grouped_rows: dict[tuple[str, str, str, Decimal, Decimal], dict[str, object]] = {}
    normalized_operation = normalize_operation_type(operation_type)

    for source_row in detailed_rows:
        if normalize_operation_type(str(source_row.get("operation_type", ""))) != normalized_operation:
            continue
        cst_pis = str(source_row.get("cst_pis", "")).strip()
        cst_cofins = str(source_row.get("cst_cofins", "")).strip()
        cfop = str(source_row.get("cfop", "")).strip()
        aliquota_pis = source_row.get("aliquota_pis", Decimal("0"))
        aliquota_cofins = source_row.get("aliquota_cofins", Decimal("0"))
        aliquota_pis = aliquota_pis if isinstance(aliquota_pis, Decimal) else Decimal("0")
        aliquota_cofins = aliquota_cofins if isinstance(aliquota_cofins, Decimal) else Decimal("0")
        key = (
            cst_pis,
            cst_cofins,
            cfop,
            aliquota_pis.quantize(Decimal("0.01")),
            aliquota_cofins.quantize(Decimal("0.01")),
        )
        sale_value = source_row.get("sale_value", Decimal("0"))
        base_pis = source_row.get("base_pis", Decimal("0"))
        base_cofins = source_row.get("base_cofins", Decimal("0"))
        pis_value = source_row.get("pis_value", Decimal("0"))
        cofins_value = source_row.get("cofins_value", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_pis = base_pis if isinstance(base_pis, Decimal) else Decimal("0")
        base_cofins = base_cofins if isinstance(base_cofins, Decimal) else Decimal("0")
        pis_value = pis_value if isinstance(pis_value, Decimal) else Decimal("0")
        cofins_value = cofins_value if isinstance(cofins_value, Decimal) else Decimal("0")

        if key not in grouped_rows:
            grouped_rows[key] = {
                "operation_type": normalized_operation,
                "cst_pis": cst_pis,
                "cst_cofins": cst_cofins,
                "cfop": cfop,
                "aliquota_pis": key[3],
                "aliquota_cofins": key[4],
                "sale_value": Decimal("0"),
                "base_pis": Decimal("0"),
                "pis_value": Decimal("0"),
                "base_cofins": Decimal("0"),
                "cofins_value": Decimal("0"),
                "document_keys": set(),
                "launch_count": 0,
            }

        grouped_rows[key]["sale_value"] += sale_value
        grouped_rows[key]["base_pis"] += base_pis
        grouped_rows[key]["pis_value"] += pis_value
        grouped_rows[key]["base_cofins"] += base_cofins
        grouped_rows[key]["cofins_value"] += cofins_value
        grouped_rows[key]["launch_count"] += 1
        document_identity = normalize_document_key(str(source_row.get("document_key", ""))) or str(source_row.get("document_number", "")).strip()
        if document_identity:
            grouped_rows[key]["document_keys"].add(document_identity)

    rows = sorted(
        grouped_rows.values(),
        key=lambda row: (
            str(row["cst_pis"]),
            str(row["cst_cofins"]),
            str(row["cfop"]),
            Decimal(row["aliquota_pis"]),
            Decimal(row["aliquota_cofins"]),
        ),
    )
    totals = {
        "sale_value": sum((Decimal(row["sale_value"]) for row in rows), Decimal("0")),
        "base_pis": sum((Decimal(row["base_pis"]) for row in rows), Decimal("0")),
        "pis_value": sum((Decimal(row["pis_value"]) for row in rows), Decimal("0")),
        "base_cofins": sum((Decimal(row["base_cofins"]) for row in rows), Decimal("0")),
        "cofins_value": sum((Decimal(row["cofins_value"]) for row in rows), Decimal("0")),
    }
    return rows, totals


def build_contrib_operation_launch_details_map(
    detailed_rows: list[dict[str, object]],
    operation_type: str,
) -> dict[tuple[str, str, str, Decimal, Decimal], list[dict[str, object]]]:
    grouped_details: dict[tuple[str, str, str, Decimal, Decimal], list[dict[str, object]]] = defaultdict(list)
    normalized_operation = normalize_operation_type(operation_type)

    for item in detailed_rows:
        if normalize_operation_type(str(item.get("operation_type", ""))) != normalized_operation:
            continue
        cst_pis = str(item.get("cst_pis", "")).strip()
        cst_cofins = str(item.get("cst_cofins", "")).strip()
        cfop = str(item.get("cfop", "")).strip()
        aliquota_pis = item.get("aliquota_pis", Decimal("0"))
        aliquota_cofins = item.get("aliquota_cofins", Decimal("0"))
        aliquota_pis = aliquota_pis if isinstance(aliquota_pis, Decimal) else Decimal("0")
        aliquota_cofins = aliquota_cofins if isinstance(aliquota_cofins, Decimal) else Decimal("0")
        key = (
            cst_pis,
            cst_cofins,
            cfop,
            aliquota_pis.quantize(Decimal("0.01")),
            aliquota_cofins.quantize(Decimal("0.01")),
        )
        grouped_details[key].append(dict(item))

    return grouped_details


def build_contrib_product_monthly_summary_rows(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> list[dict[str, object]]:
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    totals_by_code: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))

    for row in filtered_rows:
        code = str(row.get("code", "")).strip()
        period = str(row.get("period", "")).strip()
        if not code or not period:
            continue

        key = (code, period)
        bucket = grouped_rows.setdefault(
            key,
            {
                "period": period,
                "operation_type": operation_type,
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "document_keys": set(),
                "launch_count": 0,
                "sale_value": Decimal("0"),
                "base_pis": Decimal("0"),
                "pis_value": Decimal("0"),
                "base_cofins": Decimal("0"),
                "cofins_value": Decimal("0"),
            },
        )

        if not bucket["description"] and str(row.get("description", "")).strip():
            bucket["description"] = str(row.get("description", "")).strip()
        if not bucket["ncm"] and str(row.get("ncm", "")).strip():
            bucket["ncm"] = str(row.get("ncm", "")).strip()
        for cfop in str(row.get("cfop", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                bucket["cfops"].add(cfop_text)
        for cst in (str(row.get("cst_pis", "")).split("|") + str(row.get("cst_cofins", "")).split("|")):
            cst_text = cst.strip()
            if cst_text:
                bucket["csts"].add(cst_text)

        bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        bucket["base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
        bucket["pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
        bucket["base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
        bucket["cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

        launch_details = list(row.get("launch_details", []))
        bucket["launch_count"] += len(launch_details)
        for detail in launch_details:
            document_identity = normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            if document_identity:
                bucket["document_keys"].add(document_identity)

        totals_by_code[code] += Decimal(row.get("sale_value", Decimal("0")))

    curve_labels = calculate_abc_curve_labels(totals_by_code)
    monthly_rows: list[dict[str, object]] = []
    for item in grouped_rows.values():
        monthly_rows.append(
            {
                "period": item["period"],
                "operation_type": item["operation_type"],
                "curve_abc": curve_labels.get(str(item["code"]).strip(), "C"),
                "code": item["code"],
                "description": item["description"],
                "ncm": item["ncm"],
                "cfops": " | ".join(sorted(item["cfops"])),
                "csts": " | ".join(sorted(item["csts"])),
                "document_count": len(item["document_keys"]),
                "launch_count": int(item["launch_count"]),
                "sale_value": Decimal(item["sale_value"]),
                "base_pis": Decimal(item["base_pis"]),
                "pis_value": Decimal(item["pis_value"]),
                "base_cofins": Decimal(item["base_cofins"]),
                "cofins_value": Decimal(item["cofins_value"]),
            }
        )

    monthly_rows.sort(
        key=lambda item: (
            normalize_text(str(item.get("period", ""))),
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        )
    )
    return monthly_rows


def build_contrib_product_monthly_linear_dataset(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]]]:
    monthly_rows = build_contrib_product_monthly_summary_rows(filtered_rows, operation_type)
    periods = sorted({str(row.get("period", "")).strip() for row in monthly_rows if str(row.get("period", "")).strip()}, key=period_label_sort_key)
    grouped_products: dict[str, dict[str, object]] = {}

    for row in monthly_rows:
        code = str(row.get("code", "")).strip()
        if not code:
            continue
        product_bucket = grouped_products.setdefault(
            code,
            {
                "curve_abc": str(row.get("curve_abc", "C")).strip() or "C",
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "periods": {},
                "total_sale_value": Decimal("0"),
                "total_base_pis": Decimal("0"),
                "total_pis_value": Decimal("0"),
                "total_base_cofins": Decimal("0"),
                "total_cofins_value": Decimal("0"),
            },
        )
        for cfop in str(row.get("cfops", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                product_bucket["cfops"].add(cfop_text)
        for cst in str(row.get("csts", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                product_bucket["csts"].add(cst_text)
        period = str(row.get("period", "")).strip()
        product_bucket["periods"][period] = {
            "sale_value": Decimal(row.get("sale_value", Decimal("0"))),
            "base_pis": Decimal(row.get("base_pis", Decimal("0"))),
            "pis_value": Decimal(row.get("pis_value", Decimal("0"))),
            "base_cofins": Decimal(row.get("base_cofins", Decimal("0"))),
            "cofins_value": Decimal(row.get("cofins_value", Decimal("0"))),
        }
        product_bucket["total_sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        product_bucket["total_base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
        product_bucket["total_pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
        product_bucket["total_base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
        product_bucket["total_cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

    fixed_headers = ["Curva ABC", "Codigo", "Descricao", "NCM", "CFOPs", "CSTs"]
    dynamic_headers: list[str] = []
    for period in periods:
        dynamic_headers.extend(
            [
                f"{period} Valor Operacao",
                f"{period} Base PIS",
                f"{period} Valor PIS",
                f"{period} Base COFINS",
                f"{period} Valor COFINS",
            ]
        )
    export_headers = [
        *fixed_headers,
        *dynamic_headers,
        "Total Valor Operacao",
        "Total Base PIS",
        "Total Valor PIS",
        "Total Base COFINS",
        "Total Valor COFINS",
    ]

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for product in sorted(
        grouped_products.values(),
        key=lambda item: (
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        ),
    ):
        values: list[object] = [
            product["curve_abc"],
            product["code"],
            product["description"],
            product["ncm"],
            " | ".join(sorted(product["cfops"])),
            " | ".join(sorted(product["csts"])),
        ]
        for period in periods:
            period_data = product["periods"].get(
                period,
                {
                    "sale_value": Decimal("0"),
                    "base_pis": Decimal("0"),
                    "pis_value": Decimal("0"),
                    "base_cofins": Decimal("0"),
                    "cofins_value": Decimal("0"),
                },
            )
            values.extend(
                [
                    Decimal(period_data["sale_value"]),
                    Decimal(period_data["base_pis"]),
                    Decimal(period_data["pis_value"]),
                    Decimal(period_data["base_cofins"]),
                    Decimal(period_data["cofins_value"]),
                ]
            )
        values.extend(
            [
                Decimal(product["total_sale_value"]),
                Decimal(product["total_base_pis"]),
                Decimal(product["total_pis_value"]),
                Decimal(product["total_base_cofins"]),
                Decimal(product["total_cofins_value"]),
            ]
        )
        display_rows.append({"curve_abc": product["curve_abc"], "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows


def build_monthly_aliquota_divergence_rows(
    headers: list[str],
    rows: list[list[object]],
    periods: list[str],
) -> list[list[object]]:
    fixed_column_count = 6
    monthly_block_size = 4
    divergence_rows: list[list[object]] = []

    for row in rows:
        valid_rates: list[Decimal] = []
        for period_index, _period in enumerate(periods):
            start = fixed_column_count + (period_index * monthly_block_size)
            if start + 3 >= len(row):
                continue
            sale_value = row[start]
            base_icms = row[start + 1]
            icms_value = row[start + 2]
            calculated_rate = row[start + 3]
            sale_value = sale_value if isinstance(sale_value, Decimal) else parse_decimal(str(sale_value or "0"))
            base_icms = base_icms if isinstance(base_icms, Decimal) else parse_decimal(str(base_icms or "0"))
            icms_value = icms_value if isinstance(icms_value, Decimal) else parse_decimal(str(icms_value or "0"))
            calculated_rate = calculated_rate if isinstance(calculated_rate, Decimal) else parse_decimal(str(calculated_rate or "0"))

            if base_icms <= 0 and sale_value <= 0 and icms_value <= 0:
                continue
            if base_icms <= 0:
                continue
            valid_rates.append(calculated_rate.quantize(Decimal("0.01")))

        distinct_rates = {rate for rate in valid_rates}
        if len(distinct_rates) > 1:
            divergence_rows.append(list(row))

    return divergence_rows


def build_summary_operation_totals_by_period(
    sped_paths: list[Path],
    operation_type: str,
) -> dict[str, dict[str, Decimal]]:
    totals_by_period: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: {
            "total_operation_value": Decimal("0"),
            "base_icms": Decimal("0"),
            "icms_value": Decimal("0"),
        }
    )
    normalized_operation = normalize_operation_type(operation_type)
    for sped_path in sped_paths:
        if not sped_path.exists():
            continue
        period_label = infer_sped_period_label(sped_path, [])
        summary_rows = read_sped_summary_register_rows(sped_path)
        for row in summary_rows:
            if normalize_operation_type(str(row.get("operation_type", ""))) != normalized_operation:
                continue
            totals_by_period[period_label]["total_operation_value"] += Decimal(row.get("total_operation_value", Decimal("0")))
            totals_by_period[period_label]["base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
            totals_by_period[period_label]["icms_value"] += Decimal(row.get("icms_value", Decimal("0")))
    return totals_by_period


def classify_credit_base_reason(
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)[-2:]
    if total_operation_value <= 0 and base_icms <= 0 and icms_value <= 0:
        return "Sem valores fiscais"
    if base_icms <= 0 and icms_value > 0:
        return "Credito sem base fiscal"
    if normalized_cst in {"40", "41", "50"}:
        return "CST sem direito a credito"
    if normalized_cst in {"60", "61"}:
        return "Substituicao tributaria / ICMS ST"
    if normalized_cst in {"90"} and base_icms <= 0:
        return "Outras entradas sem base de credito"
    if total_operation_value > 0 and base_icms <= 0:
        return "Operacao sem base de credito"
    if total_operation_value > 0 and base_icms < total_operation_value:
        return "Base reduzida em relacao ao valor da operacao"
    if source_register in {"C590", "D190", "D590", "D730"}:
        return "Servico / energia / comunicacao"
    return "Base integral"


def classify_debit_base_reason(
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)[-2:]
    if total_operation_value <= 0 and base_icms <= 0 and icms_value <= 0:
        return "Sem valores fiscais"
    if base_icms <= 0 and icms_value > 0:
        return "Debito sem base fiscal"
    if normalized_cst in {"40", "41", "50"}:
        return "CST sem debito do imposto"
    if normalized_cst in {"60", "61"}:
        return "Substituicao tributaria / ICMS ST"
    if normalized_cst in {"90"} and base_icms <= 0:
        return "Outras saidas sem base de debito"
    if total_operation_value > 0 and base_icms <= 0:
        return "Operacao sem base de debito"
    if total_operation_value > 0 and icms_value <= 0:
        return "Operacao sem debito"
    if total_operation_value > 0 and base_icms < total_operation_value:
        return "Base reduzida em relacao ao valor da operacao"
    if source_register in {"C590", "D190", "D590", "D730"}:
        return "Servico / energia / comunicacao"
    return "Base integral"


def classify_icms_diagnostic_reason(
    operation_type: str,
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    if normalize_operation_type(operation_type) == "saida":
        return classify_debit_base_reason(source_register, cst_icms, total_operation_value, base_icms, icms_value)
    return classify_credit_base_reason(source_register, cst_icms, total_operation_value, base_icms, icms_value)


def build_credit_diagnostic_datasets(
    summary_register_rows: list[dict[str, object]],
    operation_type: str,
    cst_filter_values: set[str],
    cfop_filter_values: set[str],
    detailed_sales: list[dict[str, object]] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, Decimal]]:
    normalized_operation = normalize_operation_type(operation_type)
    normalized_csts = {normalize_cst_icms_for_sped(str(value)) for value in cst_filter_values if str(value).strip()}
    normalized_cfops = {str(value).strip() for value in cfop_filter_values if str(value).strip()}

    detail_groups: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    summary_groups: dict[tuple[str, str], dict[str, object]] = {}
    totals = {
        "total_operation_value": Decimal("0"),
        "base_icms": Decimal("0"),
        "icms_value": Decimal("0"),
        "base_gap": Decimal("0"),
    }

    for row in summary_register_rows:
        if normalize_operation_type(str(row.get("operation_type", ""))) != normalized_operation:
            continue
        cst_value = normalize_cst_icms_for_sped(str(row.get("cst_icms", "")))
        cfop_value = str(row.get("cfop", "")).strip()
        if normalized_csts and cst_value not in normalized_csts:
            continue
        if normalized_cfops and cfop_value not in normalized_cfops:
            continue

        source_register = str(row.get("source_register", "")).strip()
        period = str(row.get("period", "")).strip()
        icms_rate = Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
        total_operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
        base_icms = Decimal(row.get("base_icms", Decimal("0")))
        icms_value = Decimal(row.get("icms_value", Decimal("0")))
        reason = classify_icms_diagnostic_reason(normalized_operation, source_register, cst_value, total_operation_value, base_icms, icms_value)
        base_gap = max(Decimal("0"), total_operation_value - base_icms)
        base_ratio = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )

        totals["total_operation_value"] += total_operation_value
        totals["base_icms"] += base_icms
        totals["icms_value"] += icms_value
        totals["base_gap"] += base_gap

        detail_key = (period, reason, source_register, cst_value, cfop_value, icms_rate)
        detail_bucket = detail_groups.setdefault(
            detail_key,
            {
                "operation_type": normalized_operation,
                "period": period,
                "reason": reason,
                "source_register": source_register,
                "cst_icms": cst_value,
                "cfop": cfop_value,
                "icms_rate": icms_rate,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_gap": Decimal("0"),
                "row_count": 0,
                "grouped_rows": [],
            },
        )
        detail_bucket["total_operation_value"] += total_operation_value
        detail_bucket["base_icms"] += base_icms
        detail_bucket["icms_value"] += icms_value
        detail_bucket["base_gap"] += base_gap
        detail_bucket["row_count"] += 1
        grouped_rows = detail_bucket.get("grouped_rows")
        if isinstance(grouped_rows, list):
            grouped_rows.append(dict(row))

        summary_bucket = summary_groups.setdefault(
            (period, reason),
            {
                "period": period,
                "reason": reason,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_gap": Decimal("0"),
                "row_count": 0,
            },
        )
        summary_bucket["total_operation_value"] += total_operation_value
        summary_bucket["base_icms"] += base_icms
        summary_bucket["icms_value"] += icms_value
        summary_bucket["base_gap"] += base_gap
        summary_bucket["row_count"] += 1

    detail_rows = sorted(
        detail_groups.values(),
        key=lambda row: (
            Decimal(row["base_gap"]),
            Decimal(row["total_operation_value"]),
        ),
        reverse=True,
    )
    for row in detail_rows:
        total_operation_value = Decimal(row["total_operation_value"])
        base_icms = Decimal(row["base_icms"])
        row["base_ratio"] = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        row["product_rows"] = (
            build_credit_diagnostic_product_rows(
                detailed_sales,
                normalized_operation,
                str(row["period"]),
                str(row["source_register"]),
                str(row["reason"]),
                str(row["cst_icms"]),
                str(row["cfop"]),
                Decimal(row["icms_rate"]),
            )
            if detailed_sales
            else []
        )
        row["launch_details"] = [
            dict(detail)
            for product_row in row["product_rows"]
            for detail in product_row.get("launch_details", [])
        ]
        row["document_count"] = len(
            {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in row["launch_details"]
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
        )

    summary_rows = sorted(
        summary_groups.values(),
        key=lambda row: Decimal(row["base_gap"]),
        reverse=True,
    )
    for row in summary_rows:
        total_operation_value = Decimal(row["total_operation_value"])
        base_icms = Decimal(row["base_icms"])
        row["base_ratio"] = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )

    return summary_rows, detail_rows, totals


def build_credit_diagnostic_product_rows(
    detailed_sales: list[dict[str, object]],
    operation_type: str,
    period: str,
    source_register: str,
    reason: str,
    cst_icms: str,
    cfop: str,
    icms_rate: Decimal,
) -> list[dict[str, object]]:
    normalized_operation = normalize_operation_type(operation_type)
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)
    normalized_cfop = str(cfop or "").strip()
    rate_key = Decimal(icms_rate).quantize(Decimal("0.01"))
    candidate_items: list[tuple[dict[str, object], str, Decimal, Decimal, Decimal]] = []
    fallback_items: list[tuple[dict[str, object], str, Decimal, Decimal, Decimal]] = []

    for source_item in detailed_sales:
        if normalize_operation_type(str(source_item.get("operation_type", ""))) != normalized_operation:
            continue
        if str(source_item.get("period", "")).strip() != str(period or "").strip():
            continue
        item = apply_sped_icms_consistency_rules(source_item)
        item_cst = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
        item_cfop = str(item.get("cfop", "")).strip()
        item_rate = (
            item.get("icms_rate", Decimal("0"))
            if isinstance(item.get("icms_rate", Decimal("0")), Decimal)
            else Decimal("0")
        ).quantize(Decimal("0.01"))
        if item_cst != normalized_cst or item_cfop != normalized_cfop or item_rate != rate_key:
            continue

        sale_value = item.get("sale_value", Decimal("0"))
        base_icms = item.get("base_icms", Decimal("0"))
        icms_value = item.get("icms_value", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_icms = base_icms if isinstance(base_icms, Decimal) else Decimal("0")
        icms_value = icms_value if isinstance(icms_value, Decimal) else Decimal("0")
        fallback_items.append((item, item_cst, sale_value, base_icms, icms_value))

        item_reason = classify_icms_diagnostic_reason(normalized_operation, source_register, item_cst, sale_value, base_icms, icms_value)
        if item_reason == reason:
            candidate_items.append((item, item_cst, sale_value, base_icms, icms_value))

    selected_items = candidate_items if candidate_items else fallback_items
    grouped_products: dict[tuple[str, str, str], dict[str, object]] = {}
    for item, _item_cst, sale_value, base_icms, icms_value in selected_items:
        item_rate = (
            item.get("icms_rate", Decimal("0"))
            if isinstance(item.get("icms_rate", Decimal("0")), Decimal)
            else Decimal("0")
        ).quantize(Decimal("0.01"))
        product_key = (
            str(item.get("code", "")).strip(),
            str(item.get("description", "")).strip(),
            str(item.get("ncm", "")).strip(),
        )
        bucket = grouped_products.setdefault(
            product_key,
            {
                "code": product_key[0],
                "description": product_key[1],
                "ncm": product_key[2],
                "period": period,
                "icms_rate": item_rate,
                "row_count": 0,
                "document_keys": set(),
                "launch_details": [],
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )
        bucket["row_count"] += 1
        document_identity = normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
        if document_identity:
            bucket["document_keys"].add(document_identity)
        bucket["launch_details"].append(dict(item))
        bucket["sale_value"] += sale_value
        bucket["base_icms"] += base_icms
        bucket["icms_value"] += icms_value

    return sorted(
        grouped_products.values(),
        key=lambda row: (Decimal(row["sale_value"]), str(row["code"]), str(row["description"])),
        reverse=True,
    )


def build_credit_diagnostic_period_comparison_dataset(
    summary_register_rows: list[dict[str, object]],
    detailed_sales: list[dict[str, object]],
    cst_filter_values: set[str],
    cfop_filter_values: set[str],
    operation_type: str = "Entrada",
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]], list[dict[str, object]]]:
    summary_rows, detail_rows, _totals = build_credit_diagnostic_datasets(
        summary_register_rows,
        operation_type,
        cst_filter_values,
        cfop_filter_values,
        detailed_sales,
    )
    periods = sorted(
        {
            str(row.get("period", "")).strip()
            for row in summary_rows
            if str(row.get("period", "")).strip()
        },
        key=period_label_sort_key,
    )
    grouped_reasons: dict[str, dict[str, object]] = {}
    for row in summary_rows:
        reason = str(row.get("reason", "")).strip()
        if not reason:
            continue
        bucket = grouped_reasons.setdefault(
            reason,
            {
                "reason": reason,
                "periods": {},
                "total_rows": 0,
                "total_operation_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_base_gap": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        period = str(row.get("period", "")).strip()
        bucket["periods"][period] = {
            "row_count": int(row.get("row_count", 0) or 0),
            "operation_value": Decimal(row.get("total_operation_value", Decimal("0"))),
            "base_icms": Decimal(row.get("base_icms", Decimal("0"))),
            "base_gap": Decimal(row.get("base_gap", Decimal("0"))),
            "base_ratio": Decimal(row.get("base_ratio", Decimal("0"))),
            "icms_value": Decimal(row.get("icms_value", Decimal("0"))),
        }
        bucket["total_rows"] += int(row.get("row_count", 0) or 0)
        bucket["total_operation_value"] += Decimal(row.get("total_operation_value", Decimal("0")))
        bucket["total_base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        bucket["total_base_gap"] += Decimal(row.get("base_gap", Decimal("0")))
        bucket["total_icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

    export_headers = ["Motivo"]
    for period in periods:
        export_headers.extend(
            [
                f"{period} Linhas",
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Perda de Base",
                f"{period} % Base/Oper",
                f"{period} Valor ICMS",
            ]
        )
    export_headers.extend(
        [
            "Total Linhas",
            "Total Valor Operacao",
            "Total Base ICMS",
            "Total Perda de Base",
            "Total % Base/Oper",
            "Total Valor ICMS",
        ]
    )

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for reason in sorted(grouped_reasons.keys(), key=normalize_text):
        reason_bucket = grouped_reasons[reason]
        values: list[object] = [reason]
        for period in periods:
            period_data = reason_bucket["periods"].get(period)
            if period_data is None:
                values.extend([0, Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    int(period_data["row_count"]),
                    Decimal(period_data["operation_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["base_gap"]),
                    Decimal(period_data["base_ratio"]),
                    Decimal(period_data["icms_value"]),
                ]
            )
        total_operation_value = Decimal(reason_bucket["total_operation_value"])
        total_base_icms = Decimal(reason_bucket["total_base_icms"])
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        values.extend(
            [
                int(reason_bucket["total_rows"]),
                total_operation_value,
                total_base_icms,
                Decimal(reason_bucket["total_base_gap"]),
                total_base_ratio,
                Decimal(reason_bucket["total_icms_value"]),
            ]
        )
        display_rows.append({"reason": reason, "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows, detail_rows


def merge_credit_diagnostic_detail_rows(detail_rows: list[dict[str, object]], reason: str, period: str) -> dict[str, object] | None:
    matching_rows = [
        row
        for row in detail_rows
        if str(row.get("reason", "")).strip() == str(reason).strip()
        and str(row.get("period", "")).strip() == str(period).strip()
    ]
    if not matching_rows:
        return None
    if len(matching_rows) == 1:
        return dict(matching_rows[0])

    source_registers = sorted({str(row.get("source_register", "")).strip() for row in matching_rows if str(row.get("source_register", "")).strip()})
    csts = sorted({str(row.get("cst_icms", "")).strip() for row in matching_rows if str(row.get("cst_icms", "")).strip()})
    cfops = sorted({str(row.get("cfop", "")).strip() for row in matching_rows if str(row.get("cfop", "")).strip()})
    grouped_rows: list[dict[str, object]] = []
    launch_details: list[dict[str, object]] = []
    product_buckets: dict[tuple[str, str, str, str, Decimal], dict[str, object]] = {}

    for row in matching_rows:
        grouped_rows.extend(list(row.get("grouped_rows", [])))
        launch_details.extend(list(row.get("launch_details", [])))
        for product in row.get("product_rows", []):
            product_key = (
                str(product.get("period", "")).strip(),
                str(product.get("code", "")).strip(),
                str(product.get("description", "")).strip(),
                str(product.get("ncm", "")).strip(),
                Decimal(product.get("icms_rate", Decimal("0"))),
            )
            bucket = product_buckets.setdefault(
                product_key,
                {
                    "period": product_key[0],
                    "code": product_key[1],
                    "description": product_key[2],
                    "ncm": product_key[3],
                    "icms_rate": product_key[4],
                    "row_count": 0,
                    "document_keys": set(),
                    "launch_details": [],
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            bucket["row_count"] += int(product.get("row_count", 0) or 0)
            bucket["document_keys"].update(set(product.get("document_keys", set())))
            bucket["launch_details"].extend(list(product.get("launch_details", [])))
            bucket["sale_value"] += Decimal(product.get("sale_value", Decimal("0")))
            bucket["base_icms"] += Decimal(product.get("base_icms", Decimal("0")))
            bucket["icms_value"] += Decimal(product.get("icms_value", Decimal("0")))

    total_operation_value = sum((Decimal(row.get("total_operation_value", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_icms = sum((Decimal(row.get("base_icms", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_gap = sum((Decimal(row.get("base_gap", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_icms_value = sum((Decimal(row.get("icms_value", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_ratio = (
        (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
        if total_operation_value > 0
        else Decimal("0.00")
    )
    return {
        "operation_type": str(matching_rows[0].get("operation_type", "")).strip(),
        "period": period,
        "reason": reason,
        "source_register": " | ".join(source_registers) if source_registers else "",
        "cst_icms": " | ".join(csts) if csts else "",
        "cfop": " | ".join(cfops) if cfops else "",
        "icms_rate": next(
            (
                Decimal(row.get("icms_rate", Decimal("0")))
                for row in matching_rows
                if Decimal(row.get("icms_rate", Decimal("0"))) > 0
            ),
            Decimal("0"),
        ),
        "row_count": sum((int(row.get("row_count", 0) or 0) for row in matching_rows), 0),
        "total_operation_value": total_operation_value,
        "base_icms": total_base_icms,
        "base_gap": total_base_gap,
        "base_ratio": total_base_ratio,
        "icms_value": total_icms_value,
        "grouped_rows": grouped_rows,
        "launch_details": launch_details,
        "document_count": len(
            {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in launch_details
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
        ),
        "product_rows": sorted(
            product_buckets.values(),
            key=lambda row: (Decimal(row.get("sale_value", Decimal("0"))), str(row.get("code", ""))),
            reverse=True,
        ),
    }


def build_credit_diagnostic_period_detail_comparison_dataset(
    detail_rows: list[dict[str, object]],
    periods: list[str],
) -> tuple[list[str], list[dict[str, object]], list[list[object]]]:
    grouped_details: dict[tuple[str, str, str, str, Decimal], dict[str, object]] = {}

    for row in detail_rows:
        reason = str(row.get("reason", "")).strip()
        source_register = str(row.get("source_register", "")).strip()
        cst_icms = str(row.get("cst_icms", "")).strip()
        cfop = str(row.get("cfop", "")).strip()
        icms_rate = Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
        period = str(row.get("period", "")).strip()
        if not reason:
            continue
        detail_key = (reason, source_register, cst_icms, cfop, icms_rate)
        bucket = grouped_details.setdefault(
            detail_key,
            {
                "reason": reason,
                "source_register": source_register,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": icms_rate,
                "periods": {},
                "total_rows": 0,
                "total_operation_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_base_gap": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
        base_icms = Decimal(row.get("base_icms", Decimal("0")))
        icms_value = Decimal(row.get("icms_value", Decimal("0")))
        effective_rate = compute_display_icms_rate(
            icms_rate,
            operation_value,
            base_icms,
            icms_value,
        )
        bucket["periods"][period] = {
            "row_count": int(row.get("row_count", 0) or 0),
            "operation_value": operation_value,
            "base_icms": base_icms,
            "base_gap": Decimal(row.get("base_gap", Decimal("0"))),
            "base_ratio": Decimal(row.get("base_ratio", Decimal("0"))),
            "icms_value": icms_value,
            "effective_rate": effective_rate,
        }
        bucket["total_rows"] += int(row.get("row_count", 0) or 0)
        bucket["total_operation_value"] += operation_value
        bucket["total_base_icms"] += base_icms
        bucket["total_base_gap"] += Decimal(row.get("base_gap", Decimal("0")))
        bucket["total_icms_value"] += icms_value

    export_headers = ["Motivo", "Registro", "CST", "CFOP", "Aliq"]
    for period in periods:
        export_headers.extend(
            [
                f"{period} Linhas",
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Perda de Base",
                f"{period} % Base/Oper",
                f"{period} Valor ICMS",
                f"{period} % Aliq Efetiva",
            ]
        )
    export_headers.extend(
        [
            "Total Linhas",
            "Total Valor Operacao",
            "Total Base ICMS",
            "Total Perda de Base",
            "Total % Base/Oper",
            "Total Valor ICMS",
            "Total % Aliq Efetiva",
        ]
    )

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for detail_key in sorted(
        grouped_details.keys(),
        key=lambda item: (
            normalize_text(item[0]),
            normalize_text(item[1]),
            normalize_text(item[2]),
            normalize_text(item[3]),
            item[4],
        ),
    ):
        bucket = grouped_details[detail_key]
        values: list[object] = [
            bucket["reason"],
            bucket["source_register"],
            bucket["cst_icms"],
            bucket["cfop"],
            Decimal(bucket["icms_rate"]),
        ]
        for period in periods:
            period_data = bucket["periods"].get(period)
            if period_data is None:
                values.extend([0, Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    int(period_data["row_count"]),
                    Decimal(period_data["operation_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["base_gap"]),
                    Decimal(period_data["base_ratio"]),
                    Decimal(period_data["icms_value"]),
                    Decimal(period_data["effective_rate"]),
                ]
            )
        total_operation_value = Decimal(bucket["total_operation_value"])
        total_base_icms = Decimal(bucket["total_base_icms"])
        total_icms_value = Decimal(bucket["total_icms_value"])
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        total_effective_rate = compute_display_icms_rate(
            Decimal(bucket["icms_rate"]),
            total_operation_value,
            total_base_icms,
            total_icms_value,
        )
        values.extend(
            [
                int(bucket["total_rows"]),
                total_operation_value,
                total_base_icms,
                Decimal(bucket["total_base_gap"]),
                total_base_ratio,
                total_icms_value,
                total_effective_rate,
            ]
        )
        display_rows.append(
            {
                "reason": bucket["reason"],
                "source_register": bucket["source_register"],
                "cst_icms": bucket["cst_icms"],
                "cfop": bucket["cfop"],
                "icms_rate": Decimal(bucket["icms_rate"]),
                "values": values,
            }
        )
        export_rows.append(values)

    return export_headers, display_rows, export_rows


def build_multi_sped_entry_analysis(
    sped_paths: list[Path],
) -> tuple[
    list[str],
    list[list[object]],
    list[str],
    list[list[object]],
    list[str],
    list[list[object]],
    list[str],
    list[list[object]],
]:
    if not sped_paths:
        raise ValueError("Selecione ao menos um arquivo SPED para gerar a analise.")
    if len(sped_paths) > 12:
        raise ValueError("A analise aceita no maximo 12 arquivos SPED por vez.")

    period_labels: list[str] = []
    period_entries: list[tuple[str, dict[str, dict[str, object]]]] = []
    product_catalog: dict[str, dict[str, str]] = {}
    used_labels: set[str] = set()

    for sped_path in sped_paths:
        if not sped_path.exists():
            raise ValueError(f"O arquivo SPED nao foi encontrado: {sped_path}")

        _, _, detailed_sales, _, _ = read_sped_file(sped_path)
        entry_items = [
            item for item in detailed_sales
            if str(item.get("operation_type", "")).strip() == "Entrada"
        ]
        label_base = infer_sped_period_label(sped_path, entry_items)
        label = label_base
        suffix = 2
        while label in used_labels:
            label = f"{label_base} ({suffix})"
            suffix += 1
        used_labels.add(label)
        period_labels.append(label)

        grouped_entries: dict[str, list[dict[str, object]]] = defaultdict(list)
        for item in entry_items:
            code = str(item.get("code", "")).strip()
            if not code:
                continue
            description = str(item.get("description", "")).strip()
            ncm = str(item.get("ncm", "")).strip()
            product_catalog.setdefault(
                code,
                {
                    "description": description,
                    "ncm": ncm,
                },
            )
            if description and not product_catalog[code]["description"]:
                product_catalog[code]["description"] = description
            if ncm and not product_catalog[code]["ncm"]:
                product_catalog[code]["ncm"] = ncm
            grouped_entries[code].append(item)

        summarized_entries = {
            code: summarize_entry_analysis(items)
            for code, items in grouped_entries.items()
        }
        period_entries.append((label, summarized_entries))

    headers = ["Codigo Produto", "Descricao", "NCM"]
    for label in period_labels:
        headers.extend(
            [
                f"CFOP Entrada {label}",
                f"CST Entrada {label}",
                f"Credito ICMS {label}",
                f"Status {label}",
            ]
        )
    headers.append("Observacao Geral")

    rows: list[list[object]] = []
    divergence_rows: list[list[object]] = []
    summary_headers = [
        "Codigo Produto",
        "Descricao",
        "NCM",
        "Meses com Entrada",
        "Meses sem Entrada",
        "CFOP Predominante",
        "CST Predominante",
        "Qtde CFOPs Distintos",
        "Qtde CSTs Distintos",
        "Meses sem Credito",
        "Meses com Divergencia",
        "Resumo Divergencias",
    ]
    summary_rows: list[list[object]] = []

    period_status_counts: dict[str, dict[str, int]] = {
        label: {
            "com_entrada": 0,
            "ok": 0,
            "sem_credito": 0,
            "divergencia": 0,
            "sem_entrada": 0,
        }
        for label in period_labels
    }

    for code in sorted(product_catalog):
        product = product_catalog[code]
        row: list[object] = [code, product["description"], product["ncm"]]
        cfops_seen: set[str] = set()
        csts_seen: set[str] = set()
        cfop_frequency: dict[str, int] = defaultdict(int)
        cst_frequency: dict[str, int] = defaultdict(int)
        period_notes: list[str] = []
        row_has_issue = False
        months_with_entry = 0
        months_without_entry = 0
        months_without_credit = 0
        divergent_months = 0

        for label, summarized_entries in period_entries:
            summary = summarized_entries.get(code)
            if summary is None:
                row.extend(["", "", Decimal("0"), "Sem entrada"])
                period_notes.append(f"{label}: sem entrada")
                months_without_entry += 1
                period_status_counts[label]["sem_entrada"] += 1
                continue

            cfop_value = str(summary["cfop"]).strip()
            cst_value = str(summary["cst"]).strip()
            status_value = str(summary["status"]).strip()
            credit_value = Decimal(summary["credit"])
            row.extend([cfop_value, cst_value, credit_value, status_value])
            months_with_entry += 1
            period_status_counts[label]["com_entrada"] += 1

            if cfop_value:
                cfops_seen.add(cfop_value)
                cfop_frequency[cfop_value] += 1
            if cst_value:
                csts_seen.add(cst_value)
                cst_frequency[cst_value] += 1
            if "Sem credito" in status_value:
                months_without_credit += 1
                period_status_counts[label]["sem_credito"] += 1
            if status_value == "Ok":
                period_status_counts[label]["ok"] += 1
            else:
                row_has_issue = True
                divergent_months += 1
                if "Sem entrada" not in status_value and "Sem credito" not in status_value:
                    period_status_counts[label]["divergencia"] += 1
                period_notes.append(f"{label}: {status_value}")

        overall_notes: list[str] = []
        if len(cfops_seen) > 1:
            overall_notes.append("CFOP variou entre os periodos")
            row_has_issue = True
        if len(csts_seen) > 1:
            overall_notes.append("CST variou entre os periodos")
            row_has_issue = True
        overall_notes.extend(period_notes)
        row.append(" | ".join(overall_notes))
        rows.append(row)
        if row_has_issue:
            divergence_rows.append(row.copy())

        predominant_cfop = ""
        predominant_cst = ""
        if cfop_frequency:
            predominant_cfop = sorted(cfop_frequency.items(), key=lambda item: (-item[1], item[0]))[0][0]
        if cst_frequency:
            predominant_cst = sorted(cst_frequency.items(), key=lambda item: (-item[1], item[0]))[0][0]
        summary_rows.append(
            [
                code,
                product["description"],
                product["ncm"],
                months_with_entry,
                months_without_entry,
                predominant_cfop,
                predominant_cst,
                len(cfops_seen),
                len(csts_seen),
                months_without_credit,
                divergent_months,
                " | ".join(overall_notes),
            ]
        )

    monthly_summary_headers = [
        "Periodo",
        "Produtos com Entrada",
        "Produtos sem Entrada",
        "Produtos Ok",
        "Produtos sem Credito",
        "Produtos com Divergencia",
    ]
    monthly_summary_rows = [
        [
            label,
            counts["com_entrada"],
            counts["sem_entrada"],
            counts["ok"],
            counts["sem_credito"],
            counts["divergencia"],
        ]
        for label, counts in period_status_counts.items()
    ]

    return (
        summary_headers,
        summary_rows,
        headers,
        rows,
        headers,
        divergence_rows,
        monthly_summary_headers,
        monthly_summary_rows,
    )


def build_entry_period_comparison_rows(
    sped_paths: list[Path],
    xml_sources: list[Path] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[str], list[dict[str, object]]]:
    if not sped_paths:
        raise ValueError("Selecione ao menos um arquivo SPED para processar.")
    if len(sped_paths) > 12:
        raise ValueError("A consulta aceita no maximo 12 arquivos SPED por vez.")

    period_labels: list[str] = []
    comparison_rows: list[dict[str, object]] = []
    used_labels: set[str] = set()
    prepared_periods: list[dict[str, object]] = []
    all_xml_document_keys: set[str] = set()

    total_steps = (len(sped_paths) * 2) + (1 if xml_sources else 0)
    current_step = 0
    for sped_path in sped_paths:
        if not sped_path.exists():
            raise ValueError(f"O arquivo SPED nao foi encontrado: {sped_path}")
        if progress_callback:
            progress_callback(current_step, total_steps, f"Lendo entradas: {sped_path.name}")
        current_step += 1

        _, _, detailed_sales, _, _ = read_sped_file(sped_path)
        entry_items = [
            item for item in detailed_sales
            if str(item.get("operation_type", "")).strip() == "Entrada"
        ]
        sped_documents = scan_sped_c100_documents(sped_path, "Entrada")
        xml_document_keys = {
            document_key
            for (operation_type, document_key), meta in sped_documents.items()
            if operation_type == "Entrada" and str(meta.get("cod_mod", "")).strip() in {"55", "65"}
        }
        c100_c190_fallback_rows = [
            *read_c100_c190_fallback_rows(sped_path, "Entrada", "55"),
            *read_c100_c190_fallback_rows(sped_path, "Entrada", "65"),
        ]
        fallback_rows_by_document_key: dict[str, list[dict[str, object]]] = defaultdict(list)
        for fallback_row in c100_c190_fallback_rows:
            fallback_document_key = normalize_document_key(str(fallback_row.get("document_key", "")))
            if fallback_document_key:
                fallback_rows_by_document_key[fallback_document_key].append(fallback_row)
        all_xml_document_keys.update(xml_document_keys)
        prepared_periods.append(
            {
                "sped_path": sped_path,
                "entry_items": entry_items,
                "sped_documents": sped_documents,
                "fallback_rows_by_document_key": fallback_rows_by_document_key,
            }
        )

    if progress_callback and xml_sources:
        progress_callback(current_step, total_steps, "Indexando XMLs modelos 55/65...")
    global_xml_index = (
        build_xml_fiscal_item_index(xml_sources, all_xml_document_keys)
        if xml_sources and all_xml_document_keys
        else {}
    )
    current_step += 1 if xml_sources else 0

    for prepared_period in prepared_periods:
        sped_path = prepared_period["sped_path"]
        entry_items = list(prepared_period["entry_items"])
        sped_documents = dict(prepared_period["sped_documents"])
        fallback_rows_by_document_key = dict(prepared_period["fallback_rows_by_document_key"])
        if progress_callback:
            progress_callback(current_step, total_steps, f"Montando comparativo: {sped_path.name}")
        current_step += 1

        xml_document_keys = {
            document_key
            for (operation_type, document_key), meta in sped_documents.items()
            if operation_type == "Entrada" and str(meta.get("cod_mod", "")).strip() in {"55", "65"}
        }
        xml_index = {
            document_key: global_xml_index.get(document_key, [])
            for document_key in xml_document_keys
            if global_xml_index.get(document_key)
        }
        xml_index_by_identity: dict[tuple[str, str, str], list[dict[str, object]]] = {}
        for items in xml_index.values():
            if not items:
                continue
            first_item = items[0]
            identity_key = (
                str(first_item.get("modelo", "")).strip(),
                str(first_item.get("numero_nfce", "")).strip(),
                str(first_item.get("serie", "")).strip(),
            )
            if identity_key[1]:
                xml_index_by_identity[identity_key] = items
        xml_note_snapshots = {
            document_key: build_nfce_note_snapshot(
                document_key,
                items,
                str(sped_documents.get(("Entrada", document_key), {}).get("cod_mod", "")).strip() or "55",
            )
            for document_key, items in xml_index.items()
            if items
        }

        documents_by_identity: dict[tuple[str, str], dict[str, object]] = {}
        for item in entry_items:
            document_key = normalize_document_key(str(item.get("document_key", "")))
            document_number = str(item.get("document_number", "")).strip()
            identity = (document_key or document_number, str(item.get("document_date", "")).strip())
            if not identity[0]:
                continue
            bucket = documents_by_identity.setdefault(
                identity,
                {
                    "document_number": document_number,
                    "document_key": document_key,
                    "document_date": str(item.get("document_date", "")).strip(),
                    "document_series": str(item.get("document_series", "")).strip(),
                    "document_model": str(item.get("document_model", "")).strip(),
                    "participant_code": str(item.get("participant_code", "")).strip(),
                    "participant_name": str(item.get("participant_name", "")).strip(),
                    "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                    "items": [],
                },
            )
            bucket["items"].append(dict(item))

        entry_document_keys_with_items = {
            normalize_document_key(str(item.get("document_key", "")))
            for item in entry_items
            if normalize_document_key(str(item.get("document_key", "")))
        }

        xml_rebuilt_entries: list[dict[str, object]] = []
        c190_fallback_entries: list[dict[str, object]] = []
        for (operation_type, document_key), document_meta in sped_documents.items():
            if operation_type != "Entrada":
                continue
            document_model = str(document_meta.get("cod_mod", "")).strip()
            if document_model not in {"55", "65"}:
                continue
            if document_key in entry_document_keys_with_items:
                continue
            xml_items = xml_index.get(document_key, [])
            if not xml_items:
                xml_items = xml_index_by_identity.get(
                    (
                        document_model,
                        str(document_meta.get("document_number", "")).strip(),
                        str(document_meta.get("document_series", "")).strip(),
                    ),
                    [],
                )
            if xml_items:
                note_snapshot = xml_note_snapshots.get(document_key)
                if note_snapshot is None:
                    note_snapshot = build_nfce_note_snapshot(document_key, xml_items, document_model or "55")
                document_rebuilt_items: list[dict[str, object]] = []
                for xml_item in xml_items:
                    rebuilt_item = {
                        "operation_type": "Entrada",
                        "document_number": str(document_meta.get("document_number", "")).strip() or str(xml_item.get("numero_nfce", "")).strip(),
                        "document_key": document_key,
                        "document_date": str(document_meta.get("document_date", "")).strip() or str(xml_item.get("data_emissao", "")).strip(),
                        "document_series": str(document_meta.get("document_series", "")).strip() or str(xml_item.get("serie", "")).strip(),
                        "document_model": str(document_meta.get("cod_mod", "")).strip() or str(xml_item.get("modelo", "")).strip() or "55",
                        "participant_code": str(document_meta.get("participant_code", "")).strip(),
                        "participant_name": str(xml_item.get("emitente", "")).strip() or f"XML modelo {str(document_meta.get('cod_mod', '')).strip() or '55'}",
                        "participant_tax_id": str(xml_item.get("cnpj_emitente", "")).strip(),
                        "item_number": str(xml_item.get("item", "")).strip(),
                        "code": str(xml_item.get("codigo", "")).strip() or str(xml_item.get("descricao", "")).strip(),
                        "description": str(xml_item.get("descricao", "")).strip(),
                        "ncm": str(xml_item.get("ncm", "")).strip(),
                        "cest": str(xml_item.get("cest", "")).strip(),
                        "cst_icms": compose_xml_icms_cst_for_sped(xml_item),
                        "cfop": str(xml_item.get("cfop", "")).strip(),
                        "quantity": Decimal(xml_item.get("quantidade", Decimal("0"))),
                        "icms_rate": Decimal(xml_item.get("aliquota_icms", Decimal("0"))),
                        "sale_value": Decimal(xml_item.get("valor_operacao", xml_item.get("valor_produto", Decimal("0")))),
                        "base_icms": Decimal(xml_item.get("base_icms", Decimal("0"))),
                        "icms_value": Decimal(xml_item.get("valor_icms", Decimal("0"))),
                        "base_icms_st": Decimal(xml_item.get("base_icms_st", Decimal("0"))),
                        "icms_st_rate": Decimal(xml_item.get("aliquota_icms_st", Decimal("0"))),
                        "icms_st_value": Decimal(xml_item.get("valor_icms_st", Decimal("0"))),
                        "ipi_value": Decimal(xml_item.get("valor_ipi", Decimal("0"))),
                        "note_snapshot": note_snapshot,
                    }
                    document_rebuilt_items.append(rebuilt_item)
                    identity = (
                        document_key or str(rebuilt_item["document_number"]).strip(),
                        str(rebuilt_item["document_date"]).strip(),
                    )
                    bucket = documents_by_identity.setdefault(
                        identity,
                        {
                            "document_number": rebuilt_item["document_number"],
                            "document_key": document_key,
                            "document_date": rebuilt_item["document_date"],
                            "document_series": rebuilt_item["document_series"],
                            "document_model": rebuilt_item["document_model"],
                            "participant_code": rebuilt_item["participant_code"],
                            "participant_name": rebuilt_item["participant_name"],
                            "participant_tax_id": rebuilt_item["participant_tax_id"],
                            "items": [],
                        },
                    )
                    bucket["items"].append(rebuilt_item)
                document_rebuilt_items = normalize_xml_rebuilt_items_with_fallback(
                    document_rebuilt_items,
                    fallback_rows_by_document_key.get(document_key, []),
                )
                xml_rebuilt_entries.extend(document_rebuilt_items)
                continue

            fallback_rows = fallback_rows_by_document_key.get(document_key, [])
            for fallback_row in fallback_rows:
                rebuilt_fallback_row = dict(fallback_row)
                c190_fallback_entries.append(rebuilt_fallback_row)
                identity = (
                    document_key or str(rebuilt_fallback_row["document_number"]).strip(),
                    str(rebuilt_fallback_row["document_date"]).strip(),
                )
                bucket = documents_by_identity.setdefault(
                    identity,
                    {
                        "document_number": rebuilt_fallback_row["document_number"],
                        "document_key": document_key,
                        "document_date": rebuilt_fallback_row["document_date"],
                        "document_series": rebuilt_fallback_row["document_series"],
                        "document_model": rebuilt_fallback_row["document_model"],
                        "participant_code": rebuilt_fallback_row["participant_code"],
                        "participant_name": rebuilt_fallback_row["participant_name"],
                        "participant_tax_id": rebuilt_fallback_row["participant_tax_id"],
                        "items": [],
                    },
                )
                bucket["items"].append(dict(rebuilt_fallback_row))

        all_entry_items = list(entry_items)
        for item in all_entry_items:
            document_key = normalize_document_key(str(item.get("document_key", "")))
            if document_key in xml_note_snapshots:
                item["note_snapshot"] = xml_note_snapshots[document_key]
        all_entry_items.extend(xml_rebuilt_entries)
        all_entry_items.extend(c190_fallback_entries)

        label_base = infer_sped_period_label(sped_path, all_entry_items)
        label = label_base
        suffix = 2
        while label in used_labels:
            label = f"{label_base} ({suffix})"
            suffix += 1
        used_labels.add(label)
        period_labels.append(label)

        grouped_items: dict[str, list[dict[str, object]]] = defaultdict(list)
        for item in all_entry_items:
            code = str(item.get("code", "")).strip()
            if code:
                grouped_items[code].append(item)

        for code, items in grouped_items.items():
            cfops = sorted({str(item.get("cfop", "")).strip() for item in items if str(item.get("cfop", "")).strip()})
            csts = sorted({str(item.get("cst_icms", "")).strip() for item in items if str(item.get("cst_icms", "")).strip()})
            descriptions = sorted({str(item.get("description", "")).strip() for item in items if str(item.get("description", "")).strip()})
            ncms = sorted({str(item.get("ncm", "")).strip() for item in items if str(item.get("ncm", "")).strip()})
            cests = collect_cest_values(items)
            document_keys = {
                normalize_document_key(str(item.get("document_key", "")))
                for item in items
                if normalize_document_key(str(item.get("document_key", "")))
            }
            launch_details = [
                {
                    "operation_type": "Entrada",
                    "document_number": str(item.get("document_number", "")).strip(),
                    "document_key": str(item.get("document_key", "")).strip(),
                    "document_date": str(item.get("document_date", "")).strip(),
                    "item_number": str(item.get("item_number", "")).strip(),
                    "code": str(item.get("code", "")).strip(),
                    "ncm": str(item.get("ncm", "")).strip(),
                    "cest": str(item.get("cest", "")).strip(),
                    "cfop": str(item.get("cfop", "")).strip(),
                    "cst_icms": str(item.get("cst_icms", "")).strip(),
                    "icms_rate": Decimal(item.get("icms_rate", Decimal("0"))),
                    "quantity": Decimal(item.get("quantity", Decimal("0"))),
                    "sale_value": Decimal(item.get("sale_value", Decimal("0"))),
                    "base_icms": Decimal(item.get("base_icms", Decimal("0"))),
                    "icms_value": Decimal(item.get("icms_value", Decimal("0"))),
                    "base_icms_st": Decimal(item.get("base_icms_st", Decimal("0"))),
                    "icms_st_value": Decimal(item.get("icms_st_value", Decimal("0"))),
                    "ipi_value": Decimal(item.get("ipi_value", Decimal("0"))),
                    "description": str(item.get("description", "")).strip(),
                    "document_series": str(item.get("document_series", "")).strip(),
                    "document_model": str(item.get("document_model", "")).strip(),
                    "participant_code": str(item.get("participant_code", "")).strip(),
                    "participant_name": str(item.get("participant_name", "")).strip(),
                    "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                    "note_snapshot": item.get("note_snapshot") or documents_by_identity.get(
                        (
                            normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip(),
                            str(item.get("document_date", "")).strip(),
                        )
                    ),
                }
                for item in sorted(
                    items,
                    key=lambda current: (
                        str(current.get("document_date", "")),
                        str(current.get("document_number", "")),
                        str(current.get("item_number", "")),
                    ),
                )
            ]
            comparison_rows.append(
                {
                    "period": label,
                    "file_name": sped_path.name,
                    "curve_abc": "C",
                    "code": code,
                    "description": " | ".join(descriptions),
                    "ncm": " | ".join(ncms),
                    "cest": " | ".join(cests),
                    "cfop": " | ".join(cfops),
                    "cst_icms": " | ".join(csts),
                    "suppliers": " | ".join(
                        sorted(
                            {
                                str(item.get("participant_name", "")).strip()
                                for item in items
                                if str(item.get("participant_name", "")).strip()
                            }
                        )
                    ),
                    "supplier_count": len(
                        {
                            (
                                str(item.get("participant_name", "")).strip(),
                                str(item.get("participant_tax_id", "")).strip(),
                            )
                            for item in items
                            if str(item.get("participant_name", "")).strip() or str(item.get("participant_tax_id", "")).strip()
                        }
                    ),
                    "quantity": sum((Decimal(item.get("quantity", Decimal("0"))) for item in items), Decimal("0")),
                    "sale_value": sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                    "base_icms": sum((Decimal(item.get("base_icms", Decimal("0"))) for item in items), Decimal("0")),
                    "icms_value": sum((Decimal(item.get("icms_value", Decimal("0"))) for item in items), Decimal("0")),
                    "display_icms_rate": compute_display_icms_rate(
                        next((Decimal(detail.get("icms_rate", Decimal("0"))) for detail in launch_details if Decimal(detail.get("icms_rate", Decimal("0"))) > 0), Decimal("0")),
                        sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                        sum((Decimal(item.get("base_icms", Decimal("0"))) for item in items), Decimal("0")),
                        sum((Decimal(item.get("icms_value", Decimal("0"))) for item in items), Decimal("0")),
                    ),
                    "document_count": len(document_keys),
                    "launch_count": len(items),
                    "status": summarize_entry_analysis(items)["status"],
                    "has_icms_reduction": any(
                        has_icms_reduction(
                            detail.get("icms_rate", Decimal("0")),
                            detail.get("sale_value", Decimal("0")),
                            detail.get("base_icms", Decimal("0")),
                            detail.get("icms_value", Decimal("0")),
                        )
                        for detail in launch_details
                    ),
                    "launch_details": launch_details,
                }
            )

    curve_labels = calculate_abc_curve_labels(
        {
            str(row.get("code", "")).strip(): Decimal(row.get("sale_value", Decimal("0")))
            for row in comparison_rows
            if str(row.get("code", "")).strip()
        }
    )
    for row in comparison_rows:
        row["curve_abc"] = curve_labels.get(str(row.get("code", "")).strip(), "C")

    comparison_rows.sort(
        key=lambda item: (
            str(item["code"]),
            str(item["period"]),
            str(item["cfop"]),
            str(item["cst_icms"]),
        )
    )
    return period_labels, comparison_rows


def summarize_sale_analysis(sale_items: list[dict[str, object]]) -> dict[str, object]:
    cfops = sorted({str(item.get("cfop", "")).strip() for item in sale_items if str(item.get("cfop", "")).strip()})
    csts = sorted({str(item.get("cst_icms", "")).strip() for item in sale_items if str(item.get("cst_icms", "")).strip()})
    total_debit = sum((Decimal(item.get("icms_value", Decimal("0"))) for item in sale_items), Decimal("0"))

    status_parts: list[str] = []
    if not sale_items:
        status_parts.append("Sem saida")
    else:
        if not cfops:
            status_parts.append("Erro: CFOP vazio")
        elif len(cfops) > 1:
            status_parts.append("Erro: multiplos CFOPs")
        if not csts:
            status_parts.append("Erro: CST vazio")
        elif len(csts) > 1:
            status_parts.append("Erro: multiplos CSTs")
        if total_debit < Decimal("0"):
            status_parts.append("Erro: debito negativo")
        elif total_debit == Decimal("0"):
            status_parts.append("Sem debito")
        if not status_parts:
            status_parts.append("Ok")

    return {
        "cfop": " | ".join(cfops),
        "cst": " | ".join(csts),
        "debit": total_debit,
        "has_sale": bool(sale_items),
        "status": " | ".join(status_parts),
    }


def build_nfce_note_snapshot(
    document_key: str,
    xml_items: list[dict[str, object]],
    document_model: str = "65",
) -> dict[str, object]:
    first_item = xml_items[0] if xml_items else {}
    note_items = [
        {
            "item_number": str(item.get("item", "")).strip(),
            "code": str(item.get("codigo", "")).strip(),
            "description": str(item.get("descricao", "")).strip(),
            "ncm": str(item.get("ncm", "")).strip(),
            "cest": str(item.get("cest", "")).strip(),
            "cfop": str(item.get("cfop", "")).strip(),
            "cst_icms": compose_xml_icms_cst_for_sped(item),
            "icms_rate": Decimal(item.get("aliquota_icms", Decimal("0"))),
            "quantity": Decimal(item.get("quantidade", Decimal("0"))),
            "sale_value": Decimal(item.get("valor_operacao", item.get("valor_produto", Decimal("0")))),
            "base_icms": Decimal(item.get("base_icms", Decimal("0"))),
            "icms_value": Decimal(item.get("valor_icms", Decimal("0"))),
        }
        for item in xml_items
    ]
    return {
        "document_number": str(first_item.get("numero_nfce", "")).strip(),
        "document_key": document_key,
        "document_date": str(first_item.get("data_emissao", "")).strip(),
        "document_series": str(first_item.get("serie", "")).strip(),
        "document_model": document_model,
        "participant_code": "",
        "participant_name": str(first_item.get("emitente", "")).strip(),
        "participant_tax_id": str(first_item.get("cnpj_emitente", "")).strip(),
        "items": note_items,
    }


def build_sale_period_comparison_rows(
    sped_paths: list[Path],
    xml_sources: list[Path],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[str], list[dict[str, object]]]:
    if not sped_paths:
        raise ValueError("Selecione ao menos um arquivo SPED para processar.")
    if len(sped_paths) > 12:
        raise ValueError("A consulta aceita no maximo 12 arquivos SPED por vez.")

    period_labels: list[str] = []
    comparison_rows: list[dict[str, object]] = []
    used_labels: set[str] = set()
    prepared_periods: list[dict[str, object]] = []
    all_xml_document_keys: set[str] = set()

    total_steps = (len(sped_paths) * 2) + (1 if xml_sources else 0)
    current_step = 0
    for sped_path in sped_paths:
        if not sped_path.exists():
            raise ValueError(f"O arquivo SPED nao foi encontrado: {sped_path}")
        if progress_callback:
            progress_callback(current_step, total_steps, f"Lendo saidas: {sped_path.name}")
        current_step += 1

        _, _, detailed_sales, _, _ = read_sped_file(sped_path)
        sale_items = [
            item for item in detailed_sales
            if str(item.get("operation_type", "")).strip() == "Saida"
        ]
        sped_documents = scan_sped_c100_documents(sped_path, "Saida")
        xml_document_keys = {
            document_key
            for (operation_type, document_key), meta in sped_documents.items()
            if operation_type == "Saida" and str(meta.get("cod_mod", "")).strip() in {"55", "65"}
        }
        c100_c190_fallback_rows = [
            *read_c100_c190_fallback_rows(sped_path, "Saida", "55"),
            *read_c100_c190_fallback_rows(sped_path, "Saida", "65"),
        ]
        fallback_rows_by_document_key: dict[str, list[dict[str, object]]] = defaultdict(list)
        for fallback_row in c100_c190_fallback_rows:
            fallback_document_key = normalize_document_key(str(fallback_row.get("document_key", "")))
            if fallback_document_key:
                fallback_rows_by_document_key[fallback_document_key].append(fallback_row)
        all_xml_document_keys.update(xml_document_keys)
        prepared_periods.append(
            {
                "sped_path": sped_path,
                "sale_items": sale_items,
                "sped_documents": sped_documents,
                "fallback_rows_by_document_key": fallback_rows_by_document_key,
            }
        )

    if progress_callback and xml_sources:
        progress_callback(current_step, total_steps, "Indexando XMLs modelos 55/65...")
    global_xml_index = (
        build_xml_fiscal_item_index(xml_sources, all_xml_document_keys)
        if xml_sources and all_xml_document_keys
        else {}
    )
    current_step += 1 if xml_sources else 0

    for prepared_period in prepared_periods:
        sped_path = prepared_period["sped_path"]
        sale_items = list(prepared_period["sale_items"])
        sped_documents = dict(prepared_period["sped_documents"])
        fallback_rows_by_document_key = dict(prepared_period["fallback_rows_by_document_key"])
        if progress_callback:
            progress_callback(current_step, total_steps, f"Montando comparativo: {sped_path.name}")
        current_step += 1

        xml_document_keys = {
            document_key
            for (operation_type, document_key), meta in sped_documents.items()
            if operation_type == "Saida" and str(meta.get("cod_mod", "")).strip() in {"55", "65"}
        }
        xml_index = {
            document_key: global_xml_index.get(document_key, [])
            for document_key in xml_document_keys
            if global_xml_index.get(document_key)
        }
        xml_index_by_identity: dict[tuple[str, str, str], list[dict[str, object]]] = {}
        for items in xml_index.values():
            if not items:
                continue
            first_item = items[0]
            identity_key = (
                str(first_item.get("modelo", "")).strip(),
                str(first_item.get("numero_nfce", "")).strip(),
                str(first_item.get("serie", "")).strip(),
            )
            if identity_key[1]:
                xml_index_by_identity[identity_key] = items
        xml_note_snapshots = {
            document_key: build_nfce_note_snapshot(
                document_key,
                items,
                str(sped_documents.get(("Saida", document_key), {}).get("cod_mod", "")).strip() or "65",
            )
            for document_key, items in xml_index.items()
            if items
        }

        documents_by_identity: dict[tuple[str, str], dict[str, object]] = {}
        for item in sale_items:
            document_key = normalize_document_key(str(item.get("document_key", "")))
            document_number = str(item.get("document_number", "")).strip()
            identity = (document_key or document_number, str(item.get("document_date", "")).strip())
            if not identity[0]:
                continue
            bucket = documents_by_identity.setdefault(
                identity,
                {
                    "document_number": document_number,
                    "document_key": document_key,
                    "document_date": str(item.get("document_date", "")).strip(),
                    "document_series": str(item.get("document_series", "")).strip(),
                    "document_model": str(item.get("document_model", "")).strip(),
                    "participant_code": str(item.get("participant_code", "")).strip(),
                    "participant_name": str(item.get("participant_name", "")).strip(),
                    "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                    "items": [],
                },
            )
            bucket["items"].append(dict(item))

        sale_document_keys_with_items = {
            normalize_document_key(str(item.get("document_key", "")))
            for item in sale_items
            if normalize_document_key(str(item.get("document_key", "")))
        }

        xml_rebuilt_sales: list[dict[str, object]] = []
        c190_fallback_sales: list[dict[str, object]] = []
        for (operation_type, document_key), document_meta in sped_documents.items():
            if operation_type != "Saida":
                continue
            document_model = str(document_meta.get("cod_mod", "")).strip()
            if document_model not in {"55", "65"}:
                continue
            if document_key in sale_document_keys_with_items:
                continue
            xml_items = xml_index.get(document_key, [])
            if not xml_items:
                xml_items = xml_index_by_identity.get(
                    (
                        document_model,
                        str(document_meta.get("document_number", "")).strip(),
                        str(document_meta.get("document_series", "")).strip(),
                    ),
                    [],
                )
            if xml_items:
                note_snapshot = xml_note_snapshots.get(document_key)
                if note_snapshot is None:
                    note_snapshot = build_nfce_note_snapshot(document_key, xml_items, document_model or "65")
                document_rebuilt_items: list[dict[str, object]] = []
                for xml_item in xml_items:
                    rebuilt_item = {
                        "operation_type": "Saida",
                        "document_number": str(document_meta.get("document_number", "")).strip() or str(xml_item.get("numero_nfce", "")).strip(),
                        "document_key": document_key,
                        "document_date": str(document_meta.get("document_date", "")).strip() or str(xml_item.get("data_emissao", "")).strip(),
                        "document_series": str(document_meta.get("document_series", "")).strip() or str(xml_item.get("serie", "")).strip(),
                        "document_model": str(document_meta.get("cod_mod", "")).strip() or str(xml_item.get("modelo", "")).strip() or "55",
                        "participant_code": "",
                        "participant_name": str(xml_item.get("emitente", "")).strip() or f"XML modelo {str(document_meta.get('cod_mod', '')).strip() or '55'}",
                        "participant_tax_id": str(xml_item.get("cnpj_emitente", "")).strip(),
                        "item_number": str(xml_item.get("item", "")).strip(),
                        "code": str(xml_item.get("codigo", "")).strip() or str(xml_item.get("descricao", "")).strip(),
                        "description": str(xml_item.get("descricao", "")).strip(),
                        "ncm": str(xml_item.get("ncm", "")).strip(),
                        "cest": str(xml_item.get("cest", "")).strip(),
                        "cst_icms": compose_xml_icms_cst_for_sped(xml_item),
                        "cfop": str(xml_item.get("cfop", "")).strip(),
                        "quantity": Decimal(xml_item.get("quantidade", Decimal("0"))),
                        "icms_rate": Decimal(xml_item.get("aliquota_icms", Decimal("0"))),
                        "sale_value": Decimal(xml_item.get("valor_operacao", xml_item.get("valor_produto", Decimal("0")))),
                        "base_icms": Decimal(xml_item.get("base_icms", Decimal("0"))),
                        "icms_value": Decimal(xml_item.get("valor_icms", Decimal("0"))),
                        "base_icms_st": Decimal(xml_item.get("base_icms_st", Decimal("0"))),
                        "icms_st_rate": Decimal(xml_item.get("aliquota_icms_st", Decimal("0"))),
                        "icms_st_value": Decimal(xml_item.get("valor_icms_st", Decimal("0"))),
                        "ipi_value": Decimal(xml_item.get("valor_ipi", Decimal("0"))),
                        "note_snapshot": note_snapshot,
                    }
                    document_rebuilt_items.append(rebuilt_item)
                    identity = (
                        document_key or str(rebuilt_item["document_number"]).strip(),
                        str(rebuilt_item["document_date"]).strip(),
                    )
                    bucket = documents_by_identity.setdefault(
                        identity,
                        {
                            "document_number": rebuilt_item["document_number"],
                            "document_key": document_key,
                            "document_date": rebuilt_item["document_date"],
                            "document_series": rebuilt_item["document_series"],
                            "document_model": rebuilt_item["document_model"],
                            "participant_code": "",
                            "participant_name": rebuilt_item["participant_name"],
                            "participant_tax_id": rebuilt_item["participant_tax_id"],
                            "items": [],
                        },
                    )
                    bucket["items"].append(rebuilt_item)
                document_rebuilt_items = normalize_xml_rebuilt_items_with_fallback(
                    document_rebuilt_items,
                    fallback_rows_by_document_key.get(document_key, []),
                )
                xml_rebuilt_sales.extend(document_rebuilt_items)
                continue

            fallback_rows = fallback_rows_by_document_key.get(document_key, [])
            for fallback_row in fallback_rows:
                rebuilt_fallback_row = dict(fallback_row)
                c190_fallback_sales.append(rebuilt_fallback_row)
                identity = (
                    document_key or str(rebuilt_fallback_row["document_number"]).strip(),
                    str(rebuilt_fallback_row["document_date"]).strip(),
                )
                bucket = documents_by_identity.setdefault(
                    identity,
                    {
                        "document_number": rebuilt_fallback_row["document_number"],
                        "document_key": document_key,
                        "document_date": rebuilt_fallback_row["document_date"],
                        "document_series": rebuilt_fallback_row["document_series"],
                        "document_model": rebuilt_fallback_row["document_model"],
                        "participant_code": rebuilt_fallback_row["participant_code"],
                        "participant_name": rebuilt_fallback_row["participant_name"],
                        "participant_tax_id": rebuilt_fallback_row["participant_tax_id"],
                        "items": [],
                    },
                )
                bucket["items"].append(dict(rebuilt_fallback_row))

        all_sale_items = list(sale_items)
        for item in all_sale_items:
            document_key = normalize_document_key(str(item.get("document_key", "")))
            if document_key in xml_note_snapshots:
                item["note_snapshot"] = xml_note_snapshots[document_key]
        all_sale_items.extend(xml_rebuilt_sales)
        all_sale_items.extend(c190_fallback_sales)

        label_base = infer_sped_period_label(sped_path, all_sale_items)
        label = label_base
        suffix = 2
        while label in used_labels:
            label = f"{label_base} ({suffix})"
            suffix += 1
        used_labels.add(label)
        period_labels.append(label)

        grouped_items: dict[str, list[dict[str, object]]] = defaultdict(list)
        for item in all_sale_items:
            code = str(item.get("code", "")).strip() or str(item.get("description", "")).strip()
            if code:
                grouped_items[code].append(item)

        for code, items in grouped_items.items():
            cfops = sorted({str(item.get("cfop", "")).strip() for item in items if str(item.get("cfop", "")).strip()})
            csts = sorted({str(item.get("cst_icms", "")).strip() for item in items if str(item.get("cst_icms", "")).strip()})
            descriptions = sorted({str(item.get("description", "")).strip() for item in items if str(item.get("description", "")).strip()})
            ncms = sorted({str(item.get("ncm", "")).strip() for item in items if str(item.get("ncm", "")).strip()})
            cests = collect_cest_values(items)
            document_keys = {
                normalize_document_key(str(item.get("document_key", "")))
                for item in items
                if normalize_document_key(str(item.get("document_key", "")))
            }
            launch_details = [
                {
                    "operation_type": "Saida",
                    "document_number": str(item.get("document_number", "")).strip(),
                    "document_key": str(item.get("document_key", "")).strip(),
                    "document_date": str(item.get("document_date", "")).strip(),
                    "item_number": str(item.get("item_number", "")).strip(),
                    "code": str(item.get("code", "")).strip(),
                    "cfop": str(item.get("cfop", "")).strip(),
                    "cst_icms": str(item.get("cst_icms", "")).strip(),
                    "icms_rate": Decimal(item.get("icms_rate", Decimal("0"))),
                    "quantity": Decimal(item.get("quantity", Decimal("0"))),
                    "sale_value": Decimal(item.get("sale_value", Decimal("0"))),
                    "base_icms": Decimal(item.get("base_icms", Decimal("0"))),
                    "icms_value": Decimal(item.get("icms_value", Decimal("0"))),
                    "base_icms_st": Decimal(item.get("base_icms_st", Decimal("0"))),
                    "icms_st_value": Decimal(item.get("icms_st_value", Decimal("0"))),
                    "ipi_value": Decimal(item.get("ipi_value", Decimal("0"))),
                    "description": str(item.get("description", "")).strip(),
                    "ncm": str(item.get("ncm", "")).strip(),
                    "cest": str(item.get("cest", "")).strip(),
                    "document_series": str(item.get("document_series", "")).strip(),
                    "document_model": str(item.get("document_model", "")).strip(),
                    "participant_code": str(item.get("participant_code", "")).strip(),
                    "participant_name": str(item.get("participant_name", "")).strip(),
                    "participant_tax_id": str(item.get("participant_tax_id", "")).strip(),
                    "note_snapshot": item.get("note_snapshot") or documents_by_identity.get(
                        (
                            normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip(),
                            str(item.get("document_date", "")).strip(),
                        )
                    ),
                }
                for item in sorted(
                    items,
                    key=lambda current: (
                        str(current.get("document_date", "")),
                        str(current.get("document_number", "")),
                        str(current.get("item_number", "")),
                    ),
                )
            ]
            comparison_rows.append(
                {
                    "period": label,
                    "file_name": sped_path.name,
                    "curve_abc": "C",
                    "code": code,
                    "description": " | ".join(descriptions),
                    "ncm": " | ".join(ncms),
                    "cest": " | ".join(cests),
                    "cfop": " | ".join(cfops),
                    "cst_icms": " | ".join(csts),
                    "suppliers": " | ".join(
                        sorted(
                            {
                                str(item.get("participant_name", "")).strip()
                                for item in items
                                if str(item.get("participant_name", "")).strip()
                            }
                        )
                    ),
                    "supplier_count": len(
                        {
                            (
                                str(item.get("participant_name", "")).strip(),
                                str(item.get("participant_tax_id", "")).strip(),
                            )
                            for item in items
                            if str(item.get("participant_name", "")).strip() or str(item.get("participant_tax_id", "")).strip()
                        }
                    ),
                    "quantity": sum((Decimal(item.get("quantity", Decimal("0"))) for item in items), Decimal("0")),
                    "sale_value": sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                    "base_icms": sum((Decimal(item.get("base_icms", Decimal("0"))) for item in items), Decimal("0")),
                    "icms_value": sum((Decimal(item.get("icms_value", Decimal("0"))) for item in items), Decimal("0")),
                    "display_icms_rate": compute_display_icms_rate(
                        next((Decimal(detail.get("icms_rate", Decimal("0"))) for detail in launch_details if Decimal(detail.get("icms_rate", Decimal("0"))) > 0), Decimal("0")),
                        sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                        sum((Decimal(item.get("base_icms", Decimal("0"))) for item in items), Decimal("0")),
                        sum((Decimal(item.get("icms_value", Decimal("0"))) for item in items), Decimal("0")),
                    ),
                    "document_count": len(document_keys),
                    "launch_count": len(items),
                    "status": summarize_sale_analysis(items)["status"],
                    "has_icms_reduction": any(
                        has_icms_reduction(
                            detail.get("icms_rate", Decimal("0")),
                            detail.get("sale_value", Decimal("0")),
                            detail.get("base_icms", Decimal("0")),
                            detail.get("icms_value", Decimal("0")),
                        )
                        for detail in launch_details
                    ),
                    "launch_details": launch_details,
                }
            )

        if progress_callback:
            progress_callback(current_step, total_steps, f"Saidas processadas: {sped_path.name}")

    curve_labels = calculate_abc_curve_labels(
        {
            str(row.get("code", "")).strip(): Decimal(row.get("sale_value", Decimal("0")))
            for row in comparison_rows
            if str(row.get("code", "")).strip()
        }
    )
    for row in comparison_rows:
        row["curve_abc"] = curve_labels.get(str(row.get("code", "")).strip(), "C")

    comparison_rows.sort(
        key=lambda item: (
            str(item["code"]),
            str(item["period"]),
            str(item["cfop"]),
            str(item["cst_icms"]),
        )
    )
    return period_labels, comparison_rows


NUMERIC_FIELD_PATTERN = re.compile(r"^-?\d+(?:[.,]\d+)?$")


def normalize_field_for_log_compare(value: str) -> tuple[str, object]:
    text = str(value or "").strip()
    if not text:
        return ("empty", "")
    if NUMERIC_FIELD_PATTERN.fullmatch(text):
        return ("num", parse_decimal(text))
    return ("text", text)


def lines_differ_meaningfully(original_line: str, adjusted_line: str) -> bool:
    original_text = str(original_line or "").strip()
    adjusted_text = str(adjusted_line or "").strip()
    if not original_text and not adjusted_text:
        return False
    if not original_text or not adjusted_text:
        return True
    if not original_text.startswith("|") or not adjusted_text.startswith("|"):
        return original_text != adjusted_text

    original_fields = normalize_sped_line(original_text)
    adjusted_fields = normalize_sped_line(adjusted_text)
    max_len = max(len(original_fields), len(adjusted_fields))
    for index in range(max_len):
        original_field = original_fields[index] if index < len(original_fields) else ""
        adjusted_field = adjusted_fields[index] if index < len(adjusted_fields) else ""
        if normalize_field_for_log_compare(original_field) != normalize_field_for_log_compare(adjusted_field):
            return True
    return False


def filter_sales(
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
    cst_values: set[str],
    cfop_values: set[str],
) -> list[dict[str, object]]:
    filtered_sales = filter_detailed_sales(
        detailed_sales,
        descriptions,
        cst_values,
        cfop_values,
    )
    grouped_filtered: dict[tuple[str, str, Decimal], dict[str, object]] = {}

    for item in filtered_sales:
        key = (str(item["code"]), str(item["cst_icms"]), item["icms_rate"])
        if key not in grouped_filtered:
            grouped_filtered[key] = {
                "code": item["code"],
                "description": item["description"],
                "ncm": item["ncm"],
                "cst_icms": item["cst_icms"],
                "quantity": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_rate": item["icms_rate"],
                "icms_value": Decimal("0"),
            }

        grouped_filtered[key]["quantity"] += item["quantity"]
        grouped_filtered[key]["base_icms"] += item["base_icms"]
        grouped_filtered[key]["icms_value"] += item["icms_value"]

    return sorted(
        grouped_filtered.values(),
        key=lambda item: (str(item["description"]), str(item["cst_icms"]), item["icms_rate"]),
    )


def filter_detailed_sales_by_descriptions(
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
) -> list[dict[str, object]]:
    return filter_detailed_sales(detailed_sales, descriptions, set(), set())


def filter_detailed_sales(
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
    cst_values: set[str],
    cfop_values: set[str],
) -> list[dict[str, object]]:
    # O filtro principal aceita descricao, CST e CFOP em conjunto.
    normalized_targets = {normalize_text(description) for description in descriptions if normalize_text(description)}
    if not normalized_targets:
        normalized_targets = set()
    normalized_csts = {str(value).strip() for value in cst_values if str(value).strip()}
    normalized_cfops = {str(value).strip() for value in cfop_values if str(value).strip()}

    filtered_items: list[dict[str, object]] = []
    for item in detailed_sales:
        if normalized_targets and normalize_text(str(item["description"])) not in normalized_targets:
            continue
        if normalized_csts and str(item["cst_icms"]).strip() not in normalized_csts:
            continue
        if normalized_cfops and str(item["cfop"]).strip() not in normalized_cfops:
            continue
        filtered_items.append(item)
    return filtered_items


def apply_icms_rate_override(
    detailed_sales: list[dict[str, object]],
    icms_rate: Decimal | None,
) -> list[dict[str, object]]:
    # Sobrescreve a aliquota dos itens filtrados e recalcula o ICMS
    # a partir da base de ICMS ja existente no documento.
    if icms_rate is None:
        return list(detailed_sales)

    recalculated_sales: list[dict[str, object]] = []
    for item in detailed_sales:
        base_icms = item["base_icms"] if isinstance(item["base_icms"], Decimal) else Decimal("0")
        recalculated_sales.append(
            {
                **item,
                "icms_rate": icms_rate,
                "icms_value": (base_icms * icms_rate / Decimal("100")).quantize(Decimal("0.01")),
            }
        )
    return recalculated_sales


def rebuild_detailed_sales_with_override(
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
    cst_values: set[str],
    cfop_values: set[str],
    icms_rate: Decimal | None,
    new_cst: str,
    new_cfop: str,
    zero_icms_cfops: set[str],
    selected_rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> list[dict[str, object]]:
    normalized_targets = {normalize_text(description) for description in descriptions if normalize_text(description)}
    normalized_csts = {str(value).strip() for value in cst_values if str(value).strip()}
    normalized_cfops = {str(value).strip() for value in cfop_values if str(value).strip()}
    rebuilt_sales: list[dict[str, object]] = []

    for item in detailed_sales:
        description_matches = not normalized_targets or normalize_text(str(item["description"])) in normalized_targets
        cst_matches = not normalized_csts or str(item["cst_icms"]).strip() in normalized_csts
        cfop_matches = not normalized_cfops or str(item["cfop"]).strip() in normalized_cfops
        should_apply_default_rule = has_configured_icms_rule(selected_rule_profile, runtime_rules, item)
        if should_apply_default_rule or (description_matches and cst_matches and cfop_matches):
            source_items = (
                expand_configured_icms_rule_items(selected_rule_profile, runtime_rules, item)
                if should_apply_default_rule
                else [dict(item)]
            )
            for source_item in source_items:
                rebuilt_item = dict(source_item)
                rebuilt_item.setdefault("original_cfop", str(source_item.get("original_cfop", source_item.get("cfop", ""))).strip())
                rebuilt_item.setdefault("original_cst_icms", str(source_item.get("original_cst_icms", source_item.get("cst_icms", ""))).strip())
                base_icms = rebuilt_item["base_icms"] if isinstance(rebuilt_item["base_icms"], Decimal) else Decimal("0")
                sale_value = rebuilt_item["sale_value"] if isinstance(rebuilt_item["sale_value"], Decimal) else Decimal("0")
                preserved_icms_value = rebuilt_item["icms_value"] if isinstance(rebuilt_item.get("icms_value"), Decimal) else Decimal("0")
                preserve_icms_value = bool(rebuilt_item.get("_preserve_icms_value"))
                had_fiscal_override = bool(new_cst or new_cfop or icms_rate is not None)
                if new_cst:
                    rebuilt_item["cst_icms"] = new_cst
                if new_cfop:
                    rebuilt_item["cfop"] = new_cfop
                final_cfop = str(rebuilt_item["cfop"]).strip()
                if final_cfop in zero_icms_cfops:
                    rebuilt_item["base_icms"] = Decimal("0")
                    rebuilt_item["icms_rate"] = Decimal("0")
                    rebuilt_item["icms_value"] = Decimal("0")
                else:
                    final_rate = icms_rate if icms_rate is not None else (
                        rebuilt_item["icms_rate"] if isinstance(rebuilt_item["icms_rate"], Decimal) else Decimal("0")
                    )
                    if had_fiscal_override and base_icms == Decimal("0") and sale_value > Decimal("0") and final_rate > Decimal("0"):
                        base_icms = sale_value
                    rebuilt_item["base_icms"] = base_icms
                    rebuilt_item["icms_rate"] = final_rate
                    if preserve_icms_value:
                        rebuilt_item["icms_value"] = preserved_icms_value
                    else:
                        rebuilt_item["icms_value"] = (
                            (base_icms * final_rate / Decimal("100")).quantize(Decimal("0.01"))
                            if base_icms > Decimal("0") and final_rate > Decimal("0")
                            else Decimal("0")
                        )
                rebuilt_sales.append(apply_sped_icms_consistency_rules(rebuilt_item, selected_rule_profile, runtime_rules))
            continue
        preserved_item = dict(item)
        preserved_item.setdefault("original_cfop", str(item.get("original_cfop", item.get("cfop", ""))).strip())
        preserved_item.setdefault("original_cst_icms", str(item.get("original_cst_icms", item.get("cst_icms", ""))).strip())
        rebuilt_sales.append(apply_sped_icms_consistency_rules(preserved_item, selected_rule_profile, runtime_rules))

    return rebalance_document_rule_icms_values(rebuilt_sales)


def rebalance_document_rule_icms_values(
    detailed_sales: list[dict[str, object]],
) -> list[dict[str, object]]:
    document_totals: dict[tuple[str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    explicit_target_groups: dict[str, list[int]] = defaultdict(list)
    recalculated_groups: dict[tuple[str, str, str], list[int]] = defaultdict(list)

    for index, item in enumerate(detailed_sales):
        operation_type = str(item.get("operation_type", "")).strip().lower()
        document_id = normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
        if not document_id:
            continue
        sale_value = item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else Decimal("0")
        document_totals[(operation_type, document_id)] += sale_value
        rule_signature = str(item.get("_configured_rule_signature", "")).strip()
        if not rule_signature:
            continue
        if isinstance(item.get("_document_target_icms_total"), Decimal):
            explicit_target_groups[rule_signature].append(index)
        elif item.get("_document_recalculate_icms"):
            recalculated_groups[(operation_type, document_id, rule_signature)].append(index)

    if not explicit_target_groups and not recalculated_groups:
        return detailed_sales

    adjusted_sales = list(detailed_sales)
    for rule_signature, indexes in explicit_target_groups.items():
        matched_items = [adjusted_sales[index] for index in indexes]
        explicit_targets = [
            item.get("_document_target_icms_total")
            for item in matched_items
            if isinstance(item.get("_document_target_icms_total"), Decimal)
        ]
        if not explicit_targets:
            continue
        target_total = explicit_targets[0].quantize(Decimal("0.01"))
        total_capacity = sum(
            (
                item["sale_value"]
                if isinstance(item.get("sale_value"), Decimal) and item["sale_value"] > Decimal("0")
                else Decimal("0")
            )
            for item in matched_items
        ).quantize(Decimal("0.01"))
        if total_capacity <= Decimal("0"):
            continue
        target_total = min(target_total, total_capacity)
        allocated_values = distribute_icms_with_caps(matched_items, target_total)
        if not allocated_values:
            continue

        applied_total = sum(allocated_values, Decimal("0")).quantize(Decimal("0.01"))
        for index, new_icms_value in zip(indexes, allocated_values):
            adjusted_item = dict(adjusted_sales[index])
            operation_type = str(adjusted_item.get("operation_type", "")).strip().lower()
            document_id = normalize_document_key(str(adjusted_item.get("document_key", ""))) or str(adjusted_item.get("document_number", "")).strip()
            adjusted_item["icms_value"] = new_icms_value
            adjusted_item["_preserve_icms_value"] = True
            adjusted_item["_rateio_group"] = rule_signature
            adjusted_item["_rateio_target_total"] = target_total
            adjusted_item["_rateio_applied_total"] = applied_total
            adjusted_item["_rateio_document_cap"] = document_totals[(operation_type, document_id)].quantize(Decimal("0.01"))
            adjusted_sales[index] = adjusted_item

    for group_key, indexes in recalculated_groups.items():
        operation_type, document_id, rule_signature = group_key
        matched_items = [adjusted_sales[index] for index in indexes]
        document_total = document_totals[(operation_type, document_id)].quantize(Decimal("0.01"))
        if document_total <= Decimal("0"):
            continue

        current_total = sum(
            (
                item["icms_value"]
                if isinstance(item.get("icms_value"), Decimal)
                else Decimal("0")
            )
            for item in matched_items
        ).quantize(Decimal("0.01"))
        target_total: Decimal | None = document_total if current_total > document_total else None
        if target_total is None or current_total == target_total:
            continue
        adjusted_values = distribute_icms_with_caps(matched_items, target_total)
        if not adjusted_values:
            continue
        applied_total = sum(adjusted_values, Decimal("0")).quantize(Decimal("0.01"))
        for index, new_icms_value in zip(indexes, adjusted_values):
            adjusted_item = dict(adjusted_sales[index])
            adjusted_item["icms_value"] = new_icms_value
            adjusted_item["_preserve_icms_value"] = True
            adjusted_item["_rateio_group"] = rule_signature
            adjusted_item["_rateio_target_total"] = target_total
            adjusted_item["_rateio_applied_total"] = applied_total
            adjusted_item["_rateio_document_cap"] = document_total
            adjusted_sales[index] = adjusted_item

    return adjusted_sales


def distribute_icms_with_caps(
    items: list[dict[str, object]],
    target_total: Decimal,
) -> list[Decimal]:
    if not items:
        return []

    capped_target = target_total.quantize(Decimal("0.01"))
    weights: list[Decimal] = []
    capacities: list[Decimal] = []
    for item in items:
        base_icms = item["base_icms"] if isinstance(item.get("base_icms"), Decimal) else Decimal("0")
        sale_value = item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else Decimal("0")
        quantity = item["quantity"] if isinstance(item.get("quantity"), Decimal) else Decimal("0")
        current_icms = item["icms_value"] if isinstance(item.get("icms_value"), Decimal) else Decimal("0")
        if base_icms > Decimal("0"):
            weights.append(base_icms)
        elif sale_value > Decimal("0"):
            weights.append(sale_value)
        elif quantity > Decimal("0"):
            weights.append(quantity)
        elif current_icms > Decimal("0"):
            weights.append(current_icms)
        else:
            weights.append(Decimal("1"))
        capacities.append(max(sale_value, Decimal("0")).quantize(Decimal("0.01")))

    values = [Decimal("0.00")] * len(items)
    active_indexes = {index for index, capacity in enumerate(capacities) if capacity > Decimal("0")}
    remaining_target = min(capped_target, sum(capacities, Decimal("0")).quantize(Decimal("0.01")))

    while active_indexes and remaining_target > Decimal("0"):
        total_weight = sum(weights[index] for index in active_indexes)
        if total_weight <= Decimal("0"):
            total_weight = Decimal(len(active_indexes))
            for index in active_indexes:
                weights[index] = Decimal("1")

        allocated_this_round = Decimal("0")
        clamped_indexes: set[int] = set()
        ordered_indexes = sorted(active_indexes)
        for position, index in enumerate(ordered_indexes):
            remaining_capacity = (capacities[index] - values[index]).quantize(Decimal("0.01"))
            if remaining_capacity <= Decimal("0"):
                clamped_indexes.add(index)
                continue
            if position == len(ordered_indexes) - 1:
                proposed = (remaining_target - allocated_this_round).quantize(Decimal("0.01"))
            else:
                proposed = (remaining_target * weights[index] / total_weight).quantize(Decimal("0.01"))
            if proposed > remaining_capacity:
                proposed = remaining_capacity
                clamped_indexes.add(index)
            values[index] = (values[index] + proposed).quantize(Decimal("0.01"))
            allocated_this_round += proposed

        remaining_target = (remaining_target - allocated_this_round).quantize(Decimal("0.01"))
        active_indexes -= clamped_indexes
        if allocated_this_round <= Decimal("0"):
            break

    return values


def prefer_excel_details_for_sped(
    detailed_sales: list[dict[str, object]],
    affected_documents: set[tuple[str, str]],
) -> list[dict[str, object]]:
    # Quando SPED e planilha sao usados juntos no ajuste, a planilha vira a fonte
    # preferencial dos valores fiscais e o SPED preserva a estrutura do item.
    excel_details_by_document: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    sped_details: list[dict[str, object]] = []

    for item in detailed_sales:
        document_identity = (
            str(item["operation_type"]),
            normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
        )
        if str(item.get("item_number", "")).strip():
            sped_details.append(item)
            continue
        excel_details_by_document[document_identity].append(dict(item))

    sped_by_document: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    for item in sped_details:
        document_identity = (
            str(item["operation_type"]),
            normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
        )
        sped_by_document[document_identity].append(item)

    preferred_details: list[dict[str, object]] = []
    for document_identity, sped_items in sped_by_document.items():
        if document_identity not in affected_documents:
            preferred_details.extend(dict(item) for item in sped_items)
            continue

        candidates = [dict(item) for item in excel_details_by_document.get(document_identity, [])]
        if not candidates:
            preferred_details.extend(dict(item) for item in sped_items)
            continue

        matched_items: list[dict[str, object]] = []
        remaining_sped = [dict(item) for item in sped_items]

        for sped_item in remaining_sped:
            match_index = None
            for index, candidate in enumerate(candidates):
                if (
                    normalize_text(str(candidate["description"])) == normalize_text(str(sped_item["description"]))
                    and candidate["sale_value"] == sped_item["sale_value"]
                ):
                    match_index = index
                    break
            if match_index is None:
                for index, candidate in enumerate(candidates):
                    if normalize_text(str(candidate["description"])) == normalize_text(str(sped_item["description"])):
                        match_index = index
                        break
            if match_index is None:
                for index, candidate in enumerate(candidates):
                    if candidate["sale_value"] == sped_item["sale_value"]:
                        match_index = index
                        break
            if match_index is None and candidates:
                match_index = 0

            if match_index is None:
                matched_items.append(sped_item)
                continue

            matched = candidates.pop(match_index)
            matched_items.append(
                {
                    **sped_item,
                    "description": matched["description"] or sped_item["description"],
                    "ncm": matched["ncm"] or sped_item["ncm"],
                    "cst_icms": matched["cst_icms"] or sped_item["cst_icms"],
                    "cfop": matched["cfop"] or sped_item["cfop"],
                    "quantity": matched["quantity"] if isinstance(matched["quantity"], Decimal) else sped_item["quantity"],
                    "sale_value": matched["sale_value"] if isinstance(matched["sale_value"], Decimal) else sped_item["sale_value"],
                    "base_icms": matched["base_icms"] if isinstance(matched["base_icms"], Decimal) else sped_item["base_icms"],
                    "icms_rate": matched["icms_rate"] if isinstance(matched["icms_rate"], Decimal) else sped_item["icms_rate"],
                    "icms_value": matched["icms_value"] if isinstance(matched["icms_value"], Decimal) else sped_item["icms_value"],
                    "base_icms_st": matched["base_icms_st"] if isinstance(matched.get("base_icms_st"), Decimal) else sped_item.get("base_icms_st", Decimal("0")),
                    "icms_st_rate": matched["icms_st_rate"] if isinstance(matched.get("icms_st_rate"), Decimal) else sped_item.get("icms_st_rate", Decimal("0")),
                    "icms_st_value": matched["icms_st_value"] if isinstance(matched.get("icms_st_value"), Decimal) else sped_item.get("icms_st_value", Decimal("0")),
                }
            )

        preferred_details.extend(matched_items)

    for document_identity, excel_items in excel_details_by_document.items():
        if document_identity not in affected_documents:
            continue
        if document_identity in sped_by_document:
            continue
        preferred_details.extend(dict(item) for item in excel_items)

    return preferred_details


def rebalance_bella_citta_1252_1253(detailed_sales: list[dict[str, object]], rule_profile: str) -> list[dict[str, object]]:
    if rule_profile not in {"BELLA CITTA 0001-02 - ENTRADAS-Saidas", "CASA DE PAES - ENTRADAS"}:
        return list(detailed_sales)

    documents: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    for item in detailed_sales:
        key = (
            str(item["operation_type"]),
            normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
        )
        documents[key].append(item)

    rebalanced_sales: list[dict[str, object]] = []
    for key, items in documents.items():
        if rule_profile == "CASA DE PAES - ENTRADAS":
            pooled_items = [
                item for item in items
                if str(item.get("operation_type", "")).strip().lower() == "entrada"
                and normalize_cst_icms_for_sped(str(item.get("cst_icms", ""))) in {"000", "090"}
                and str(item.get("cfop", "")).strip() in {"1252", "1253"}
            ]
        else:
            pooled_items = [
                item for item in items
                if str(item.get("operation_type", "")).strip().lower() == "entrada"
                and (
                    (normalize_cst_icms_for_sped(str(item.get("cst_icms", ""))) == "000" and str(item.get("cfop", "")).strip() in {"1252", "1253"})
                    or (normalize_cst_icms_for_sped(str(item.get("cst_icms", ""))) == "090" and str(item.get("cfop", "")).strip() == "1253")
                )
            ]
        if not pooled_items:
            rebalanced_sales.extend(items)
            continue

        remaining_items = [item for item in items if item not in pooled_items]
        template_item = dict(pooled_items[0])
        total_sale_value = sum(
            (item["sale_value"] for item in pooled_items if isinstance(item.get("sale_value"), Decimal)),
            Decimal("0"),
        ).quantize(Decimal("0.01"))
        total_quantity = sum(
            (item["quantity"] for item in pooled_items if isinstance(item.get("quantity"), Decimal)),
            Decimal("0"),
        ).quantize(Decimal("0.01"))

        ten_percent_value = (total_sale_value * Decimal("0.10")).quantize(Decimal("0.01"))
        ninety_percent_value = (total_sale_value - ten_percent_value).quantize(Decimal("0.01"))
        ten_percent_quantity = (total_quantity * Decimal("0.10")).quantize(Decimal("0.01"))
        ninety_percent_quantity = (total_quantity - ten_percent_quantity).quantize(Decimal("0.01"))

        item_1253 = {
            **template_item,
            "cst_icms": "090",
            "cfop": "1253",
            "quantity": ten_percent_quantity,
            "sale_value": ten_percent_value,
            "total_operation_value": ten_percent_value,
            "base_icms": Decimal("0"),
            "icms_rate": Decimal("0"),
            "icms_value": Decimal("0"),
        }
        item_1252 = {
            **template_item,
            "cst_icms": "000",
            "cfop": "1252",
            "quantity": ninety_percent_quantity,
            "sale_value": ninety_percent_value,
            "total_operation_value": ninety_percent_value,
            "base_icms": ninety_percent_value,
            "icms_rate": Decimal("18"),
            "icms_value": (ninety_percent_value * Decimal("18") / Decimal("100")).quantize(Decimal("0.01")),
        }

        rebalanced_sales.extend(remaining_items)
        rebalanced_sales.extend([item_1253, item_1252])

    return rebalanced_sales


def force_icms_total_for_target_group(
    detailed_sales: list[dict[str, object]],
    operation_type: str,
    cst_icms: str,
    cfop: str,
    target_total: Decimal,
    single_item_deduction: Decimal = Decimal("0"),
) -> list[dict[str, object]]:
    normalized_operation = str(operation_type).strip().lower()
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)
    normalized_cfop = str(cfop).strip()

    matched_indexes: list[int] = []
    matched_items: list[dict[str, object]] = []
    current_total = Decimal("0")

    for index, item in enumerate(detailed_sales):
        item_operation = str(item.get("operation_type", "")).strip().lower()
        item_cst = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
        item_cfop = str(item.get("cfop", "")).strip()
        if item_operation != normalized_operation or item_cst != normalized_cst or item_cfop != normalized_cfop:
            continue
        matched_indexes.append(index)
        matched_items.append(dict(item))
        current_total += item["icms_value"] if isinstance(item.get("icms_value"), Decimal) else Decimal("0")

    if not matched_items:
        return detailed_sales

    quantized_target = target_total.quantize(Decimal("0.01"))
    if current_total.quantize(Decimal("0.01")) == quantized_target:
        return detailed_sales

    weights: list[Decimal] = []
    for item in matched_items:
        icms_value = item["icms_value"] if isinstance(item.get("icms_value"), Decimal) else Decimal("0")
        quantity = item["quantity"] if isinstance(item.get("quantity"), Decimal) else Decimal("0")
        sale_value = item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else Decimal("0")
        if icms_value > Decimal("0"):
            weights.append(icms_value)
        elif quantity > Decimal("0"):
            weights.append(quantity)
        elif sale_value > Decimal("0"):
            weights.append(sale_value)
        else:
            weights.append(Decimal("1"))

    total_weight = sum(weights, Decimal("0"))
    if total_weight <= Decimal("0"):
        weights = [Decimal("1")] * len(matched_items)
        total_weight = Decimal(len(matched_items))

    adjusted_values: list[Decimal] = []
    allocated_total = Decimal("0")
    for position, weight in enumerate(weights):
        if position == len(weights) - 1:
            new_value = (quantized_target - allocated_total).quantize(Decimal("0.01"))
        else:
            new_value = (quantized_target * weight / total_weight).quantize(Decimal("0.01"))
            allocated_total += new_value
        adjusted_values.append(new_value)

    if adjusted_values and single_item_deduction != Decimal("0"):
        adjusted_values[-1] = (adjusted_values[-1] - single_item_deduction).quantize(Decimal("0.01"))

    adjusted_sales = list(detailed_sales)
    for index, new_icms_value in zip(matched_indexes, adjusted_values):
        adjusted_item = dict(adjusted_sales[index])
        adjusted_item["icms_value"] = new_icms_value
        adjusted_item["_preserve_icms_value"] = True
        adjusted_sales[index] = adjusted_item

    return adjusted_sales


def build_document_c190_map(
    c190_rows: list[dict[str, object]],
) -> dict[tuple[str, str], list[dict[str, object]]]:
    documents: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    for item in c190_rows:
        key = (
            str(item["operation_type"]),
            normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
        )
        documents[key].append(item)
    return documents


def consolidate_c190_items(c190_items: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped_items: dict[tuple[str, str, str, str, Decimal], dict[str, object]] = {}

    for item in c190_items:
        icms_rate = item["icms_rate"] if isinstance(item.get("icms_rate"), Decimal) else Decimal("0")
        key = (
            str(item.get("operation_type", "")),
            str(item.get("document_number", "")),
            normalize_cst_icms_for_sped(str(item.get("cst_icms", ""))),
            str(item.get("cfop", "")),
            icms_rate.quantize(Decimal("0.01")),
        )
        if key not in grouped_items:
            grouped_items[key] = {
                **item,
                "cst_icms": normalize_cst_icms_for_sped(str(item.get("cst_icms", ""))),
                "icms_rate": key[-1],
                "total_operation_value": Decimal("0"),
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "reduction_value": Decimal("0"),
                "ipi_value": Decimal("0"),
            }

        grouped_items[key]["total_operation_value"] += (
            item["total_operation_value"] if isinstance(item.get("total_operation_value"), Decimal) else Decimal("0")
        )
        grouped_items[key]["sale_value"] += item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else Decimal("0")
        grouped_items[key]["base_icms"] += item["base_icms"] if isinstance(item.get("base_icms"), Decimal) else Decimal("0")
        grouped_items[key]["icms_value"] += item["icms_value"] if isinstance(item.get("icms_value"), Decimal) else Decimal("0")
        grouped_items[key]["base_icms_st"] += (
            item["base_icms_st"] if isinstance(item.get("base_icms_st"), Decimal) else Decimal("0")
        )
        grouped_items[key]["icms_st_value"] += (
            item["icms_st_value"] if isinstance(item.get("icms_st_value"), Decimal) else Decimal("0")
        )
        grouped_items[key]["reduction_value"] += (
            item["reduction_value"] if isinstance(item.get("reduction_value"), Decimal) else Decimal("0")
        )
        grouped_items[key]["ipi_value"] += item["ipi_value"] if isinstance(item.get("ipi_value"), Decimal) else Decimal("0")

    return list(grouped_items.values())


def consolidate_summary_items(summary_items: list[dict[str, object]]) -> list[dict[str, object]]:
    return consolidate_c190_items(summary_items)


def build_document_detail_map(
    detailed_sales: list[dict[str, object]],
) -> dict[tuple[str, str], dict[str, list[dict[str, object]]]]:
    documents: dict[tuple[str, str], dict[str, list[dict[str, object]]]] = defaultdict(lambda: defaultdict(list))
    for item in detailed_sales:
        item_number = str(item.get("item_number", "")).strip()
        if not item_number:
            continue
        key = (
            str(item["operation_type"]),
            normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
        )
        documents[key][item_number].append(item)
    return documents


def parse_c190_item_from_line(
    line: str,
    operation_type: str,
    document_number: str,
    document_key: str,
    document_date: str,
) -> dict[str, object]:
    fields = normalize_sped_line(line)
    return {
        "operation_type": operation_type,
        "document_number": document_number,
        "document_key": document_key,
        "document_date": document_date,
        "item_number": "",
        "code": "",
        "description": "",
        "ncm": "",
        "cst_icms": get_field(fields, 2),
        "cfop": get_field(fields, 3),
        "quantity": Decimal("0"),
        "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
        "total_operation_value": parse_decimal(get_field(fields, 5)),
        "sale_value": parse_decimal(get_field(fields, 5)),
        "base_icms": parse_decimal(get_field(fields, 6)),
        "icms_value": parse_decimal(get_field(fields, 7)),
        "base_icms_st": parse_decimal(get_field(fields, 8)),
        "icms_st_value": parse_decimal(get_field(fields, 9)),
        "reduction_value": parse_decimal(get_field(fields, 10)),
        "ipi_value": parse_decimal(get_field(fields, 11)),
    }


def parse_c590_item_from_line(
    line: str,
    operation_type: str,
    document_number: str,
    document_key: str,
    document_date: str,
) -> dict[str, object]:
    fields = normalize_sped_line(line)
    return {
        "operation_type": operation_type,
        "document_number": document_number,
        "document_key": document_key,
        "document_date": document_date,
        "item_number": "",
        "code": "",
        "description": "",
        "ncm": "",
        "cst_icms": get_field(fields, 2),
        "cfop": get_field(fields, 3),
        "quantity": Decimal("0"),
        "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
        "total_operation_value": parse_decimal(get_field(fields, 5)),
        "sale_value": parse_decimal(get_field(fields, 5)),
        "base_icms": parse_decimal(get_field(fields, 6)),
        "icms_value": parse_decimal(get_field(fields, 7)),
        "base_icms_st": parse_decimal(get_field(fields, 8)),
        "icms_st_value": parse_decimal(get_field(fields, 9)),
        "reduction_value": Decimal("0"),
        "ipi_value": Decimal("0"),
    }


def scan_default_filial_rule_documents(
    sped_path: Path,
    operation_scope: str,
    rule_profile: str,
    runtime_rules: list[dict[str, object]] | None = None,
) -> set[tuple[str, str]]:
    scoped_operations = {str(item["operation_type"]) for item in filter_details_by_operation_scope(
        [{"operation_type": "Entrada"}, {"operation_type": "Saida"}],
        operation_scope,
    )}
    affected_documents: set[tuple[str, str]] = set()
    current_operation = ""
    current_document = ""
    current_document_key = ""

    for line in read_sped_plain_lines(sped_path):
        if line.startswith("|C100|"):
            fields = normalize_sped_line(line)
            ind_oper = get_field(fields, 2)
            current_operation = "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else ""
            current_document = get_field(fields, 8)
            current_document_key = normalize_document_key(get_field(fields, 9))
            continue
        if line.startswith("|C500|"):
            fields = normalize_sped_line(line)
            ind_oper = get_field(fields, 2)
            current_operation = "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else ""
            current_document = get_field(fields, 10)
            current_document_key = normalize_document_key(get_field(fields, 10))
            continue

        if current_operation not in scoped_operations:
            continue

        if line.startswith("|C190|"):
            c190_item = parse_c190_item_from_line(line, current_operation, current_document, current_document_key, "")
            if has_configured_icms_rule(rule_profile, runtime_rules, c190_item):
                affected_documents.add((current_operation, current_document_key or current_document))
            continue

        if line.startswith("|C590|"):
            c590_item = parse_c590_item_from_line(line, current_operation, current_document, current_document_key, "")
            if has_configured_icms_rule(rule_profile, runtime_rules, c590_item):
                affected_documents.add((current_operation, current_document_key or current_document))

    return affected_documents


def apply_override_to_c170_line(line: str, detail: dict[str, object], item_number: str | None = None) -> str:
    fields = normalize_sped_line(line)
    if item_number is not None:
        set_field(fields, 2, item_number)
    set_field(fields, 5, format_decimal_sped(detail["quantity"] if isinstance(detail.get("quantity"), Decimal) else Decimal("0")))
    set_field(fields, 7, format_decimal_sped(detail["sale_value"] if isinstance(detail.get("sale_value"), Decimal) else Decimal("0")))
    set_field(fields, 10, normalize_cst_icms_for_sped(str(detail["cst_icms"])))
    set_field(fields, 11, str(detail["cfop"]))
    set_field(fields, 13, format_decimal_sped(detail["base_icms"] if isinstance(detail["base_icms"], Decimal) else Decimal("0")))
    set_field(fields, 14, format_decimal_sped(detail["icms_rate"] if isinstance(detail["icms_rate"], Decimal) else Decimal("0")))
    set_field(fields, 15, format_decimal_sped(detail["icms_value"] if isinstance(detail["icms_value"], Decimal) else Decimal("0")))
    set_field(fields, 16, format_decimal_sped(detail["base_icms_st"] if isinstance(detail.get("base_icms_st"), Decimal) else Decimal("0")))
    set_field(fields, 17, format_decimal_sped(detail["icms_st_rate"] if isinstance(detail.get("icms_st_rate"), Decimal) else Decimal("0")))
    set_field(fields, 18, format_decimal_sped(detail["icms_st_value"] if isinstance(detail.get("icms_st_value"), Decimal) else Decimal("0")))
    if isinstance(detail.get("ipi_value"), Decimal):
        set_field(fields, 24, format_decimal_sped(detail["ipi_value"]))
    return "|".join(fields)


def build_c190_line(item: dict[str, object]) -> str:
    # Converte a estrutura interna de um C190 de volta para a linha textual do SPED.
    fields = [
        "",
        "C190",
        normalize_cst_icms_for_sped(str(item["cst_icms"])),
        str(item["cfop"]),
        format_decimal_sped(item["icms_rate"] if isinstance(item["icms_rate"], Decimal) else Decimal("0")),
        format_decimal_sped(item["total_operation_value"]),
        format_decimal_sped(item["base_icms"]),
        format_decimal_sped(item["icms_value"]),
        format_decimal_sped(item["base_icms_st"]),
        format_decimal_sped(item["icms_st_value"]),
        format_decimal_sped(item["reduction_value"]),
        format_decimal_sped(item["ipi_value"]),
        "",
        "",
    ]
    return "|".join(fields)


def build_c590_line(item: dict[str, object]) -> str:
    fields = [
        "",
        "C590",
        normalize_cst_icms_for_sped(str(item["cst_icms"])),
        str(item["cfop"]),
        format_decimal_sped(item["icms_rate"] if isinstance(item["icms_rate"], Decimal) else Decimal("0")),
        format_decimal_sped(item["total_operation_value"]),
        format_decimal_sped(item["base_icms"]),
        format_decimal_sped(item["icms_value"]),
        format_decimal_sped(item["base_icms_st"] if isinstance(item.get("base_icms_st"), Decimal) else Decimal("0")),
        format_decimal_sped(item["icms_st_value"] if isinstance(item.get("icms_st_value"), Decimal) else Decimal("0")),
        "",
        "",
        "",
    ]
    return "|".join(fields)


def update_c100_totals(fields: list[str], c190_items: list[dict[str, object]]) -> None:
    # Atualiza os totais do C100 com base nos C190 recompostos.
    # Para NFC-e (modelo 65), os campos finais proibidos pelo PVA ficam em branco.
    total_sale_value = sum(
        (
            item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else item["total_operation_value"]
            for item in c190_items
            if isinstance(item.get("sale_value"), Decimal) or isinstance(item.get("total_operation_value"), Decimal)
        ),
        Decimal("0"),
    )
    total_base_icms = sum((item["base_icms"] for item in c190_items), Decimal("0"))
    total_icms = sum((item["icms_value"] for item in c190_items), Decimal("0"))
    total_base_icms_st = sum((item["base_icms_st"] for item in c190_items), Decimal("0"))
    total_icms_st = sum((item["icms_st_value"] for item in c190_items), Decimal("0"))
    total_ipi = sum((item["ipi_value"] for item in c190_items), Decimal("0"))
    cod_mod = get_field(fields, 5)
    total_discount = parse_decimal(get_field(fields, 14))
    total_abatement = parse_decimal(get_field(fields, 15))
    total_freight = parse_decimal(get_field(fields, 18))
    total_insurance = parse_decimal(get_field(fields, 19))
    total_other = parse_decimal(get_field(fields, 20))

    set_field(fields, 16, format_decimal_sped(total_sale_value))
    total_document_value = (
        total_sale_value
        - total_discount
        - total_abatement
        + total_freight
        + total_insurance
        + total_other
        + total_icms_st
        + total_ipi
    )
    if total_document_value < Decimal("0"):
        total_document_value = Decimal("0")
    set_field(fields, 12, format_decimal_sped(total_document_value))

    set_field(fields, 21, format_decimal_sped(total_base_icms))
    set_field(fields, 22, format_decimal_sped(total_icms))

    if cod_mod == "65":
        for index in range(23, 30):
            set_field(fields, index, "")
        return

    set_field(fields, 23, format_decimal_sped(total_base_icms_st))
    set_field(fields, 24, format_decimal_sped(total_icms_st))
    set_field(fields, 25, format_decimal_sped(total_ipi))


def update_c500_totals(fields: list[str], c590_items: list[dict[str, object]]) -> None:
    total_base_icms = sum((item["base_icms"] for item in c590_items), Decimal("0"))
    total_icms = sum((item["icms_value"] for item in c590_items), Decimal("0"))
    set_field(fields, 19, format_decimal_sped(total_base_icms))
    set_field(fields, 20, format_decimal_sped(total_icms))


def calculate_e110_totals(lines: list[str]) -> tuple[Decimal, Decimal]:
    total_debitos = Decimal("0")
    total_creditos = Decimal("0")
    credit_debit_registers = {"C190", "C590", "D190", "D590", "D730"}

    for line in lines:
        if not line.startswith("|"):
            continue
        fields = normalize_sped_line(line)
        register = get_field(fields, 1)
        if register not in credit_debit_registers:
            continue

        cfop = get_field(fields, 3)
        icms_value = parse_decimal(get_field(fields, 7))

        if cfop == "1605" or cfop.startswith("5") or cfop.startswith("6") or cfop.startswith("7"):
            total_debitos += icms_value
            continue
        if cfop == "5605" or cfop.startswith("1") or cfop.startswith("2") or cfop.startswith("3"):
            total_creditos += icms_value

    return total_debitos, total_creditos


EXCLUDED_E210_OUT_CREDIT_CFOPS = {
    "1410", "1411", "1414", "1415",
    "1603", "2603",
    "1660", "1661", "1662",
    "2410", "2411", "2414", "2415",
    "2660", "2661", "2662",
}


def is_e210_other_credit_cfop(cfop: str) -> bool:
    normalized_cfop = normalize_cfop(cfop)
    return bool(normalized_cfop) and normalized_cfop[0] in {"1", "2"} and normalized_cfop not in EXCLUDED_E210_OUT_CREDIT_CFOPS


def calculate_e210_out_credit_from_c190(lines: list[str]) -> Decimal:
    total_out_credit = Decimal("0")
    for line in lines:
        if not line.startswith("|C190|"):
            continue
        fields = normalize_sped_line(line)
        cfop = get_field(fields, 3)
        if not is_e210_other_credit_cfop(cfop):
            continue
        total_out_credit += parse_decimal(get_field(fields, 9))
    return total_out_credit


def calculate_e210_e220_out_credit(lines: list[str], e210_index: int) -> Decimal:
    total_out_credit = Decimal("0")
    index = e210_index + 1
    while index < len(lines):
        line = lines[index]
        if not line.startswith("|"):
            break
        if line.startswith("|E200|") or line.startswith("|E210|") or line.startswith("|E001|") or line.startswith("|E100|"):
            break
        if not line.startswith("|E220|"):
            index += 1
            continue

        fields = normalize_sped_line(line)
        cod_aj_apur = get_field(fields, 2)
        vl_aj_apur = parse_decimal(get_field(fields, 4))
        if len(cod_aj_apur) >= 4 and cod_aj_apur[2] == "1" and cod_aj_apur[3] in {"2", "3"}:
            total_out_credit += vl_aj_apur
        index += 1
    return total_out_credit


def recalculate_e210(lines: list[str]) -> None:
    c190_out_credit_total = calculate_e210_out_credit_from_c190(lines)

    for index, line in enumerate(lines):
        if not line.startswith("|E210|"):
            continue

        fields = normalize_sped_line(line)
        e220_out_credit_total = calculate_e210_e220_out_credit(lines, index)

        saldo_credor_anterior = parse_decimal(get_field(fields, 3))
        devolucao_st = parse_decimal(get_field(fields, 4))
        ressarcimento_st = parse_decimal(get_field(fields, 5))
        ajustes_credito_st = parse_decimal(get_field(fields, 7))
        retencao_st = parse_decimal(get_field(fields, 8))
        outros_debitos_st = parse_decimal(get_field(fields, 9))
        ajustes_debito_st = parse_decimal(get_field(fields, 10))
        deducoes_st = parse_decimal(get_field(fields, 12))
        deb_esp_st = parse_decimal(get_field(fields, 15))

        vl_out_cred_st = c190_out_credit_total + e220_out_credit_total

        saldo_devedor_antes_deducoes = (
            retencao_st
            + outros_debitos_st
            + ajustes_debito_st
            - saldo_credor_anterior
            - devolucao_st
            - ressarcimento_st
            - vl_out_cred_st
            - ajustes_credito_st
        )
        if saldo_devedor_antes_deducoes < Decimal("0"):
            saldo_devedor_antes_deducoes = Decimal("0")

        vl_icms_recol_st = saldo_devedor_antes_deducoes - deducoes_st
        if vl_icms_recol_st < Decimal("0"):
            vl_icms_recol_st = Decimal("0")

        vl_sld_cred_st_transportar = (
            saldo_credor_anterior
            + devolucao_st
            + ressarcimento_st
            + vl_out_cred_st
            + ajustes_credito_st
            + deducoes_st
            - retencao_st
            - outros_debitos_st
            - ajustes_debito_st
        )
        if vl_sld_cred_st_transportar < Decimal("0"):
            vl_sld_cred_st_transportar = Decimal("0")

        ind_mov_st = "1" if any(
            value != Decimal("0")
            for value in (
                saldo_credor_anterior,
                devolucao_st,
                ressarcimento_st,
                vl_out_cred_st,
                ajustes_credito_st,
                retencao_st,
                outros_debitos_st,
                ajustes_debito_st,
                deducoes_st,
                vl_icms_recol_st,
                vl_sld_cred_st_transportar,
                deb_esp_st,
            )
        ) else "0"

        set_field(fields, 2, ind_mov_st)
        set_field(fields, 6, format_decimal_sped(vl_out_cred_st))
        set_field(fields, 11, format_decimal_sped(saldo_devedor_antes_deducoes))
        set_field(fields, 12, format_decimal_sped(deducoes_st))
        set_field(fields, 13, format_decimal_sped(vl_icms_recol_st))
        set_field(fields, 14, format_decimal_sped(vl_sld_cred_st_transportar))
        lines[index] = "|".join(fields)


def recalculate_e110(lines: list[str]) -> None:
    # Recalcula a apuracao do ICMS a partir dos C190 ajustados,
    # mantendo E110 e E116 consistentes entre si.
    total_debitos, total_creditos = calculate_e110_totals(lines)

    for index, line in enumerate(lines):
        if not line.startswith("|E110|"):
            continue
        fields = normalize_sped_line(line)

        ajustes_debito_doc = parse_decimal(get_field(fields, 3))
        total_ajustes_debito = parse_decimal(get_field(fields, 4))
        estornos_credito = parse_decimal(get_field(fields, 5))
        ajustes_credito_doc = parse_decimal(get_field(fields, 7))
        total_ajustes_credito = parse_decimal(get_field(fields, 8))
        estornos_debito = parse_decimal(get_field(fields, 9))
        saldo_credor_anterior = parse_decimal(get_field(fields, 10))
        total_deducoes = parse_decimal(get_field(fields, 12))

        set_field(fields, 2, format_decimal_sped(total_debitos))
        set_field(fields, 6, format_decimal_sped(total_creditos))

        saldo_bruto = (
            total_debitos
            + ajustes_debito_doc
            + total_ajustes_debito
            + estornos_credito
            - total_creditos
            - ajustes_credito_doc
            - total_ajustes_credito
            - estornos_debito
            - saldo_credor_anterior
        )

        if saldo_bruto > Decimal("0"):
            saldo_apurado = saldo_bruto
            icms_recolher = saldo_apurado - total_deducoes
            if icms_recolher < Decimal("0"):
                icms_recolher = Decimal("0")

            set_field(fields, 11, format_decimal_sped(saldo_apurado))
            set_field(fields, 13, format_decimal_sped(icms_recolher))
            set_field(fields, 14, format_decimal_sped(Decimal("0")))
            e116_target = icms_recolher
        else:
            set_field(fields, 11, format_decimal_sped(Decimal("0")))
            set_field(fields, 13, format_decimal_sped(Decimal("0")))
            set_field(fields, 14, format_decimal_sped(abs(saldo_bruto)))
            e116_target = Decimal("0")
        lines[index] = "|".join(fields)
        recalculate_e116(lines, e116_target)
        break


def recalculate_e116(lines: list[str], target_total: Decimal) -> None:
    e116_indexes = [index for index, line in enumerate(lines) if line.startswith("|E116|")]
    if not e116_indexes:
        return

    current_values = []
    current_total = Decimal("0")
    for index in e116_indexes:
        fields = normalize_sped_line(lines[index])
        value = parse_decimal(get_field(fields, 3))
        current_values.append((index, fields, value))
        current_total += value

    assigned_total = Decimal("0")
    for position, (index, fields, value) in enumerate(current_values):
        if position == len(current_values) - 1:
            new_value = target_total - assigned_total
        elif current_total > 0:
            new_value = (target_total * value / current_total).quantize(Decimal("0.01"))
            assigned_total += new_value
        else:
            new_value = target_total if position == 0 else Decimal("0")
            assigned_total += new_value
        set_field(fields, 3, format_decimal_sped(new_value))
        lines[index] = "|".join(fields)


def normalize_sped_warning_fields(lines: list[str]) -> None:
    current_operation = ""
    for index, line in enumerate(lines):
        if line.startswith("|C100|"):
            fields = normalize_sped_line(line)
            current_operation = "entrada" if get_field(fields, 2) == "0" else "saida" if get_field(fields, 2) == "1" else ""
            continue

        if line.startswith("|C170|"):
            fields = normalize_sped_line(line)
            cst_suffix = normalize_cst_icms_for_sped(get_field(fields, 10))[-2:]
            cfop = get_field(fields, 11)
            current_icms = parse_decimal(get_field(fields, 15))
            preserve_special_icms = (
                current_operation == "entrada"
                and cst_suffix == "61"
                and (cfop in {"1651", "1653"} or current_icms > Decimal("0"))
            )
            if current_operation == "entrada" and cst_suffix in {"02", "15", "53", "61"}:
                set_field(fields, 13, format_decimal_sped(Decimal("0")))
                if not preserve_special_icms:
                    set_field(fields, 15, format_decimal_sped(Decimal("0")))
                if cst_suffix == "61":
                    set_field(fields, 14, format_decimal_sped(Decimal("0")))
                lines[index] = "|".join(fields)
            continue

        if not line.startswith("|C190|"):
            continue

        fields = normalize_sped_line(line)
        cst_suffix = normalize_cst_icms_for_sped(get_field(fields, 2))[-2:]
        cfop = get_field(fields, 3)
        vl_opr = parse_decimal(get_field(fields, 5))
        vl_bc = parse_decimal(get_field(fields, 6))
        vl_icms = parse_decimal(get_field(fields, 7))

        if cst_suffix in {"20", "70"}:
            reduction = vl_opr - vl_bc if vl_opr > vl_bc else Decimal("0")
            set_field(fields, 10, format_decimal_sped(reduction))

        if current_operation == "entrada" and cst_suffix in {"02", "15", "53", "61"}:
            set_field(fields, 6, format_decimal_sped(Decimal("0")))
            preserve_special_icms = cst_suffix == "61" and (cfop in {"1651", "1653"} or vl_icms > Decimal("0"))
            if not preserve_special_icms:
                set_field(fields, 7, format_decimal_sped(Decimal("0")))
            if cst_suffix == "61":
                set_field(fields, 4, format_decimal_sped(Decimal("0")))

        lines[index] = "|".join(fields)


def write_adjustment_log(
    output_path: Path,
    original_lines: list[str],
    adjusted_lines: list[str],
    affected_documents: set[tuple[str, str]],
    rateio_rows: list[list[object]] | None = None,
) -> None:
    log_path = output_path.with_name(f"{output_path.stem}_log_ajustes.xlsx")
    headers = ["tipo", "documento", "chave_nfe", "registro", "classificacao", "original", "ajustado"]
    changed_rows: list[list[object]] = []
    rewritten_rows: list[list[object]] = []

    original_map: dict[tuple[str, str], list[tuple[str, str, str]]] = defaultdict(list)
    adjusted_map: dict[tuple[str, str], list[tuple[str, str, str]]] = defaultdict(list)

    def collect(lines: list[str], target: dict[tuple[str, str], list[tuple[str, str, str]]]) -> None:
        current_operation = ""
        current_document = ""
        current_key = ""
        for line in lines:
            if line.startswith("|C100|"):
                fields = normalize_sped_line(line)
                current_operation = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
                current_document = get_field(fields, 8)
                current_key = normalize_document_key(get_field(fields, 9))
                key = (current_operation, current_key or current_document)
                if key in affected_documents:
                    target[key].append((line, current_document, current_key))
            elif line.startswith("|C170|") or line.startswith("|C190|") or line.startswith("|E110|"):
                key = (current_operation, current_key or current_document)
                if key in affected_documents or line.startswith("|E110|"):
                    target[key].append((line, current_document, current_key))

    collect(original_lines, original_map)
    collect(adjusted_lines, adjusted_map)

    for key in sorted(affected_documents):
        original_block = original_map.get(key, [])
        adjusted_block = adjusted_map.get(key, [])
        max_len = max(len(original_block), len(adjusted_block))
        for index in range(max_len):
            original_entry = original_block[index] if index < len(original_block) else ("", "", "")
            adjusted_entry = adjusted_block[index] if index < len(adjusted_block) else ("", "", "")
            original_line, original_document, original_key = original_entry
            adjusted_line, adjusted_document, adjusted_key = adjusted_entry
            document_number = adjusted_document or original_document
            document_key = adjusted_key or original_key
            register = ""
            if adjusted_line.startswith("|"):
                register = get_field(normalize_sped_line(adjusted_line), 1)
            elif original_line.startswith("|"):
                register = get_field(normalize_sped_line(original_line), 1)
            row = [
                key[0],
                document_number,
                document_key,
                register,
                "Alteracao Real" if lines_differ_meaningfully(original_line, adjusted_line) else "Regravacao Padrao",
                original_line,
                adjusted_line,
            ]
            if row[4] == "Alteracao Real":
                changed_rows.append(row)
            else:
                rewritten_rows.append(row)

    original_e110 = next((line for line in original_lines if line.startswith("|E110|")), "")
    adjusted_e110 = next((line for line in adjusted_lines if line.startswith("|E110|")), "")
    e110_row = [
        "Apuracao",
        "",
        "",
        "E110",
        "Alteracao Real" if lines_differ_meaningfully(original_e110, adjusted_e110) else "Regravacao Padrao",
        original_e110,
        adjusted_e110,
    ]
    if e110_row[4] == "Alteracao Real":
        changed_rows.append(e110_row)
    else:
        rewritten_rows.append(e110_row)

    workbook_sheets: list[tuple[str, list[str], list[list[object]]]] = [
        ("Alteracoes Reais", headers, changed_rows),
        ("Regravacoes Padrao", headers, rewritten_rows),
    ]
    if rateio_rows:
        workbook_sheets.append(
            (
                "Rateios",
                ["tipo", "documento", "chave_nfe", "codigo", "regra", "valor_documento", "total_solicitado", "total_aplicado", "valor_icms_item"],
                rateio_rows,
            )
        )
    write_simple_excel_workbook(log_path, workbook_sheets)


def build_rateio_log_rows(detailed_sales: list[dict[str, object]]) -> list[list[object]]:
    rows: list[list[object]] = []
    seen_keys: set[tuple[str, str, str, str]] = set()
    for item in detailed_sales:
        target_total = item.get("_rateio_target_total")
        applied_total = item.get("_rateio_applied_total")
        rule_signature = str(item.get("_rateio_group", "")).strip()
        if not isinstance(target_total, Decimal) or not isinstance(applied_total, Decimal) or not rule_signature:
            continue
        row_key = (
            str(item.get("operation_type", "")).strip(),
            str(item.get("document_number", "")).strip(),
            normalize_document_key(str(item.get("document_key", ""))),
            str(item.get("code", "")).strip(),
        )
        if row_key in seen_keys:
            continue
        seen_keys.add(row_key)
        rows.append(
            [
                str(item.get("operation_type", "")).strip(),
                str(item.get("document_number", "")).strip(),
                normalize_document_key(str(item.get("document_key", ""))),
                str(item.get("code", "")).strip(),
                rule_signature,
                format_decimal_sped(item["sale_value"] if isinstance(item.get("sale_value"), Decimal) else Decimal("0")),
                format_decimal_sped(target_total),
                format_decimal_sped(applied_total),
                format_decimal_sped(item["icms_value"] if isinstance(item.get("icms_value"), Decimal) else Decimal("0")),
            ]
        )
    return rows


def write_xml_st_adjustment_log(output_path: Path, adjustment_rows: list[dict[str, object]]) -> None:
    if not adjustment_rows:
        return

    log_path = output_path.with_name(f"{output_path.stem}_log_st_xml.xlsx")
    headers = [
        "tipo",
        "documento",
        "chave_nfe",
        "item",
        "codigo_item",
        "cfop_sped",
        "cfop_xml",
        "vl_bc_st_original",
        "aliq_st_original",
        "vl_icms_st_original",
        "vl_bc_st_ajustado",
        "aliq_st_ajustado",
        "vl_icms_st_ajustado",
    ]
    rows: list[list[object]] = []

    for item in sorted(
        adjustment_rows,
        key=lambda row: (
            str(row.get("operation_type", "")),
            str(row.get("document_number", "")),
            str(row.get("document_key", "")),
            str(row.get("item_number", "")),
        ),
    ):
        rows.append([
            str(item.get("operation_type", "")),
            str(item.get("document_number", "")),
            str(item.get("document_key", "")),
            str(item.get("item_number", "")),
            str(item.get("code", "")),
            str(item.get("cfop_sped", "")),
            str(item.get("cfop_xml", "")),
            format_decimal_sped(item.get("base_icms_st_original", Decimal("0")) if isinstance(item.get("base_icms_st_original"), Decimal) else Decimal("0")),
            format_decimal_sped(item.get("icms_st_rate_original", Decimal("0")) if isinstance(item.get("icms_st_rate_original"), Decimal) else Decimal("0")),
            format_decimal_sped(item.get("icms_st_value_original", Decimal("0")) if isinstance(item.get("icms_st_value_original"), Decimal) else Decimal("0")),
            format_decimal_sped(item.get("base_icms_st_ajustado", Decimal("0")) if isinstance(item.get("base_icms_st_ajustado"), Decimal) else Decimal("0")),
            format_decimal_sped(item.get("icms_st_rate_ajustado", Decimal("0")) if isinstance(item.get("icms_st_rate_ajustado"), Decimal) else Decimal("0")),
            format_decimal_sped(item.get("icms_st_value_ajustado", Decimal("0")) if isinstance(item.get("icms_st_value_ajustado"), Decimal) else Decimal("0")),
        ])

    write_simple_excel_workbook(log_path, [("Ajustes ST XML", headers, rows)])


def generate_adjusted_sped_lines(
    sped_path: Path,
    rebuilt_detailed_sales: list[dict[str, object]],
    affected_documents: set[tuple[str, str]],
    selected_rule_profile: str = "",
    runtime_rules: list[dict[str, object]] | None = None,
) -> tuple[list[str], list[str]]:
    # Reescreve somente os blocos C100 afetados, preservando o restante do SPED.
    c100_children = {
        "C101", "C105", "C110", "C111", "C112", "C113", "C114", "C115", "C116",
        "C120", "C130", "C140", "C141", "C160", "C165", "C170", "C171", "C172",
        "C173", "C174", "C175", "C176", "C177", "C178", "C179", "C190", "C191",
        "C195", "C197",
    }
    c500_children = {"C510", "C590", "C591", "C595", "C597"}
    recalculated_c190_rows, _ = build_c190_rows_from_details(rebuilt_detailed_sales, selected_rule_profile, runtime_rules)
    document_map = build_document_c190_map(recalculated_c190_rows)
    document_detail_map = build_document_detail_map(rebuilt_detailed_sales)

    source_lines = read_sped_plain_lines(sped_path)
    output_lines: list[str] = []
    index = 0

    while index < len(source_lines):
        current_line = source_lines[index]
        if not current_line.startswith("|C100|"):
            if current_line.startswith("|C500|"):
                block_lines = [current_line]
                index += 1
                while index < len(source_lines):
                    candidate = source_lines[index]
                    candidate_fields = normalize_sped_line(candidate) if candidate.startswith("|") else []
                    candidate_register = get_field(candidate_fields, 1)
                    if not candidate.startswith("|"):
                        break
                    if candidate_register and candidate_register.startswith("C") and candidate_register not in c500_children:
                        break
                    if candidate.startswith("|D001|"):
                        break
                    block_lines.append(candidate)
                    index += 1

                c500_fields = normalize_sped_line(block_lines[0])
                operation_type = "Entrada" if get_field(c500_fields, 2) == "0" else "Saida" if get_field(c500_fields, 2) == "1" else ""
                document_number = get_field(c500_fields, 10)
                document_key = normalize_document_key(get_field(c500_fields, 10))
                cod_sit = get_field(c500_fields, 6)
                document_identity = (operation_type, document_key or document_number)
                if document_identity not in affected_documents or cod_sit in {"02", "03", "04"}:
                    output_lines.extend(block_lines)
                    continue

                direct_c590_items: list[dict[str, object]] = []
                rebuilt_block = ["|".join(c500_fields)]
                non_c590_lines: list[str] = []
                insert_position = None

                for line in block_lines[1:]:
                    if line.startswith("|C590|"):
                        if insert_position is None:
                            insert_position = len(non_c590_lines)
                        parsed_item = parse_c590_item_from_line(
                            line,
                            operation_type,
                            document_number,
                            document_key,
                            get_field(c500_fields, 11),
                        )
                        expanded_items = expand_configured_icms_rule_items(selected_rule_profile, runtime_rules, parsed_item)
                        for expanded_item in expanded_items:
                            direct_c590_items.append(apply_sped_icms_consistency_rules(expanded_item, selected_rule_profile, runtime_rules))
                        continue
                    non_c590_lines.append(line)

                if not direct_c590_items:
                    output_lines.extend(block_lines)
                    continue

                direct_c590_items = rebalance_bella_citta_1252_1253(direct_c590_items, selected_rule_profile)
                direct_c590_items = consolidate_summary_items(direct_c590_items)
                update_c500_totals(c500_fields, direct_c590_items)
                if insert_position is None:
                    insert_position = len(non_c590_lines)
                    for position, line in enumerate(non_c590_lines):
                        if line.startswith("|C591|") or line.startswith("|C595|") or line.startswith("|C597|"):
                            insert_position = position
                            break

                rebuilt_lines = list(non_c590_lines)
                rebuilt_lines[insert_position:insert_position] = [build_c590_line(item) for item in direct_c590_items]
                rebuilt_block.extend(rebuilt_lines)
                output_lines.extend(rebuilt_block)
                continue

            output_lines.append(current_line)
            index += 1
            continue

        block_lines = [current_line]
        index += 1
        while index < len(source_lines):
            candidate = source_lines[index]
            candidate_fields = normalize_sped_line(candidate) if candidate.startswith("|") else []
            candidate_register = get_field(candidate_fields, 1)
            if not candidate.startswith("|"):
                break
            if candidate_register and candidate_register.startswith("C") and candidate_register not in c100_children:
                break
            if candidate.startswith("|D001|"):
                break
            block_lines.append(candidate)
            index += 1

        c100_fields = normalize_sped_line(block_lines[0])
        operation_type = "Entrada" if get_field(c100_fields, 2) == "0" else "Saida" if get_field(c100_fields, 2) == "1" else ""
        document_number = get_field(c100_fields, 8)
        document_key = normalize_document_key(get_field(c100_fields, 9))
        cod_sit = get_field(c100_fields, 6)
        document_identity = (operation_type, document_key or document_number)
        if document_identity not in affected_documents or cod_sit in {"02", "03", "04"}:
            output_lines.extend(block_lines)
            continue

        document_items = document_map.get(document_identity)
        document_details = document_detail_map.get(document_identity, {})

        if not document_items and (selected_rule_profile or runtime_rules):
            direct_c190_items: list[dict[str, object]] = []
            rebuilt_block = ["|".join(c100_fields)]
            non_c190_lines: list[str] = []
            insert_position = None

            for line in block_lines[1:]:
                if line.startswith("|C190|"):
                    if insert_position is None:
                        insert_position = len(non_c190_lines)
                    direct_c190_items.append(
                        apply_sped_icms_consistency_rules(
                            parse_c190_item_from_line(
                                line,
                                operation_type,
                                document_number,
                                document_key,
                                get_field(c100_fields, 10),
                            ),
                            selected_rule_profile,
                            runtime_rules,
                        )
                    )
                    continue
                non_c190_lines.append(line)

            if not direct_c190_items:
                output_lines.extend(block_lines)
                continue

            direct_c190_items = consolidate_c190_items(direct_c190_items)
            update_c100_totals(c100_fields, direct_c190_items)
            rebuilt_block = ["|".join(c100_fields)]
            if insert_position is None:
                insert_position = len(non_c190_lines)
                for position, line in enumerate(non_c190_lines):
                    if line.startswith("|C191|") or line.startswith("|C195|") or line.startswith("|C197|") or line.startswith("|C500|"):
                        insert_position = position
                        break

            rebuilt_lines = list(non_c190_lines)
            rebuilt_lines[insert_position:insert_position] = [build_c190_line(item) for item in direct_c190_items]
            rebuilt_block.extend(rebuilt_lines)
            output_lines.extend(rebuilt_block)
            continue

        if not document_items:
            output_lines.extend(block_lines)
            continue

        update_c100_totals(c100_fields, document_items)
        rebuilt_block = ["|".join(c100_fields)]
        non_c190_lines: list[str] = []
        insert_position = None

        next_item_number = 1
        for position, line in enumerate(block_lines[1:]):
            if line.startswith("|C170|"):
                c170_fields = normalize_sped_line(line)
                item_number = get_field(c170_fields, 2)
                details = document_details.get(item_number, [])
                if details:
                    for detail in details:
                        non_c190_lines.append(apply_override_to_c170_line(line, detail, str(next_item_number)))
                        next_item_number += 1
                else:
                    non_c190_lines.append(line)
                    next_item_number += 1
                continue
            if line.startswith("|C190|"):
                if insert_position is None:
                    insert_position = len(non_c190_lines)
                continue
            non_c190_lines.append(line)

        if insert_position is None:
            insert_position = len(non_c190_lines)
            for position, line in enumerate(non_c190_lines):
                if line.startswith("|C191|") or line.startswith("|C195|") or line.startswith("|C197|") or line.startswith("|C500|"):
                    insert_position = position
                    break

        rebuilt_lines = list(non_c190_lines)
        c190_lines = [build_c190_line(item) for item in document_items]
        rebuilt_lines[insert_position:insert_position] = c190_lines
        rebuilt_block.extend(rebuilt_lines)
        output_lines.extend(rebuilt_block)

    return source_lines, output_lines


def write_adjusted_sped(
    output_path: Path,
    sped_path: Path,
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
    cst_values: set[str],
    cfop_values: set[str],
    icms_rate: Decimal | None,
    new_cst: str,
    new_cfop: str,
    zero_icms_cfops: set[str],
    operation_scope: str,
    selected_rule_profile: str = "",
    xml_source_paths: list[Path] | None = None,
    xml_ncm_filters: set[str] | None = None,
    runtime_rules: list[dict[str, object]] | None = None,
) -> list[dict[str, object]]:
    # Ponto central da geracao do SPED ajustado:
    # filtra, recalcula, regrava o arquivo e produz o log de auditoria.
    scoped_detailed_sales = filter_details_by_operation_scope(detailed_sales, operation_scope)
    xml_ncm_filters = xml_ncm_filters or set()
    xml_fiscal_adjustment_rows: list[dict[str, object]] = []
    xml_st_adjustment_rows: list[dict[str, object]] = []
    missing_xml_rows: list[dict[str, object]] = []
    filtered_detailed_sales = (
        [item for item in scoped_detailed_sales if has_configured_icms_rule(selected_rule_profile, runtime_rules, item)]
        if (selected_rule_profile or runtime_rules)
        else filter_detailed_sales(scoped_detailed_sales, descriptions, cst_values, cfop_values)
    )
    affected_documents = (
        scan_default_filial_rule_documents(sped_path, operation_scope, selected_rule_profile, runtime_rules)
        |
        {
            (
                str(item["operation_type"]),
                normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
            )
            for item in filtered_detailed_sales
        }
        if (selected_rule_profile or runtime_rules)
        else {
            (
                str(item["operation_type"]),
                normalize_document_key(str(item["document_key"])) or str(item["document_number"]),
            )
            for item in filtered_detailed_sales
        }
    )
    sped_documents = scan_sped_c100_documents(sped_path, operation_scope)
    source_detailed_sales = prefer_excel_details_for_sped(scoped_detailed_sales, affected_documents)
    rebuilt_detailed_sales = rebuild_detailed_sales_with_override(
        source_detailed_sales,
        descriptions,
        cst_values,
        cfop_values,
        icms_rate,
        new_cst,
        new_cfop,
        zero_icms_cfops,
        selected_rule_profile,
        runtime_rules,
    )
    if xml_source_paths and xml_ncm_filters:
        rebuilt_detailed_sales, xml_fiscal_affected_documents, xml_fiscal_adjustment_rows = apply_xml_fiscal_adjustments_to_details(
            rebuilt_detailed_sales,
            sped_documents,
            xml_source_paths,
            xml_ncm_filters,
        )
        affected_documents |= xml_fiscal_affected_documents
    if xml_source_paths:
        rebuilt_detailed_sales, xml_affected_documents, xml_st_adjustment_rows, missing_xml_rows = apply_xml_st_adjustments_to_details(
            rebuilt_detailed_sales,
            xml_source_paths,
        )
        affected_documents |= xml_affected_documents
    rebuilt_detailed_sales = rebalance_bella_citta_1252_1253(rebuilt_detailed_sales, selected_rule_profile)
    original_lines, adjusted_lines = generate_adjusted_sped_lines(
        sped_path,
        rebuilt_detailed_sales,
        affected_documents,
        selected_rule_profile,
        runtime_rules,
    )
    normalize_sped_warning_fields(adjusted_lines)
    recalculate_e110(adjusted_lines)
    recalculate_e210(adjusted_lines)
    adjusted_lines = recalculate_sped_summaries(adjusted_lines)
    recalculate_e110(adjusted_lines)
    recalculate_e210(adjusted_lines)
    output_path.write_text("\n".join(adjusted_lines) + "\n", encoding="latin-1")
    write_adjustment_log(
        output_path,
        original_lines,
        adjusted_lines,
        affected_documents,
        build_rateio_log_rows(rebuilt_detailed_sales),
    )
    write_xml_st_adjustment_log(output_path, xml_st_adjustment_rows)
    if xml_fiscal_adjustment_rows:
        write_simple_excel_workbook(
            output_path.with_name(f"{output_path.stem}_log_xml_fiscal.xlsx"),
            [
                (
                    "Ajustes XML Fiscal",
                    [
                        "Tipo",
                        "Documento",
                        "Chave NFe",
                        "Item",
                        "Codigo",
                        "NCM Original",
                        "NCM XML",
                        "CFOP Original",
                        "CFOP XML",
                        "CST Original",
                        "CST XML",
                        "Valor Item Original",
                        "Valor Item XML",
                        "Base ICMS Original",
                        "Base ICMS XML",
                        "Aliquota Original",
                        "Aliquota XML",
                        "Valor ICMS Original",
                        "Valor ICMS XML",
                        "Base ICMS ST Original",
                        "Base ICMS ST XML",
                        "Valor ICMS ST Original",
                        "Valor ICMS ST XML",
                        "Valor IPI Original",
                        "Valor IPI XML",
                    ],
                    [
                        [
                            row["operation_type"],
                            row["document_number"],
                            row["document_key"],
                            row["item_number"],
                            row["code"],
                            row["ncm_original"],
                            row["ncm_xml"],
                            row["cfop_original"],
                            row["cfop_xml"],
                            row["cst_original"],
                            row["cst_xml"],
                            row["vl_item_original"],
                            row["vl_item_xml"],
                            row["vl_bc_original"],
                            row["vl_bc_xml"],
                            row["aliq_original"],
                            row["aliq_xml"],
                            row["vl_icms_original"],
                            row["vl_icms_xml"],
                            row["vl_bc_st_original"],
                            row["vl_bc_st_xml"],
                            row["vl_icms_st_original"],
                            row["vl_icms_st_xml"],
                            row["vl_ipi_original"],
                            row["vl_ipi_xml"],
                        ]
                        for row in xml_fiscal_adjustment_rows
                    ],
                )
            ],
        )
    missing_xml_output = output_path.with_name(f"{output_path.stem}_st_1401_1403_sem_xml.xlsx")
    if missing_xml_rows:
        write_missing_xml_st_excel(missing_xml_output, missing_xml_rows)
    elif missing_xml_output.exists():
        missing_xml_output.unlink()
    return rebuilt_detailed_sales


def recalculate_sped_summaries(lines: list[str]) -> list[str]:
    register_counts: dict[str, int] = defaultdict(int)
    for line in lines:
        if not line.startswith("|"):
            continue
        register = get_field(normalize_sped_line(line), 1)
        if register:
            register_counts[register] += 1

    c001_index = next((index for index, line in enumerate(lines) if line.startswith("|C001|")), None)
    c990_index = next((index for index, line in enumerate(lines) if line.startswith("|C990|")), None)
    if c001_index is not None and c990_index is not None and c990_index >= c001_index:
        c990_fields = normalize_sped_line(lines[c990_index])
        set_field(c990_fields, 2, str(c990_index - c001_index + 1))
        lines[c990_index] = "|".join(c990_fields)

    register_counts = defaultdict(int)
    for line in lines:
        if not line.startswith("|"):
            continue
        register = get_field(normalize_sped_line(line), 1)
        if register:
            register_counts[register] += 1

    for index, line in enumerate(lines):
        if not line.startswith("|9900|"):
            continue
        fields = normalize_sped_line(line)
        target_register = get_field(fields, 2)
        set_field(fields, 3, str(register_counts.get(target_register, 0)))
        lines[index] = "|".join(fields)

    b9001_index = next((index for index, line in enumerate(lines) if line.startswith("|9001|")), None)
    b9990_index = next((index for index, line in enumerate(lines) if line.startswith("|9990|")), None)
    line_9999_index = next((index for index, line in enumerate(lines) if line.startswith("|9999|")), None)
    if b9001_index is not None and b9990_index is not None and line_9999_index is not None and line_9999_index >= b9001_index:
        fields = normalize_sped_line(lines[b9990_index])
        set_field(fields, 2, str(line_9999_index - b9001_index + 1))
        lines[b9990_index] = "|".join(fields)

    if line_9999_index is not None:
        fields = normalize_sped_line(lines[line_9999_index])
        set_field(fields, 2, str(len(lines)))
        lines[line_9999_index] = "|".join(fields)

    return lines


def write_excel(
    output_path: Path,
    product_rows: list[ProductRecord],
    sales_rows: list[dict[str, object]],
    detailed_sales: list[dict[str, object]],
    filtered_rows: list[dict[str, object]],
    c190_rows: list[dict[str, object]],
    c190_product_rows: list[dict[str, object]],
    recalculated_c190_product_rows: list[dict[str, object]],
) -> None:
    # Gera o arquivo Excel final com as visoes de cadastro, filtro, C190 original,
    # C190 por produto e C190 recalculado.
    product_totals: dict[str, dict[str, Decimal]] = defaultdict(lambda: {"base_icms": Decimal("0"), "icms_value": Decimal("0")})
    for item in sales_rows:
        code = str(item["code"])
        product_totals[code]["base_icms"] += item["base_icms"]
        product_totals[code]["icms_value"] += item["icms_value"]

    cadastro_headers = ["Codigo Produto", "Descricao", "NCM", "CST ICMS", "Aliquota ICMS", "Base Calculo ICMS", "Valor ICMS"]
    cadastro_data = [
        [
            product.code,
            product.description,
            product.ncm,
            product.cst_icms,
            product.icms_rate if product.icms_rate is not None else "",
            product_totals[product.code]["base_icms"],
            product_totals[product.code]["icms_value"],
        ]
        for product in product_rows
    ]

    vendas_headers = [
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "Aliquota ICMS",
        "Quantidade Total",
        "Valor Venda",
        "Base Calculo ICMS",
        "Valor ICMS",
    ]
    vendas_data = [
        [
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
        ]
        for item in sales_rows
    ]
    filtro_headers = [
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "Quantidade Total",
        "Base Calculo ICMS",
        "Aliquota ICMS",
        "Valor ICMS",
    ]
    filtro_data = [
        [
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["quantity"],
            item["base_icms"],
            item["icms_rate"],
            item["icms_value"],
        ]
        for item in filtered_rows
    ]
    c190_grouped_products: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    for item in c190_product_rows:
        key = (
            str(item["operation_type"]),
            str(item["document_number"]),
            str(item["document_key"]),
            str(item["cst_icms"]),
            str(item["cfop"]),
            item["icms_rate"],
        )
        bucket = c190_grouped_products.setdefault(
            key,
            {
                "codes": [],
                "descriptions": [],
                "ncms": [],
                "quantity": Decimal("0"),
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )
        if item["code"] and item["code"] not in bucket["codes"]:
            bucket["codes"].append(item["code"])
        if item["description"] and item["description"] not in bucket["descriptions"]:
            bucket["descriptions"].append(item["description"])
        if item["ncm"] and item["ncm"] not in bucket["ncms"]:
            bucket["ncms"].append(item["ncm"])
        bucket["quantity"] += item["quantity"]
        bucket["sale_value"] += item["sale_value"]
        bucket["base_icms"] += item["base_icms"]
        bucket["icms_value"] += item["icms_value"]

    c190_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Codigo Produto",
        "Produto",
        "NCM",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS",
        "Quantidade Produto",
        "Valor Produto",
        "Base ICMS Produto",
        "Valor ICMS Produto",
        "Valor Operacao C190",
        "Base Calculo ICMS C190",
        "Valor ICMS C190",
        "Base Calculo ICMS ST C190",
        "Valor ICMS ST C190",
        "Valor Reducao Base C190",
        "Valor IPI C190",
    ]
    c190_data = []
    for item in c190_rows:
        key = (
            str(item["operation_type"]),
            str(item["document_number"]),
            str(item["document_key"]),
            str(item["cst_icms"]),
            str(item["cfop"]),
            item["icms_rate"],
        )
        product_match = c190_grouped_products.get(
            key,
            {
                "codes": [],
                "descriptions": [],
                "ncms": [],
                "quantity": Decimal("0"),
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )
        has_products = bool(product_match["descriptions"])
        c190_data.append(
            [
                item["operation_type"],
                item["document_number"],
                item["document_key"],
                item["document_date"],
                " | ".join(product_match["codes"]) if product_match["codes"] else "",
                " | ".join(product_match["descriptions"]) if has_products else "Nao foi possivel identificar os produtos pelo SPED",
                " | ".join(product_match["ncms"]) if has_products else "",
                item["cst_icms"],
                item["cfop"],
                item["icms_rate"],
                product_match["quantity"],
                product_match["sale_value"],
                product_match["base_icms"],
                product_match["icms_value"],
                item["total_operation_value"],
                item["base_icms"],
                item["icms_value"],
                item["base_icms_st"],
                item["icms_st_value"],
                item["reduction_value"],
                item["ipi_value"],
            ]
        )
    c190_product_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS",
        "Quantidade Produto",
        "Valor Produto",
        "Base ICMS Produto",
        "Valor ICMS Produto",
        "Valor Operacao C190",
        "Base ICMS C190",
        "Valor ICMS C190",
        "Base ICMS ST C190",
        "Valor ICMS ST C190",
        "Reducao Base C190",
        "Valor IPI C190",
    ]
    c190_product_data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["cfop"],
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
            item["c190_total_operation_value"],
            item["c190_base_icms"],
            item["c190_icms_value"],
            item["c190_base_icms_st"],
            item["c190_icms_st_value"],
            item["c190_reduction_value"],
            item["c190_ipi_value"],
        ]
        for item in c190_product_rows
    ]
    c190_saida_data = [row for row in c190_product_data if row[0] == "Saida"]
    c190_sped_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Numero Cupom",
        "Chave NFe",
        "Data Documento",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS",
        "Valor Operacao",
        "Base ICMS",
        "Valor ICMS",
        "Base ICMS ST",
        "Valor ICMS ST",
        "Valor Reducao Base",
        "Valor IPI",
    ]
    c190_sped_data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["cst_icms"],
            item["cfop"],
            item["icms_rate"],
            item["total_operation_value"],
            item["base_icms"],
            item["icms_value"],
            item["base_icms_st"],
            item["icms_st_value"],
            item["reduction_value"],
            item["ipi_value"],
        ]
        for item in c190_rows
    ]
    c190_product_recalc_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS Recalculada",
        "Quantidade Produto",
        "Valor Produto",
        "Base ICMS Produto",
        "Valor ICMS Produto Recalculado",
        "Valor Operacao C190",
        "Base ICMS C190",
        "Valor ICMS C190 Recalculado",
        "Base ICMS ST C190",
        "Valor ICMS ST C190",
        "Reducao Base C190",
        "Valor IPI C190",
    ]
    c190_product_recalc_data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["cfop"],
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
            item["c190_total_operation_value"],
            item["c190_base_icms"],
            item["c190_icms_value"],
            item["c190_base_icms_st"],
            item["c190_icms_st_value"],
            item["c190_reduction_value"],
            item["c190_ipi_value"],
        ]
        for item in recalculated_c190_product_rows
    ]
    cst_061_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Item",
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS",
        "Quantidade",
        "Valor Operacao",
        "Base ICMS",
        "Valor ICMS",
    ]
    cst_061_data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["item_number"],
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["cfop"],
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
        ]
        for item in detailed_sales
        if str(item["operation_type"]).strip().lower() == "entrada"
        and normalize_cst_icms_for_sped(str(item["cst_icms"])) == "061"
    ]
    write_simple_excel_workbook(
        output_path,
        [
            ("Cadastro Produtos", cadastro_headers, cadastro_data),
            ("Resumo ICMS", vendas_headers, vendas_data),
            ("Filtro Descricoes", filtro_headers, filtro_data),
            ("Registro C190", c190_headers, c190_data),
            ("C190 por Produto", c190_product_headers, c190_product_data),
            ("C190 Saida", c190_product_headers, c190_saida_data),
            ("C190 SPED Filtrado", c190_sped_headers, c190_sped_data),
            ("C190 Prod Recalc", c190_product_recalc_headers, c190_product_recalc_data),
            ("Entradas CST 061", cst_061_headers, cst_061_data),
        ],
    )


def write_cst_061_excel(output_path: Path, detailed_sales: list[dict[str, object]]) -> None:
    special_codes_061 = {"0078957238", "0078957271"}
    cst_061_headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Item",
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "CFOP",
        "CFOP Ajustado",
        "Aliquota ICMS",
        "Quantidade",
        "Valor Operacao",
        "Base ICMS",
        "Valor ICMS",
        "Valor ICMS Calc Especial",
    ]
    cst_061_data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["item_number"],
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["cfop"],
            (
                "1651"
                if str(item["cfop"]).strip() == "1401" and str(item["code"]).strip() in special_codes_061
                else item["cfop"]
            ),
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
            (
                (
                    item["quantity"]
                    * (
                        Decimal("1.47")
                        if str(item["code"]).strip() == "0078957271" and str(item["cfop"]).strip() == "1651"
                        else Decimal("1.12")
                    )
                ).quantize(Decimal("0.01"))
                if str(item["code"]).strip() in special_codes_061 and str(item["cfop"]).strip() in {"1401", "1651", "1653"}
                else ""
            ),
        ]
        for item in detailed_sales
        if str(item["operation_type"]).strip().lower() == "entrada"
        and normalize_cst_icms_for_sped(str(item["cst_icms"])) == "061"
    ]

    write_simple_excel_workbook(output_path, [("Entradas CST 061", cst_061_headers, cst_061_data)])


def write_cfop_1252_1253_excel(output_path: Path, detailed_sales: list[dict[str, object]]) -> None:
    headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Item",
        "Codigo Produto",
        "Descricao",
        "NCM",
        "CST ICMS",
        "CFOP",
        "Aliquota ICMS",
        "Quantidade",
        "Valor Operacao",
        "Base ICMS",
        "Valor ICMS",
    ]
    data = [
        [
            item["operation_type"],
            item["document_number"],
            item["document_key"],
            item["document_date"],
            item["item_number"],
            item["code"],
            item["description"],
            item["ncm"],
            item["cst_icms"],
            item["cfop"],
            item["icms_rate"],
            item["quantity"],
            item["sale_value"],
            item["base_icms"],
            item["icms_value"],
        ]
        for item in detailed_sales
        if str(item["operation_type"]).strip().lower() == "entrada"
        and str(item["cfop"]).strip() in {"1252", "1253"}
    ]

    write_simple_excel_workbook(output_path, [("Entradas 1252 1253", headers, data)])


def write_missing_xml_st_excel(output_path: Path, missing_rows: list[dict[str, object]]) -> None:
    headers = [
        "Tipo Operacao",
        "Numero Documento",
        "Chave NFe",
        "Data Documento",
        "Item",
        "Codigo Produto",
        "Descricao",
        "CFOP Original",
        "CFOP Ajustado",
        "Motivo",
    ]
    data = [
        [
            item.get("operation_type", ""),
            item.get("document_number", ""),
            item.get("document_key", ""),
            item.get("document_date", ""),
            item.get("item_number", ""),
            item.get("code", ""),
            item.get("description", ""),
            item.get("cfop_original", ""),
            item.get("cfop_ajustado", ""),
            item.get("missing_reason", ""),
        ]
        for item in missing_rows
    ]

    write_simple_excel_workbook(output_path, [("ST sem XML", headers, data)])


def format_rule_decimal(value: object) -> str:
    if not isinstance(value, Decimal):
        return str(value or "").strip()
    normalized = value.quantize(Decimal("0.01"))
    text = format(normalized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", ",")


def describe_operation_type(value: object) -> str:
    normalized = normalize_operation_type(str(value or "")).strip()
    return normalized or "Qualquer"


def describe_codes(value: object) -> str:
    if not isinstance(value, set):
        return ""
    codes = sorted(str(code).strip() for code in value if str(code).strip())
    return ", ".join(codes)


def describe_rule_conditions(rule: dict[str, object]) -> str:
    parts: list[str] = []
    operation_type = str(rule.get("operation_type", "")).strip()
    if operation_type:
        parts.append(f"Tipo {describe_operation_type(operation_type)}")
    cst_icms = str(rule.get("cst_icms", "")).strip()
    if cst_icms:
        parts.append(f"CST {cst_icms}")
    cfop = str(rule.get("cfop", "")).strip()
    if cfop:
        parts.append(f"CFOP {cfop}")
    match_rate = rule.get("match_rate")
    if isinstance(match_rate, Decimal):
        parts.append(f"Aliquota {format_rule_decimal(match_rate)}")
    match_rates = rule.get("match_rates")
    if isinstance(match_rates, set) and match_rates:
        formatted = ", ".join(sorted(format_rule_decimal(rate) for rate in match_rates))
        parts.append(f"Aliquotas {formatted}")
    match_sale_value = rule.get("match_sale_value")
    if isinstance(match_sale_value, Decimal):
        parts.append(f"Valor da operacao {format_rule_decimal(match_sale_value)}")
    match_base_icms = rule.get("match_base_icms")
    if isinstance(match_base_icms, Decimal):
        parts.append(f"Base ICMS {format_rule_decimal(match_base_icms)}")
    match_codes = describe_codes(rule.get("match_codes"))
    if match_codes:
        parts.append(f"Codigos {match_codes}")
    return " | ".join(parts) if parts else "Sem filtro especifico"


def describe_rule_actions(rule: dict[str, object]) -> str:
    parts: list[str] = []
    new_cst = str(rule.get("new_cst", "")).strip()
    if new_cst:
        parts.append(f"alterar CST para {new_cst}")
    new_cfop = str(rule.get("new_cfop", "")).strip()
    if new_cfop:
        parts.append(f"alterar CFOP para {new_cfop}")
    force_rate = rule.get("force_rate")
    if isinstance(force_rate, Decimal):
        parts.append(f"definir aliquota ICMS em {format_rule_decimal(force_rate)}")
    set_base_icms = rule.get("set_base_icms")
    if isinstance(set_base_icms, Decimal):
        parts.append(f"definir base ICMS em {format_rule_decimal(set_base_icms)}")
    set_icms_value = rule.get("set_icms_value")
    if isinstance(set_icms_value, Decimal):
        parts.append(f"definir valor do ICMS em {format_rule_decimal(set_icms_value)}")
    if rule.get("use_sale_value_as_base"):
        parts.append("usar o valor da operacao como base do ICMS")
    set_sale_value = rule.get("set_sale_value")
    if isinstance(set_sale_value, Decimal):
        parts.append(f"definir valor da operacao em {format_rule_decimal(set_sale_value)}")
    sale_value_factor = rule.get("sale_value_factor")
    if isinstance(sale_value_factor, Decimal):
        parts.append(f"aplicar fator {format_rule_decimal(sale_value_factor)} ao valor da operacao")
    icms_per_quantity = rule.get("icms_per_quantity")
    if isinstance(icms_per_quantity, Decimal):
        parts.append(f"calcular ICMS por quantidade em {format_rule_decimal(icms_per_quantity)} por unidade")
    if rule.get("recalculate_icms_value"):
        parts.append("recalcular o valor do ICMS")
    if rule.get("zero_icms"):
        parts.append("zerar base, aliquota e valor do ICMS")
    if rule.get("preserve_icms_value"):
        parts.append("preservar o valor final do ICMS apos o calculo especifico")
    return "; ".join(parts) if parts else "Sem acao de alteracao"


def describe_rule_notes(rule: dict[str, object]) -> str:
    notes: list[str] = ["Nao altera o valor do IPI."]
    split_into = rule.get("split_into")
    if isinstance(split_into, list) and split_into:
        notes.append(f"Desdobra o item em {len(split_into)} lancamento(s).")
    return " ".join(notes)


def build_rule_report_entries(rule: dict[str, object], prefix: str = "") -> list[str]:
    summary_line = f"{prefix}{describe_rule_conditions(rule)} -> {describe_rule_actions(rule)}. {describe_rule_notes(rule)}"
    lines = [summary_line]
    split_into = rule.get("split_into")
    if isinstance(split_into, list) and split_into:
        for split_index, split_rule in enumerate(split_into, start=1):
            if isinstance(split_rule, dict):
                lines.extend(build_rule_report_entries(split_rule, prefix=f"{prefix}Desdobramento {split_index}: "))
    return lines


def make_word_paragraph(text: str, style: str | None = None) -> str:
    escaped = html.escape(str(text or ""))
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return (
        "<w:p>"
        f"{style_xml}"
        f"<w:r><w:t xml:space=\"preserve\">{escaped}</w:t></w:r>"
        "</w:p>"
    )


def write_rules_report_docx(
    output_path: Path,
    report_title: str,
    sections: list[dict[str, object]],
) -> None:
    body_parts: list[str] = [
        make_word_paragraph(report_title, "Heading1"),
        make_word_paragraph(f"Gerado em: {dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"),
        make_word_paragraph("Observacao geral: as regras documentadas neste relatorio nao alteram o valor do IPI."),
    ]

    for section in sections:
        title = str(section.get("title", "")).strip()
        lines = [str(line).strip() for line in section.get("lines", []) if str(line).strip()]
        if not title or not lines:
            continue
        body_parts.append(make_word_paragraph(""))
        body_parts.append(make_word_paragraph(title, "Heading2"))
        for line in lines:
            body_parts.append(make_word_paragraph(line))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14">'
        "<w:body>"
        + "".join(body_parts)
        + (
            "<w:sectPr>"
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="708" w:footer="708" w:gutter="0"/>'
            "</w:sectPr>"
            "</w:body></w:document>"
        )
    )

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
</w:styles>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Word</Application>
</Properties>
"""
    created_at = dt.datetime.now().replace(microsecond=0).isoformat()
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{html.escape(report_title)}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>
</cp:coreProperties>
"""

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("docProps/app.xml", app_xml)
        docx.writestr("docProps/core.xml", core_xml)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/_rels/document.xml.rels", document_rels)
