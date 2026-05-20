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
from app.exporters.rules_report_exporter import (
    format_rule_decimal,
    describe_operation_type,
    describe_codes,
    describe_rule_conditions,
    describe_rule_actions,
    describe_rule_notes,
    build_rule_report_entries,
    make_word_paragraph,
    write_rules_report_docx,
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
    ProductRecord,
)
from app.parsers.sped_fiscal_parser import read_sped_file
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
    read_xlsx_sheet,
    get_first_xlsx_sheet_name,
    read_filter_descriptions_file,
    parse_filter_values,
    has_active_item_filters,
)
from app.parsers.compare_sped_reader import extract_company_tax_id_from_sped
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
    compare_extract_icms,
    compare_extract_pis_cofins,
    compare_extract_ipi,
    parse_compare_nfse_file,
    extract_xml_cancellation_event,
    collect_compare_ignored_xml_rows,
)
from app.services.xml_reconciliation import (
    allocate_decimal_proportionally,
    apply_xml_fiscal_adjustments_to_details,
    apply_xml_st_adjustments_to_details,
    build_pis_cofins_period_comparison_rows,
    build_xml_fiscal_identity_index,
    build_xml_fiscal_item_index,
    build_xml_st_index,
    collect_xml_candidate_document_keys,
    compose_xml_icms_cst_for_sped,
    display_text,
    export_nfce_items_by_ncm,
    find_matching_xml_fiscal_item,
    load_nfce_xml_items_for_index,
    map_xml_cfop_to_entry_cfops,
    normalize_xml_rebuilt_items_with_fallback,
    parse_nfe_xml_st_items,
    read_c100_c190_fallback_rows,
    read_xml_document_key,
    rebuild_sped_contrib_items_from_xml,
    scan_sped_c100_documents,
    scan_sped_contrib_c100_documents,
    set_field,
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
from app.services.runtime_rules import (
    DEFAULT_ICMS_RULE_PROFILES,
    apply_configured_icms_rules,
    apply_default_icms_rule_actions,
    apply_default_icms_rules,
    apply_sped_icms_consistency_rules,
    build_rule_signature,
    decimal_rule_matches,
    expand_configured_icms_rule_items,
    expand_default_icms_rule_items,
    extract_tax_id_from_document_key,
    get_configured_icms_rule,
    get_default_icms_rule,
    get_first_matching_icms_rule,
    has_configured_icms_rule,
    has_default_icms_rule,
    normalize_runtime_rule_tax_id,
    parse_bool_flag,
    parse_replacement_value,
    parse_runtime_rule_lines,
    runtime_rule_summary,
)
from app.services.analysis_utils import (
    calculate_abc_curve_labels,
    filter_details_by_operation_scope,
    infer_sped_period_label,
    parse_sped_document_date,
)
from app.services.compare_workflows import (
    build_xml_cfop_summary_rows,
    build_xml_entry_credit_rows,
    compare_sped_with_sheet,
    compare_sped_with_xml_folder,
)
from app.services.compare_matching import (
    build_compare_sped_number_indexes,
    compare_decimal_value,
    compare_sped_c100_has_generated_nfse_service,
    describe_compare_xml_sped_value_difference,
    find_compare_sped_document_by_key_or_number,
)
from app.services.compare_operations import (
    build_compare_invoice_note_snapshot,
    classify_xml_item_operation,
    filter_xml_summary_rows_by_scope,
    infer_compare_invoice_operation,
    normalize_compare_operation_scope,
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
from app.services.path_selection import get_xml_worker_count
from app.repositories.mysql_cadastro import MysqlCadastroRepository



















































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


