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
from app.services.adjusted_sped import (
    apply_icms_rate_override,
    apply_override_to_c170_line,
    build_c190_line,
    build_c590_line,
    build_document_c190_map,
    build_document_detail_map,
    build_rateio_log_rows,
    calculate_e110_totals,
    calculate_e210_e220_out_credit,
    calculate_e210_out_credit_from_c190,
    consolidate_c190_items,
    consolidate_summary_items,
    distribute_icms_with_caps,
    filter_detailed_sales,
    filter_detailed_sales_by_descriptions,
    filter_sales,
    force_icms_total_for_target_group,
    generate_adjusted_sped_lines,
    is_e210_other_credit_cfop,
    lines_differ_meaningfully,
    normalize_field_for_log_compare,
    normalize_sped_warning_fields,
    parse_c190_item_from_line,
    parse_c590_item_from_line,
    prefer_excel_details_for_sped,
    rebalance_bella_citta_1252_1253,
    rebalance_document_rule_icms_values,
    rebuild_detailed_sales_with_override,
    recalculate_e110,
    recalculate_e116,
    recalculate_e210,
    recalculate_sped_summaries,
    scan_default_filial_rule_documents,
    update_c100_totals,
    update_c500_totals,
    write_adjusted_sped,
    write_cfop_1252_1253_excel,
    write_cst_061_excel,
    write_excel,
    write_missing_xml_st_excel,
    write_adjustment_log,
    write_xml_st_adjustment_log,
)
from app.services.analysis_reports import (
    build_contrib_operation_launch_details_map,
    build_contrib_operation_summary_rows,
    build_contrib_product_monthly_linear_dataset,
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_credit_diagnostic_period_detail_comparison_dataset,
    build_credit_diagnostic_product_sheet_payload,
    build_entry_exit_analysis_rows,
    build_entry_exit_footer_rows,
    build_entry_period_comparison_rows,
    build_monthly_aliquota_divergence_rows,
    build_multi_sped_entry_analysis,
    build_product_monthly_linear_dataset,
    build_sale_period_comparison_rows,
    build_summary_operation_totals_by_period,
    merge_credit_diagnostic_detail_rows,
    period_label_sort_key,
    summarize_entry_analysis,
    summarize_sale_analysis,
    write_entry_exit_analysis_excel,
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
from app.services.operation_summary import (
    attach_c190_totals_to_products,
    build_c190_rows_from_details,
    build_filtered_apuracao_rows,
    build_operation_launch_details_map,
    build_operation_summary_rows,
    build_operation_summary_rows_from_c190,
    build_reduction_launch_rows,
    build_reduction_rows_from_c190,
    build_sales_rows_from_details,
    build_synthetic_launch_details_from_c190,
    combine_imported_data,
    describe_operation_base_difference,
    enrich_details_with_note_snapshots,
    filter_c190_rows,
    get_launch_total_operation_value,
    get_note_item_display_sale_value,
    get_operation_base_difference,
    read_detailed_product_excel,
)
from app.repositories.mysql_cadastro import MysqlCadastroRepository


















































































































































