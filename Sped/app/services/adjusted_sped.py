from __future__ import annotations

import csv
import datetime as dt
import html
import re
import shutil
import tempfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from app.config import COMPARE_MARK_CHECKED, COMPARE_MARK_UNCHECKED
from app.exporters.excel_base import (
    EXCEL_STYLE_BODY_BLUE,
    EXCEL_STYLE_BODY_YELLOW,
    EXCEL_STYLE_CURRENCY,
    EXCEL_STYLE_CURRENCY_BLUE,
    EXCEL_STYLE_CURRENCY_YELLOW,
    EXCEL_STYLE_DEFAULT,
    EXCEL_STYLE_HEADER,
    EXCEL_STYLE_HEADER_BLUE,
    EXCEL_STYLE_HEADER_YELLOW,
    EXCEL_STYLE_NUMBER,
    EXCEL_STYLE_NUMBER_BLUE,
    EXCEL_STYLE_NUMBER_YELLOW,
    EXCEL_STYLE_PERCENT,
    EXCEL_STYLE_PERCENT_BLUE,
    EXCEL_STYLE_PERCENT_YELLOW,
    build_cell,
    build_excel_styles_xml,
    build_sheet_rows_with_metadata,
    build_sheet_xml,
    decimal_from_excel_value,
    detect_excel_column_kind,
    excel_column_name,
    is_excel_numeric_value,
    normalize_excel_header_text,
    sanitize_excel_sheet_name,
    should_total_excel_column,
    xml_escape,
)
from app.exporters.workbook_exporter import write_simple_csv_file, write_simple_excel_workbook
from app.parsers.sped_parser import get_field, normalize_cfop, normalize_document_key, normalize_sped_line, parse_decimal, parse_rate, read_sped_plain_lines
from app.services.analysis_utils import filter_details_by_operation_scope
from app.services.analysis_reports import build_nfce_note_snapshot
from app.services.operation_summary import build_c190_rows_from_details, build_operation_summary_rows_from_c190, get_launch_total_operation_value
from app.services.runtime_rules import (
    apply_sped_icms_consistency_rules,
    build_rule_signature,
    expand_configured_icms_rule_items,
    get_configured_icms_rule,
    has_configured_icms_rule,
)
from app.services.tax_rules import format_decimal_sped, normalize_cst_icms_for_sped, normalize_operation_type, normalize_text
from app.services.xml_reconciliation import (
    apply_xml_fiscal_adjustments_to_details,
    apply_xml_st_adjustments_to_details,
    compose_xml_icms_cst_for_sped,
    read_c100_c190_fallback_rows,
    scan_sped_c100_documents,
    set_field,
)


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

def filter_detailed_sales_by_descriptions(
    detailed_sales: list[dict[str, object]],
    descriptions: list[str],
) -> list[dict[str, object]]:
    return filter_detailed_sales(detailed_sales, descriptions, set(), set())

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
