from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.models import ProductRecord
from app.parsers.excel_parser import get_first_xlsx_sheet_name, read_xlsx_sheet
from app.parsers.sped_parser import first_non_empty, normalize_document_key, parse_decimal, parse_rate
from app.services.runtime_rules import apply_sped_icms_consistency_rules
from app.services.tax_rules import (
    has_icms_reduction,
    merge_product_cst,
    merge_product_rate,
    normalize_cst_icms_for_sped,
    normalize_header,
    normalize_operation_type,
    normalize_text,
    resolve_header,
)

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
