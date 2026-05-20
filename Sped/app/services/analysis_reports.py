from __future__ import annotations

import re
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.parsers.sped_fiscal_parser import read_sped_file
from app.parsers.sped_parser import normalize_document_key, parse_decimal, read_sped_summary_register_rows
from app.services.credit_diagnostics import (
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_credit_diagnostic_period_detail_comparison_dataset,
    build_credit_diagnostic_product_rows,
    build_credit_diagnostic_product_sheet_payload,
    classify_credit_base_reason,
    classify_debit_base_reason,
    classify_icms_diagnostic_reason,
    format_grouped_line_numbers,
    merge_credit_diagnostic_detail_rows,
)
from app.services.entry_exit_reports import (
    build_entry_exit_analysis_rows,
    build_entry_exit_footer_rows,
    classify_entry_exit_category,
    classify_entry_exit_marker,
    sum_entry_exit_analysis_rows,
    write_entry_exit_analysis_excel,
)
from app.services.analysis_utils import calculate_abc_curve_labels, infer_sped_period_label, parse_sped_document_date, period_label_sort_key
from app.services.operation_summary import (
    build_c190_rows_from_details,
    build_operation_summary_rows_from_c190,
    get_launch_total_operation_value,
)
from app.services.product_import import collect_cest_values
from app.services.runtime_rules import apply_sped_icms_consistency_rules
from app.services.tax_rules import (
    compute_display_icms_rate,
    format_decimal_sped,
    has_icms_reduction,
    normalize_cst_icms_for_sped,
    normalize_operation_type,
    normalize_text,
)
from app.services.xml_reconciliation import (
    build_xml_fiscal_identity_index,
    build_xml_fiscal_item_index,
    collect_xml_candidate_document_keys,
    compose_xml_icms_cst_for_sped,
    normalize_xml_rebuilt_items_with_fallback,
    read_c100_c190_fallback_rows,
    scan_sped_c100_documents,
)



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
