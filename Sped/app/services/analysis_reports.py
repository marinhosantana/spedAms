from __future__ import annotations

import re
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.parsers.sped_fiscal_parser import read_sped_file
from app.parsers.sped_parser import normalize_document_key, parse_decimal, read_sped_summary_register_rows
from app.services.analysis_utils import calculate_abc_curve_labels, infer_sped_period_label, parse_sped_document_date
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
