from __future__ import annotations

import re
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.parsers.sped_fiscal_parser import read_sped_file
from app.services.tax_rules import format_decimal_sped, normalize_cst_icms_for_sped, normalize_operation_type


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
