from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.models import ProductRecord
from app.services.tax_rules import normalize_cst_icms_for_sped


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
