from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.models import ProductRecord
from app.parsers.sped_parser import (
    first_non_empty,
    get_field,
    normalize_document_key,
    normalize_sped_line,
    parse_decimal,
    parse_rate,
)
from app.services.tax_rules import merge_product_cst, merge_product_rate


def extract_tax_id_from_document_key(document_key: object) -> str:
    normalized_key = normalize_document_key(str(document_key or ""))
    if len(normalized_key) == 44 and normalized_key.isdigit():
        return normalized_key[6:20]
    return ""


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


def read_sped_0200_products(
    file_path: Path,
) -> tuple[str, str, str, str, list[dict], dict[str, list[dict]]]:
    """
    Le um arquivo SPED EFD fiscal em uma unica passagem e extrai:
    - CNPJ e nome da empresa (registro 0000)
    - Periodo inicio e fim (registro 0000)
    - Lista de produtos do registro 0200
    - Mapa produto->fornecedores de entrada (via 0150/C100/C170)

    Retorna: (cnpj, nome, periodo_inicio, periodo_fim, produtos, entry_suppliers_by_product)
      entry_suppliers_by_product: {cod_produto: [{"nome": ..., "cnpj": ...}, ...]}
      Cada lista contem os fornecedores distintos que deram entrada nesse produto.
    """
    company_cnpj = ""
    company_name = ""
    periodo_inicio = ""
    periodo_fim = ""
    products: list[dict] = []
    participants: dict[str, dict] = {}
    entry_suppliers_by_product: dict[str, list[dict]] = {}
    seen_supplier_per_product: dict[str, set[str]] = {}

    current_is_entry = False
    current_participant_code = ""

    with file_path.open("r", encoding="latin-1", errors="replace") as f:
        for raw_line in f:
            if not raw_line.startswith("|"):
                continue
            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)

            if register == "0000":
                periodo_inicio = get_field(fields, 4)
                periodo_fim = get_field(fields, 5)
                company_name = get_field(fields, 6)
                company_cnpj = "".join(c for c in get_field(fields, 7) if c.isdigit())
                continue

            if register == "0150":
                pcode = get_field(fields, 2)
                participants[pcode] = {
                    "nome": get_field(fields, 3),
                    "cnpj": "".join(c for c in first_non_empty(get_field(fields, 5), get_field(fields, 4)) if c.isdigit()),
                }
                continue

            if register == "0200":
                products.append(
                    {
                        "codigo": get_field(fields, 2),
                        "descricao": get_field(fields, 3),
                        "ncm": get_field(fields, 8),
                        "cst_icms": get_field(fields, 11) or "",
                        "aliquota_icms": str(parse_rate(get_field(fields, 12)) or ""),
                        "cest": get_field(fields, 13),
                    }
                )
                continue

            if register == "C100":
                current_is_entry = get_field(fields, 2) == "0"
                current_participant_code = get_field(fields, 4)
                continue

            if register == "C170" and current_is_entry:
                prod_code = get_field(fields, 3)
                if not prod_code:
                    continue
                participant = participants.get(current_participant_code, {})
                sup_cnpj = participant.get("cnpj", "")
                sup_nome = participant.get("nome", current_participant_code)
                seen_key = f"{sup_cnpj}|{sup_nome}"
                if seen_key not in seen_supplier_per_product.setdefault(prod_code, set()):
                    seen_supplier_per_product[prod_code].add(seen_key)
                    entry_suppliers_by_product.setdefault(prod_code, []).append(
                        {"nome": sup_nome, "cnpj": sup_cnpj}
                    )

    return company_cnpj, company_name, periodo_inicio, periodo_fim, products, entry_suppliers_by_product
