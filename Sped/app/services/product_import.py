from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET

from app.parsers.sped_parser import (
    first_non_empty,
    get_field,
    normalize_document_key,
    normalize_sped_line,
    parse_decimal,
    parse_rate,
)
from app.services.tax_rules import (
    VALID_PIS_COFINS_CST_AMBOS,
    VALID_PIS_COFINS_CST_ENTRADA,
    VALID_PIS_COFINS_CST_SAIDA,
    build_pis_cofins_side_values,
    normalize_operation_type,
    normalize_tax_code,
    normalize_text,
)


def get_xml_local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


def collect_xml_files(xml_sources: list[Path]) -> list[Path]:
    xml_files: list[Path] = []
    for xml_source in xml_sources:
        if not xml_source.exists():
            raise ValueError(f"O caminho de XML informado nao existe: {xml_source}")
        if xml_source.is_file():
            if xml_source.suffix.lower() == ".xml":
                xml_files.append(xml_source)
            continue
        xml_files.extend(sorted(path for path in xml_source.rglob("*.xml") if path.is_file()))
    return xml_files


def normalize_ncm(value: str) -> str:
    return "".join(char for char in str(value or "").strip() if char.isdigit())


def local_name_text(parent: ET.Element | None, *names: str) -> str:
    if parent is None:
        return ""
    target_names = set(names)
    for element in parent.iter():
        local_name = get_xml_local_name(element.tag)
        if local_name in target_names:
            text = (element.text or "").strip()
            if text:
                return text
    return ""


def ncm_matches_filters(ncm: str, ncm_filters: set[str]) -> bool:
    normalized_ncm = normalize_ncm(ncm)
    if not ncm_filters:
        return True
    for raw_filter in ncm_filters:
        text = str(raw_filter or "").strip()
        if not text:
            continue
        normalized_filter = normalize_ncm(text)
        if not normalized_filter:
            continue
        if text.endswith("*"):
            if normalized_ncm.startswith(normalized_filter):
                return True
            continue
        if normalized_ncm == normalized_filter:
            return True
    return False


def parse_nfce_xml_items(file_path: Path) -> list[dict[str, object]]:
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

    emit_cnpj = first_non_empty(
        root.findtext(".//nfe:emit/nfe:CNPJ", default="", namespaces=namespace),
        root.findtext(".//nfe:emit/nfe:CPF", default="", namespaces=namespace),
    )
    emit_name = root.findtext(".//nfe:emit/nfe:xNome", default="", namespaces=namespace).strip()
    document_number = root.findtext(".//nfe:ide/nfe:nNF", default="", namespaces=namespace).strip()
    series = root.findtext(".//nfe:ide/nfe:serie", default="", namespaces=namespace).strip()
    document_model = root.findtext(".//nfe:ide/nfe:mod", default="", namespaces=namespace).strip()
    emission_date = first_non_empty(
        root.findtext(".//nfe:ide/nfe:dhEmi", default="", namespaces=namespace),
        root.findtext(".//nfe:ide/nfe:dEmi", default="", namespaces=namespace),
    )

    items: list[dict[str, object]] = []
    for det in root.findall(".//nfe:det", namespace):
        prod = det.find("nfe:prod", namespace)
        imposto = det.find("nfe:imposto", namespace)
        if prod is None:
            continue

        icms_group = imposto.find("nfe:ICMS", namespace) if imposto is not None else None
        pis_group = imposto.find("nfe:PIS", namespace) if imposto is not None else None
        cofins_group = imposto.find("nfe:COFINS", namespace) if imposto is not None else None
        ipi_group = imposto.find("nfe:IPI", namespace) if imposto is not None else None

        ncm = prod.findtext("nfe:NCM", default="", namespaces=namespace).strip()
        valor_produto = parse_decimal(prod.findtext("nfe:vProd", default="", namespaces=namespace))
        valor_desconto = parse_decimal(prod.findtext("nfe:vDesc", default="", namespaces=namespace))
        valor_frete = parse_decimal(prod.findtext("nfe:vFrete", default="", namespaces=namespace))
        valor_outros = parse_decimal(prod.findtext("nfe:vOutro", default="", namespaces=namespace))
        item_row = {
            "arquivo_xml": str(file_path),
            "chave_acesso": document_key,
            "numero_nfce": document_number,
            "serie": series,
            "modelo": document_model,
            "data_emissao": emission_date,
            "cnpj_emitente": emit_cnpj,
            "emitente": emit_name,
            "item": (det.attrib.get("nItem", "") or "").strip(),
            "codigo": prod.findtext("nfe:cProd", default="", namespaces=namespace).strip(),
            "ean": first_non_empty(
                prod.findtext("nfe:cEANTrib", default="", namespaces=namespace),
                prod.findtext("nfe:cEAN", default="", namespaces=namespace),
            ),
            "descricao": prod.findtext("nfe:xProd", default="", namespaces=namespace).strip(),
            "ncm": ncm,
            "cest": prod.findtext("nfe:CEST", default="", namespaces=namespace).strip(),
            "cfop": prod.findtext("nfe:CFOP", default="", namespaces=namespace).strip(),
            "unidade": prod.findtext("nfe:uCom", default="", namespaces=namespace).strip(),
            "quantidade": parse_decimal(prod.findtext("nfe:qCom", default="", namespaces=namespace)),
            "valor_unitario": parse_decimal(prod.findtext("nfe:vUnCom", default="", namespaces=namespace)),
            "valor_produto": valor_produto,
            "valor_desconto": valor_desconto,
            "valor_frete": valor_frete,
            "valor_outros": valor_outros,
            "valor_operacao": valor_produto - valor_desconto + valor_frete + valor_outros,
            "valor_total_tributos": parse_decimal(prod.findtext("nfe:vTotTrib", default="", namespaces=namespace)),
            "orig_icms": local_name_text(icms_group, "orig"),
            "cst_icms": local_name_text(icms_group, "CST"),
            "csosn_icms": local_name_text(icms_group, "CSOSN"),
            "base_icms": parse_decimal(local_name_text(icms_group, "vBC")),
            "aliquota_icms": parse_decimal(local_name_text(icms_group, "pICMS")),
            "valor_icms": parse_decimal(local_name_text(icms_group, "vICMS")),
            "base_icms_st": parse_decimal(local_name_text(icms_group, "vBCST", "vBCSTRet")),
            "aliquota_icms_st": parse_decimal(local_name_text(icms_group, "pICMSST")),
            "valor_icms_st": parse_decimal(local_name_text(icms_group, "vICMSST", "vICMSSTRet")),
            "cst_pis": local_name_text(pis_group, "CST"),
            "base_pis": parse_decimal(local_name_text(pis_group, "vBC")),
            "aliquota_pis": parse_decimal(local_name_text(pis_group, "pPIS")),
            "valor_pis": parse_decimal(local_name_text(pis_group, "vPIS")),
            "cst_cofins": local_name_text(cofins_group, "CST"),
            "base_cofins": parse_decimal(local_name_text(cofins_group, "vBC")),
            "aliquota_cofins": parse_decimal(local_name_text(cofins_group, "pCOFINS")),
            "valor_cofins": parse_decimal(local_name_text(cofins_group, "vCOFINS")),
            "cst_ipi": local_name_text(ipi_group, "CST"),
            "base_ipi": parse_decimal(local_name_text(ipi_group, "vBC")),
            "aliquota_ipi": parse_decimal(local_name_text(ipi_group, "pIPI")),
            "valor_ipi": parse_decimal(local_name_text(ipi_group, "vIPI")),
        }
        items.append(item_row)

    return items

def collect_cest_values(items: Iterable[dict[str, object]]) -> list[str]:
    return sorted({str(item.get("cest", "")).strip() for item in items if str(item.get("cest", "")).strip()})

def format_cest_values(items: Iterable[dict[str, object]]) -> str:
    return " | ".join(collect_cest_values(items))

def build_product_origin_candidates_from_xml_sources(xml_sources: list[Path]) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []
    valid_sources = [path for path in xml_sources if path.exists()]
    if not valid_sources:
        return candidates
    try:
        xml_files = collect_xml_files(valid_sources)
    except Exception:
        return candidates
    for xml_file in xml_files:
        try:
            items = parse_nfce_xml_items(xml_file)
        except Exception:
            continue
        for item in items:
            code = str(item.get("codigo", "")).strip()
            ean = str(item.get("ean", "")).strip()
            origin_code = ean or code
            if not origin_code:
                continue
            candidates.append(
                {
                    "origin_code": origin_code,
                    "code": code,
                    "ean": ean,
                    "description": str(item.get("descricao", "")).strip(),
                    "ncm": str(item.get("ncm", "")).strip(),
                }
            )
    return candidates

def build_import_products_from_xml_sources(
    xml_sources: list[Path],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[dict[str, object]]:
    aggregated: dict[str, dict[str, object]] = {}
    xml_files = collect_xml_files([path for path in xml_sources if path.exists()])
    total_files = max(len(xml_files), 1)
    for index, xml_file in enumerate(xml_files, start=1):
        try:
            items = parse_nfce_xml_items(xml_file)
        except Exception:
            if progress_callback:
                progress_callback(index, total_files, f"Lendo XMLs... {index}/{total_files}")
            continue
        for item in items:
            codigo_origem = str(item.get("ean", "")).strip() or str(item.get("codigo", "")).strip()
            descricao = str(item.get("descricao", "")).strip()
            ncm = str(item.get("ncm", "")).strip()
            if not codigo_origem and not descricao:
                continue
            cst_pis_entrada, cst_pis_saida, pis_entrada, pis_saida, pis_warnings = build_pis_cofins_side_values(
                item.get("cst_pis", ""),
                item.get("aliquota_pis", Decimal("0")),
            )
            cst_cofins_entrada, cst_cofins_saida, cofins_entrada, cofins_saida, cofins_warnings = build_pis_cofins_side_values(
                item.get("cst_cofins", ""),
                item.get("aliquota_cofins", Decimal("0")),
            )
            bucket_key = codigo_origem or f"{normalize_text(descricao)}|{normalize_ncm(ncm)}"
            bucket = aggregated.setdefault(
                bucket_key,
                {
                    "codigo_origem": codigo_origem,
                    "descricao": descricao,
                    "ncm": ncm,
                    "unidade": str(item.get("unidade", "")).strip() or "UN",
                    "cst_icms_entrada": normalize_tax_code(item.get("cst_icms", "") or item.get("csosn_icms", ""), 3),
                    "cst_icms_saida": normalize_tax_code(item.get("cst_icms", "") or item.get("csosn_icms", ""), 3),
                    "icms_entrada": Decimal(item.get("aliquota_icms", Decimal("0"))),
                    "icms_saida": Decimal(item.get("aliquota_icms", Decimal("0"))),
                    "cst_pis_entrada": cst_pis_entrada,
                    "cst_pis_saida": cst_pis_saida,
                    "pis_entrada": pis_entrada,
                    "pis_saida": pis_saida,
                    "cst_cofins_entrada": cst_cofins_entrada,
                    "cst_cofins_saida": cst_cofins_saida,
                    "cofins_entrada": cofins_entrada,
                    "cofins_saida": cofins_saida,
                    "import_warnings": [*pis_warnings, *cofins_warnings],
                    "ativo": 1,
                },
            )
            if not bucket["codigo_origem"] and codigo_origem:
                bucket["codigo_origem"] = codigo_origem
            if not bucket["descricao"] and descricao:
                bucket["descricao"] = descricao
            if not bucket["ncm"] and ncm:
                bucket["ncm"] = ncm
            if bucket["unidade"] == "UN" and str(item.get("unidade", "")).strip():
                bucket["unidade"] = str(item.get("unidade", "")).strip()
            if not bucket["cst_icms_entrada"]:
                bucket["cst_icms_entrada"] = normalize_tax_code(item.get("cst_icms", "") or item.get("csosn_icms", ""), 3)
            if not bucket["cst_icms_saida"]:
                bucket["cst_icms_saida"] = normalize_tax_code(item.get("cst_icms", "") or item.get("csosn_icms", ""), 3)
            if not bucket["cst_pis_entrada"] and cst_pis_entrada:
                bucket["cst_pis_entrada"] = cst_pis_entrada
            if not bucket["cst_pis_saida"] and cst_pis_saida:
                bucket["cst_pis_saida"] = cst_pis_saida
            if not bucket["cst_cofins_entrada"] and cst_cofins_entrada:
                bucket["cst_cofins_entrada"] = cst_cofins_entrada
            if not bucket["cst_cofins_saida"] and cst_cofins_saida:
                bucket["cst_cofins_saida"] = cst_cofins_saida
            bucket["icms_entrada"] = max(Decimal(bucket["icms_entrada"]), Decimal(item.get("aliquota_icms", Decimal("0"))))
            bucket["icms_saida"] = max(Decimal(bucket["icms_saida"]), Decimal(item.get("aliquota_icms", Decimal("0"))))
            bucket["pis_entrada"] = max(Decimal(bucket["pis_entrada"]), pis_entrada)
            bucket["pis_saida"] = max(Decimal(bucket["pis_saida"]), pis_saida)
            bucket["cofins_entrada"] = max(Decimal(bucket["cofins_entrada"]), cofins_entrada)
            bucket["cofins_saida"] = max(Decimal(bucket["cofins_saida"]), cofins_saida)
            bucket.setdefault("import_warnings", [])
            bucket["import_warnings"].extend(pis_warnings)
            bucket["import_warnings"].extend(cofins_warnings)
        if progress_callback:
            progress_callback(index, total_files, f"Lendo XMLs... {index}/{total_files}")
    return list(aggregated.values())


def read_sped_contrib_product_rows(sped_path: Path) -> list[dict[str, object]]:
    return [
        {
            "operation_type": str(item.get("operation_type", "")).strip(),
            "code": str(item.get("code", "")).strip(),
            "description": str(item.get("description", "")).strip(),
            "ncm": str(item.get("ncm", "")).strip(),
            "cest": str(item.get("cest", "")).strip(),
            "cst_pis": str(item.get("cst_pis", "")).strip(),
            "aliquota_pis": Decimal(item.get("aliquota_pis", Decimal("0"))),
            "cst_cofins": str(item.get("cst_cofins", "")).strip(),
            "aliquota_cofins": Decimal(item.get("aliquota_cofins", Decimal("0"))),
        }
        for item in read_sped_contrib_detailed_rows(sped_path)
    ]

def read_sped_contrib_detailed_rows(sped_path: Path) -> list[dict[str, object]]:
    products: dict[str, dict[str, str]] = {}
    participants: dict[str, dict[str, str]] = {}
    detailed_rows: list[dict[str, object]] = []
    current_operation = ""
    current_document: dict[str, str] = {}

    with sped_path.open("r", encoding="latin-1") as sped_file:
        for raw_line in sped_file:
            if not raw_line.startswith("|"):
                continue
            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)

            if register == "0200":
                code = get_field(fields, 2)
                products[code] = {
                    "description": get_field(fields, 3),
                    "ncm": get_field(fields, 8),
                    "cest": get_field(fields, 13),
                }
                continue

            if register == "0150":
                participant_code = get_field(fields, 2)
                participants[participant_code] = {
                    "name": get_field(fields, 3),
                    "tax_id": first_non_empty(get_field(fields, 5), get_field(fields, 6)),
                }
                continue

            if register == "C100":
                ind_oper = get_field(fields, 2)
                current_operation = "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else ""
                participant_code = get_field(fields, 4)
                participant = participants.get(participant_code, {})
                current_document = {
                    "document_number": get_field(fields, 8),
                    "document_key": get_field(fields, 9),
                    "document_date": first_non_empty(get_field(fields, 10), get_field(fields, 11)),
                    "document_series": "",
                    "document_model": get_field(fields, 5),
                    "participant_code": participant_code,
                    "participant_name": participant.get("name", ""),
                    "participant_tax_id": participant.get("tax_id", ""),
                }
                continue

            if register != "C170":
                continue

            code = get_field(fields, 3)
            product_meta = products.get(code, {})
            detailed_rows.append(
                {
                    "operation_type": current_operation,
                    "document_number": current_document.get("document_number", ""),
                    "document_key": current_document.get("document_key", ""),
                    "document_date": current_document.get("document_date", ""),
                    "document_series": current_document.get("document_series", ""),
                    "document_model": current_document.get("document_model", ""),
                    "participant_code": current_document.get("participant_code", ""),
                    "participant_name": current_document.get("participant_name", ""),
                    "participant_tax_id": current_document.get("participant_tax_id", ""),
                    "item_number": get_field(fields, 2),
                    "code": code,
                    "description": product_meta.get("description", ""),
                    "ncm": product_meta.get("ncm", ""),
                    "cest": product_meta.get("cest", ""),
                    "cfop": get_field(fields, 11),
                    "quantity": parse_decimal(get_field(fields, 5)),
                    "sale_value": parse_decimal(get_field(fields, 7)),
                    "cst_pis": get_field(fields, 25),
                    "base_pis": parse_decimal(get_field(fields, 26)),
                    "aliquota_pis": parse_rate(get_field(fields, 27)) or Decimal("0"),
                    "pis_value": parse_decimal(get_field(fields, 30)),
                    "cst_cofins": get_field(fields, 31),
                    "base_cofins": parse_decimal(get_field(fields, 32)),
                    "aliquota_cofins": parse_rate(get_field(fields, 33)) or Decimal("0"),
                    "cofins_value": parse_decimal(get_field(fields, 36)),
                }
            )
    return detailed_rows

def summarize_pis_cofins_analysis(items: list[dict[str, object]], operation_type: str) -> dict[str, object]:
    cfops = sorted({str(item.get("cfop", "")).strip() for item in items if str(item.get("cfop", "")).strip()})
    csts_pis = sorted({normalize_tax_code(item.get("cst_pis", ""), 2) for item in items if str(item.get("cst_pis", "")).strip()})
    csts_cofins = sorted({normalize_tax_code(item.get("cst_cofins", ""), 2) for item in items if str(item.get("cst_cofins", "")).strip()})
    total_pis = sum((Decimal(item.get("pis_value", Decimal("0"))) for item in items), Decimal("0"))
    total_cofins = sum((Decimal(item.get("cofins_value", Decimal("0"))) for item in items), Decimal("0"))

    missing_status = "Sem entrada" if normalize_operation_type(operation_type) == "Entrada" else "Sem saida"
    zero_status = "Sem credito" if normalize_operation_type(operation_type) == "Entrada" else "Sem debito"
    status_parts: list[str] = []
    if not items:
        status_parts.append(missing_status)
    else:
        if not cfops:
            status_parts.append("Erro: CFOP vazio")
        elif len(cfops) > 1:
            status_parts.append("Erro: multiplos CFOPs")
        if not csts_pis:
            status_parts.append("Erro: CST PIS vazio")
        elif len(csts_pis) > 1:
            status_parts.append("Erro: multiplos CSTs PIS")
        if not csts_cofins:
            status_parts.append("Erro: CST COFINS vazio")
        elif len(csts_cofins) > 1:
            status_parts.append("Erro: multiplos CSTs COFINS")
        if total_pis == Decimal("0") and total_cofins == Decimal("0"):
            status_parts.append(zero_status)
        if not status_parts:
            status_parts.append("Ok")

    return {
        "cfop": " | ".join(cfops),
        "cst_pis": " | ".join(csts_pis),
        "cst_cofins": " | ".join(csts_cofins),
        "pis_value": total_pis,
        "cofins_value": total_cofins,
        "status": " | ".join(status_parts),
    }

def build_import_products_from_sped_contrib_sources(
    sped_paths: list[Path],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[dict[str, object]]:
    aggregated: dict[str, dict[str, object]] = {}
    unique_paths = [path for path in dict.fromkeys(path for path in sped_paths if path.exists())]
    total_paths = max(len(unique_paths), 1)
    for index, sped_path in enumerate(unique_paths, start=1):
        for item in read_sped_contrib_product_rows(sped_path):
            codigo_origem = str(item.get("code", "")).strip()
            descricao = str(item.get("description", "")).strip()
            ncm = str(item.get("ncm", "")).strip()
            operation_type = normalize_operation_type(str(item.get("operation_type", "")))
            bucket_key = codigo_origem or f"{normalize_text(descricao)}|{normalize_ncm(ncm)}"
            bucket = aggregated.setdefault(
                bucket_key,
                {
                    "codigo_origem": codigo_origem,
                    "descricao": descricao,
                    "ncm": ncm,
                    "unidade": "UN",
                    "cst_icms_entrada": "",
                    "cst_icms_saida": "",
                    "icms_entrada": Decimal("0"),
                    "icms_saida": Decimal("0"),
                    "cst_pis_entrada": "",
                    "cst_pis_saida": "",
                    "pis_entrada": Decimal("0"),
                    "pis_saida": Decimal("0"),
                    "cst_cofins_entrada": "",
                    "cst_cofins_saida": "",
                    "cofins_entrada": Decimal("0"),
                    "cofins_saida": Decimal("0"),
                    "ativo": 1,
                },
            )
            if not bucket["descricao"] and descricao:
                bucket["descricao"] = descricao
            if not bucket["ncm"] and ncm:
                bucket["ncm"] = ncm
            cst_pis = normalize_tax_code(item.get("cst_pis", ""), 2)
            cst_cofins = normalize_tax_code(item.get("cst_cofins", ""), 2)
            aliquota_pis = Decimal(item.get("aliquota_pis", Decimal("0")) or Decimal("0"))
            aliquota_cofins = Decimal(item.get("aliquota_cofins", Decimal("0")) or Decimal("0"))
            if operation_type == "Entrada":
                if not bucket["cst_pis_entrada"] and (cst_pis in VALID_PIS_COFINS_CST_ENTRADA or cst_pis in VALID_PIS_COFINS_CST_AMBOS):
                    bucket["cst_pis_entrada"] = cst_pis
                if not bucket["cst_cofins_entrada"] and (cst_cofins in VALID_PIS_COFINS_CST_ENTRADA or cst_cofins in VALID_PIS_COFINS_CST_AMBOS):
                    bucket["cst_cofins_entrada"] = cst_cofins
                bucket["pis_entrada"] = max(Decimal(bucket["pis_entrada"]), aliquota_pis)
                bucket["cofins_entrada"] = max(Decimal(bucket["cofins_entrada"]), aliquota_cofins)
            elif operation_type == "Saida":
                if not bucket["cst_pis_saida"] and (cst_pis in VALID_PIS_COFINS_CST_SAIDA or cst_pis in VALID_PIS_COFINS_CST_AMBOS):
                    bucket["cst_pis_saida"] = cst_pis
                if not bucket["cst_cofins_saida"] and (cst_cofins in VALID_PIS_COFINS_CST_SAIDA or cst_cofins in VALID_PIS_COFINS_CST_AMBOS):
                    bucket["cst_cofins_saida"] = cst_cofins
                bucket["pis_saida"] = max(Decimal(bucket["pis_saida"]), aliquota_pis)
                bucket["cofins_saida"] = max(Decimal(bucket["cofins_saida"]), aliquota_cofins)
            warnings = bucket.setdefault("import_warnings", [])
            if cst_pis and ((operation_type == "Entrada" and cst_pis not in VALID_PIS_COFINS_CST_ENTRADA and cst_pis not in VALID_PIS_COFINS_CST_AMBOS) or (operation_type == "Saida" and cst_pis not in VALID_PIS_COFINS_CST_SAIDA and cst_pis not in VALID_PIS_COFINS_CST_AMBOS)):
                warnings.append(f"CST PIS {cst_pis} invalido para {operation_type.lower()}.")
            if cst_cofins and ((operation_type == "Entrada" and cst_cofins not in VALID_PIS_COFINS_CST_ENTRADA and cst_cofins not in VALID_PIS_COFINS_CST_AMBOS) or (operation_type == "Saida" and cst_cofins not in VALID_PIS_COFINS_CST_SAIDA and cst_cofins not in VALID_PIS_COFINS_CST_AMBOS)):
                warnings.append(f"CST COFINS {cst_cofins} invalido para {operation_type.lower()}.")
        if progress_callback:
            progress_callback(index, total_paths, f"Lendo SPED PIS/COFINS... {index}/{total_paths}")
    return list(aggregated.values())


def read_sped_file(sped_path: Path):
    from app.services.legacy_rules import read_sped_file as legacy_read_sped_file

    return legacy_read_sped_file(sped_path)


def build_product_origin_candidates_from_sped_file(sped_path: Path) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []
    try:
        products, _sales_rows, detailed_rows, _c190_rows, _c190_product_rows = read_sped_file(sped_path)
    except Exception:
        return candidates
    seen_codes: set[str] = set()
    for product in products:
        code = str(product.code).strip()
        if not code or code in seen_codes:
            continue
        seen_codes.add(code)
        candidates.append(
            {
                "origin_code": code,
                "code": code,
                "ean": "",
                "description": str(product.description).strip(),
                "ncm": str(product.ncm).strip(),
            }
        )
    for item in detailed_rows:
        code = str(item.get("code", "")).strip()
        if not code or code in seen_codes:
            continue
        seen_codes.add(code)
        candidates.append(
            {
                "origin_code": code,
                "code": code,
                "ean": "",
                "description": str(item.get("description", "")).strip(),
                "ncm": str(item.get("ncm", "")).strip(),
            }
        )
    return candidates


def build_import_products_from_sped_0200(
    sped_path: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[dict[str, object]]:
    products, _sales_rows, _detailed_rows, _c190_rows, _c190_product_rows = read_sped_file(sped_path)
    imported_products: list[dict[str, object]] = []
    total_products = max(len(products), 1)
    for index, product in enumerate(products, start=1):
        imported_products.append(
            {
                "codigo_origem": str(product.code).strip(),
                "descricao": str(product.description).strip(),
                "ncm": str(product.ncm).strip(),
                "unidade": "UN",
                "cst_icms_entrada": "",
                "cst_icms_saida": "",
                "icms_entrada": product.icms_rate if isinstance(product.icms_rate, Decimal) else Decimal("0"),
                "icms_saida": product.icms_rate if isinstance(product.icms_rate, Decimal) else Decimal("0"),
                "cst_pis_entrada": "",
                "cst_pis_saida": "",
                "pis_entrada": Decimal("0"),
                "pis_saida": Decimal("0"),
                "cst_cofins_entrada": "",
                "cst_cofins_saida": "",
                "cofins_entrada": Decimal("0"),
                "cofins_saida": Decimal("0"),
                "ativo": 1,
            }
        )
        if progress_callback:
            progress_callback(index, total_products, f"Lendo SPED 0200... {index}/{total_products}")
    return imported_products

def build_import_products_from_sped_fiscal_sources(
    sped_paths: list[Path],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[dict[str, object]]:
    aggregated: dict[str, dict[str, object]] = {}
    unique_paths = [path for path in dict.fromkeys(path for path in sped_paths if path.exists())]
    total_paths = max(len(unique_paths), 1)
    for index, sped_path in enumerate(unique_paths, start=1):
        products, _sales_rows, detailed_rows, _c190_rows, _c190_product_rows = read_sped_file(sped_path)
        for product in products:
            codigo_origem = str(product.code).strip()
            descricao = str(product.description).strip()
            ncm = str(product.ncm).strip()
            bucket_key = codigo_origem or f"{normalize_text(descricao)}|{normalize_ncm(ncm)}"
            bucket = aggregated.setdefault(
                bucket_key,
                {
                    "codigo_origem": codigo_origem,
                    "descricao": descricao,
                    "ncm": ncm,
                    "unidade": "UN",
                    "cst_icms_entrada": normalize_tax_code(product.cst_icms, 3),
                    "cst_icms_saida": normalize_tax_code(product.cst_icms, 3),
                    "icms_entrada": Decimal(product.icms_rate or Decimal("0")),
                    "icms_saida": Decimal(product.icms_rate or Decimal("0")),
                    "cst_pis_entrada": "",
                    "cst_pis_saida": "",
                    "pis_entrada": Decimal("0"),
                    "pis_saida": Decimal("0"),
                    "cst_cofins_entrada": "",
                    "cst_cofins_saida": "",
                    "cofins_entrada": Decimal("0"),
                    "cofins_saida": Decimal("0"),
                    "ativo": 1,
                },
            )
            if not bucket["descricao"] and descricao:
                bucket["descricao"] = descricao
            if not bucket["ncm"] and ncm:
                bucket["ncm"] = ncm

        for item in detailed_rows:
            codigo_origem = str(item.get("code", "")).strip()
            descricao = str(item.get("description", "")).strip()
            ncm = str(item.get("ncm", "")).strip()
            operation_type = normalize_operation_type(str(item.get("operation_type", "")))
            bucket_key = codigo_origem or f"{normalize_text(descricao)}|{normalize_ncm(ncm)}"
            bucket = aggregated.setdefault(
                bucket_key,
                {
                    "codigo_origem": codigo_origem,
                    "descricao": descricao,
                    "ncm": ncm,
                    "unidade": "UN",
                    "cst_icms_entrada": "",
                    "cst_icms_saida": "",
                    "icms_entrada": Decimal("0"),
                    "icms_saida": Decimal("0"),
                    "cst_pis_entrada": "",
                    "cst_pis_saida": "",
                    "pis_entrada": Decimal("0"),
                    "pis_saida": Decimal("0"),
                    "cst_cofins_entrada": "",
                    "cst_cofins_saida": "",
                    "cofins_entrada": Decimal("0"),
                    "cofins_saida": Decimal("0"),
                    "ativo": 1,
                },
            )
            if not bucket["descricao"] and descricao:
                bucket["descricao"] = descricao
            if not bucket["ncm"] and ncm:
                bucket["ncm"] = ncm
            cst_icms = normalize_tax_code(item.get("cst_icms", ""), 3)
            icms_rate = Decimal(item.get("icms_rate", Decimal("0")) or Decimal("0"))
            if operation_type == "Entrada":
                if not bucket["cst_icms_entrada"] and cst_icms:
                    bucket["cst_icms_entrada"] = cst_icms
                bucket["icms_entrada"] = max(Decimal(bucket["icms_entrada"]), icms_rate)
            elif operation_type == "Saida":
                if not bucket["cst_icms_saida"] and cst_icms:
                    bucket["cst_icms_saida"] = cst_icms
                bucket["icms_saida"] = max(Decimal(bucket["icms_saida"]), icms_rate)
        if progress_callback:
            progress_callback(index, total_paths, f"Lendo SPED Fiscal... {index}/{total_paths}")
    return list(aggregated.values())

def build_import_products_from_consolidated_sources(
    xml_sources: list[Path],
    sped_fiscal_paths: list[Path],
    sped_contrib_paths: list[Path],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[dict[str, object]]:
    fiscal_rows = build_import_products_from_sped_fiscal_sources(sped_fiscal_paths, progress_callback=progress_callback) if sped_fiscal_paths else []
    contrib_rows = build_import_products_from_sped_contrib_sources(sped_contrib_paths, progress_callback=progress_callback) if sped_contrib_paths else []
    xml_rows = build_import_products_from_xml_sources(xml_sources, progress_callback=progress_callback) if xml_sources else []

    if not fiscal_rows and not contrib_rows and not xml_rows:
        return []

    consolidated: dict[str, dict[str, object]] = {}

    def get_bucket_key(row: dict[str, object]) -> str:
        codigo_origem = str(row.get("codigo_origem", "")).strip()
        descricao = str(row.get("descricao", "")).strip()
        ncm = str(row.get("ncm", "")).strip()
        return codigo_origem or f"{normalize_text(descricao)}|{normalize_ncm(ncm)}"

    def get_or_create_bucket(row: dict[str, object]) -> dict[str, object]:
        bucket_key = get_bucket_key(row)
        return consolidated.setdefault(
            bucket_key,
            {
                "codigo_origem": str(row.get("codigo_origem", "")).strip(),
                "descricao": str(row.get("descricao", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "unidade": str(row.get("unidade", "UN")).strip() or "UN",
                "tipo_produto": str(row.get("tipo_produto", "Revenda")).strip() or "Revenda",
                "cst_icms_entrada": "",
                "cst_icms_saida": "",
                "icms_entrada": Decimal("0"),
                "icms_saida": Decimal("0"),
                "cst_pis_entrada": "",
                "cst_pis_saida": "",
                "pis_entrada": Decimal("0"),
                "pis_saida": Decimal("0"),
                "cst_cofins_entrada": "",
                "cst_cofins_saida": "",
                "cofins_entrada": Decimal("0"),
                "cofins_saida": Decimal("0"),
                "ativo": 1,
            },
        )

    for row in fiscal_rows:
        bucket = get_or_create_bucket(row)
        if not bucket["descricao"] and row.get("descricao"):
            bucket["descricao"] = str(row.get("descricao", "")).strip()
        if not bucket["ncm"] and row.get("ncm"):
            bucket["ncm"] = str(row.get("ncm", "")).strip()
        if not bucket["codigo_origem"] and row.get("codigo_origem"):
            bucket["codigo_origem"] = str(row.get("codigo_origem", "")).strip()
        if not bucket["cst_icms_entrada"] and row.get("cst_icms_entrada"):
            bucket["cst_icms_entrada"] = normalize_tax_code(row.get("cst_icms_entrada", ""), 3)
        if not bucket["cst_icms_saida"] and row.get("cst_icms_saida"):
            bucket["cst_icms_saida"] = normalize_tax_code(row.get("cst_icms_saida", ""), 3)
        bucket["icms_entrada"] = max(Decimal(bucket["icms_entrada"]), Decimal(row.get("icms_entrada", Decimal("0"))))
        bucket["icms_saida"] = max(Decimal(bucket["icms_saida"]), Decimal(row.get("icms_saida", Decimal("0"))))

    for row in contrib_rows:
        bucket = get_or_create_bucket(row)
        if not bucket["descricao"] and row.get("descricao"):
            bucket["descricao"] = str(row.get("descricao", "")).strip()
        if not bucket["ncm"] and row.get("ncm"):
            bucket["ncm"] = str(row.get("ncm", "")).strip()
        if not bucket["codigo_origem"] and row.get("codigo_origem"):
            bucket["codigo_origem"] = str(row.get("codigo_origem", "")).strip()
        if not bucket["cst_pis_entrada"] and row.get("cst_pis_entrada"):
            bucket["cst_pis_entrada"] = normalize_tax_code(row.get("cst_pis_entrada", ""), 2)
        if not bucket["cst_pis_saida"] and row.get("cst_pis_saida"):
            bucket["cst_pis_saida"] = normalize_tax_code(row.get("cst_pis_saida", ""), 2)
        if not bucket["cst_cofins_entrada"] and row.get("cst_cofins_entrada"):
            bucket["cst_cofins_entrada"] = normalize_tax_code(row.get("cst_cofins_entrada", ""), 2)
        if not bucket["cst_cofins_saida"] and row.get("cst_cofins_saida"):
            bucket["cst_cofins_saida"] = normalize_tax_code(row.get("cst_cofins_saida", ""), 2)
        bucket["pis_entrada"] = max(Decimal(bucket["pis_entrada"]), Decimal(row.get("pis_entrada", Decimal("0"))))
        bucket["pis_saida"] = max(Decimal(bucket["pis_saida"]), Decimal(row.get("pis_saida", Decimal("0"))))
        bucket["cofins_entrada"] = max(Decimal(bucket["cofins_entrada"]), Decimal(row.get("cofins_entrada", Decimal("0"))))
        bucket["cofins_saida"] = max(Decimal(bucket["cofins_saida"]), Decimal(row.get("cofins_saida", Decimal("0"))))

    for row in xml_rows:
        bucket = get_or_create_bucket(row)
        if not bucket["codigo_origem"] and row.get("codigo_origem"):
            bucket["codigo_origem"] = str(row.get("codigo_origem", "")).strip()
        if (not bucket["codigo_origem"] or bucket["codigo_origem"] == str(row.get("descricao", "")).strip()) and row.get("codigo_origem"):
            bucket["codigo_origem"] = str(row.get("codigo_origem", "")).strip()
        if not bucket["descricao"] and row.get("descricao"):
            bucket["descricao"] = str(row.get("descricao", "")).strip()
        if not bucket["ncm"] and row.get("ncm"):
            bucket["ncm"] = str(row.get("ncm", "")).strip()
        if bucket["unidade"] == "UN" and str(row.get("unidade", "")).strip():
            bucket["unidade"] = str(row.get("unidade", "")).strip()
        if not bucket["cst_icms_entrada"] and row.get("cst_icms_entrada"):
            bucket["cst_icms_entrada"] = normalize_tax_code(row.get("cst_icms_entrada", ""), 3)
        if not bucket["cst_icms_saida"] and row.get("cst_icms_saida"):
            bucket["cst_icms_saida"] = normalize_tax_code(row.get("cst_icms_saida", ""), 3)
        if not bucket["cst_pis_entrada"] and row.get("cst_pis_entrada"):
            bucket["cst_pis_entrada"] = normalize_tax_code(row.get("cst_pis_entrada", ""), 2)
        if not bucket["cst_pis_saida"] and row.get("cst_pis_saida"):
            bucket["cst_pis_saida"] = normalize_tax_code(row.get("cst_pis_saida", ""), 2)
        if not bucket["cst_cofins_entrada"] and row.get("cst_cofins_entrada"):
            bucket["cst_cofins_entrada"] = normalize_tax_code(row.get("cst_cofins_entrada", ""), 2)
        if not bucket["cst_cofins_saida"] and row.get("cst_cofins_saida"):
            bucket["cst_cofins_saida"] = normalize_tax_code(row.get("cst_cofins_saida", ""), 2)
        bucket["icms_entrada"] = max(Decimal(bucket["icms_entrada"]), Decimal(row.get("icms_entrada", Decimal("0"))))
        bucket["icms_saida"] = max(Decimal(bucket["icms_saida"]), Decimal(row.get("icms_saida", Decimal("0"))))
        bucket["pis_entrada"] = max(Decimal(bucket["pis_entrada"]), Decimal(row.get("pis_entrada", Decimal("0"))))
        bucket["pis_saida"] = max(Decimal(bucket["pis_saida"]), Decimal(row.get("pis_saida", Decimal("0"))))
        bucket["cofins_entrada"] = max(Decimal(bucket["cofins_entrada"]), Decimal(row.get("cofins_entrada", Decimal("0"))))
        bucket["cofins_saida"] = max(Decimal(bucket["cofins_saida"]), Decimal(row.get("cofins_saida", Decimal("0"))))

    return list(consolidated.values())
