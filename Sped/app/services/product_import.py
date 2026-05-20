from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET

from app.parsers.sped_parser import first_non_empty, normalize_document_key, parse_decimal
from app.services.tax_rules import (
    build_pis_cofins_side_values,
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
