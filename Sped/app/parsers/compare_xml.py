from __future__ import annotations

import re
from pathlib import Path
from xml.etree import ElementTree as ET

from app.config import COMPARE_KEY_PATTERN, COMPARE_NS_NFE, COMPARE_NS_NFSE
from app.models import CompareXmlInvoice, CompareXmlItem
from app.services.tax_rules import normalize_text


def compare_clean(value: object) -> str:
    return str(value or "").strip()


def compare_sanitize(value: object) -> str:
    return compare_clean(value).replace("|", " ")


def compare_to_float(value: object) -> float:
    text = compare_clean(value)
    if not text:
        return 0.0
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def compare_extract_money_values(text: object) -> list[float]:
    raw_text = compare_clean(text)
    if not raw_text:
        return []
    values: list[float] = []
    for match in re.finditer(r"(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2}|\d+\.\d{2})", raw_text, flags=re.IGNORECASE):
        parsed = compare_to_float(match.group(1))
        if parsed > 0:
            values.append(parsed)
    return values


def compare_format_float(value: float, decimals: int = 2) -> str:
    return f"{value:.{decimals}f}".replace(".", ",")


def compare_xml_text(parent: ET.Element | None, xpath: str) -> str:
    if parent is None:
        return ""
    node = parent.find(xpath, COMPARE_NS_NFE)
    return compare_clean(node.text if node is not None else "")


def compare_nfse_text(parent: ET.Element | None, xpath: str) -> str:
    if parent is None:
        return ""
    node = parent.find(xpath, COMPARE_NS_NFSE)
    return compare_clean(node.text if node is not None else "")


def compare_extract_key(root: ET.Element) -> str:
    inf_nfe = root.find(".//nfe:infNFe", COMPARE_NS_NFE)
    if inf_nfe is not None:
        raw_id = compare_clean(inf_nfe.attrib.get("Id", ""))
        if raw_id.startswith("NFe") and len(raw_id) >= 47:
            key = raw_id[3:47]
            if len(key) == 44 and key.isdigit():
                return key
    protocol_key = compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:chNFe")
    if len(protocol_key) == 44 and protocol_key.isdigit():
        return protocol_key
    return ""


def compare_extract_nfse_key(root: ET.Element) -> str:
    inf_nfse = root.find(".//nfse:infNFSe", COMPARE_NS_NFSE)
    raw_id = compare_clean(inf_nfse.attrib.get("Id", "")) if inf_nfse is not None else ""
    if raw_id:
        return raw_id[3:] if raw_id.upper().startswith("NFS") else raw_id
    number = compare_nfse_text(root, ".//nfse:infNFSe/nfse:nNFSe")
    provider = compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:CNPJ") or compare_nfse_text(root, ".//nfse:infDPS/nfse:prest/nfse:CNPJ")
    if number and provider:
        return f"NFSE{normalize_document_key(provider)}{number}"
    return ""


def compare_is_nfse_invoice(invoice: object) -> bool:
    return str(getattr(invoice, "model", "") or "").upper().startswith("NFS")


def compare_date_to_sped(value: object) -> str:
    date_text = compare_clean(value).split("T")[0]
    parts = date_text.split("-")
    if len(parts) == 3:
        return f"{parts[2]}{parts[1]}{parts[0]}"
    return date_text if len(date_text) == 8 and date_text.isdigit() else ""


def compare_extract_icms(det: ET.Element) -> tuple[str, float, float, float, float, float, float]:
    icms_parent = det.find("./nfe:imposto/nfe:ICMS", COMPARE_NS_NFE)
    if icms_parent is None or len(icms_parent) == 0:
        return "000", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    icms_node = list(icms_parent)[0]
    cst = (
        compare_clean(icms_node.findtext("nfe:CST", default="", namespaces=COMPARE_NS_NFE))
        or compare_clean(icms_node.findtext("nfe:CSOSN", default="", namespaces=COMPARE_NS_NFE))
        or "000"
    )
    return (
        cst,
        compare_to_float(icms_node.findtext("nfe:vBC", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(icms_node.findtext("nfe:pICMS", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(icms_node.findtext("nfe:vICMS", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(icms_node.findtext("nfe:vBCST", default="", namespaces=COMPARE_NS_NFE))
        or compare_to_float(icms_node.findtext("nfe:vBCSTRet", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(icms_node.findtext("nfe:pICMSST", default="", namespaces=COMPARE_NS_NFE))
        or compare_to_float(icms_node.findtext("nfe:pST", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(icms_node.findtext("nfe:vICMSST", default="", namespaces=COMPARE_NS_NFE))
        or compare_to_float(icms_node.findtext("nfe:vICMSSTRet", default="", namespaces=COMPARE_NS_NFE)),
    )


def compare_extract_pis_cofins(det: ET.Element, register_name: str) -> tuple[str, float, float, float]:
    parent = det.find(f"./nfe:imposto/nfe:{register_name}", COMPARE_NS_NFE)
    if parent is None or len(parent) == 0:
        return "99", 0.0, 0.0, 0.0
    node = list(parent)[0]
    cst = compare_clean(node.findtext("nfe:CST", default="", namespaces=COMPARE_NS_NFE)) or "99"
    base_tag = "vBC"
    rate_tag = "pPIS" if register_name == "PIS" else "pCOFINS"
    value_tag = "vPIS" if register_name == "PIS" else "vCOFINS"
    return (
        cst,
        compare_to_float(node.findtext(f"nfe:{base_tag}", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(node.findtext(f"nfe:{rate_tag}", default="", namespaces=COMPARE_NS_NFE)),
        compare_to_float(node.findtext(f"nfe:{value_tag}", default="", namespaces=COMPARE_NS_NFE)),
    )


def compare_extract_ipi(det: ET.Element) -> float:
    parent = det.find("./nfe:imposto/nfe:IPI", COMPARE_NS_NFE)
    if parent is None:
        return 0.0
    for node in list(parent):
        value = compare_to_float(node.findtext("nfe:vIPI", default="", namespaces=COMPARE_NS_NFE))
        if value:
            return value
    return 0.0


def parse_compare_nfse_file(root: ET.Element, file_path: Path) -> CompareXmlInvoice | None:
    key = compare_extract_nfse_key(root)
    if not key:
        return None

    issue_datetime = (
        compare_nfse_text(root, ".//nfse:infNFSe/nfse:dhProc")
        or compare_nfse_text(root, ".//nfse:infDPS/nfse:dhEmi")
        or compare_nfse_text(root, ".//nfse:infDPS/nfse:dCompet")
    )
    number = compare_nfse_text(root, ".//nfse:infNFSe/nfse:nNFSe") or compare_nfse_text(root, ".//nfse:infDPS/nfse:nDPS")
    series = compare_nfse_text(root, ".//nfse:infDPS/nfse:serie")
    provider_cnpj = compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:CNPJ") or compare_nfse_text(root, ".//nfse:infDPS/nfse:prest/nfse:CNPJ")
    provider_name = compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:xNome")
    taker_cnpj = compare_nfse_text(root, ".//nfse:infDPS/nfse:toma/nfse:CNPJ") or compare_nfse_text(root, ".//nfse:infDPS/nfse:toma/nfse:CPF")
    taker_name = compare_nfse_text(root, ".//nfse:infDPS/nfse:toma/nfse:xNome")
    raw_description = (
        compare_nfse_text(root, ".//nfse:infDPS/nfse:serv/nfse:cServ/nfse:xDescServ")
        or compare_nfse_text(root, ".//nfse:infNFSe/nfse:xTribNac")
        or "Servico"
    )
    service_value = compare_to_float(compare_nfse_text(root, ".//nfse:infDPS/nfse:valores/nfse:vServPrest/nfse:vServ"))
    total_doc = service_value or compare_to_float(compare_nfse_text(root, ".//nfse:infNFSe/nfse:valores/nfse:vLiq"))
    description_values = compare_extract_money_values(raw_description)
    if description_values:
        inferred_total = max(description_values)
        if inferred_total > total_doc:
            total_doc = inferred_total

    item = CompareXmlItem(
        item_no="1",
        code="32156",
        ean="",
        description="PRESTACAO DE SERVICOS ADMINISTRATIVOS",
        unit="UN",
        ncm="39181000",
        cest="",
        c_classtrib="",
        cfop="1933",
        quantity=1.0,
        value=total_doc,
        discount=0.0,
        cst_icms="090",
        vl_bc_icms=0.0,
        aliq_icms=0.0,
        vl_icms=0.0,
        vl_bc_icms_st=0.0,
        aliq_icms_st=0.0,
        vl_icms_st=0.0,
        vl_ipi=0.0,
        cst_pis="50",
        vl_bc_pis=total_doc,
        aliq_pis=1.65,
        vl_pis=round(total_doc * 0.0165, 2),
        cst_cofins="50",
        vl_bc_cofins=total_doc,
        aliq_cofins=7.6,
        vl_cofins=round(total_doc * 0.076, 2),
    )
    return CompareXmlInvoice(
        key=key,
        number=number,
        series=series,
        issue_date=issue_datetime,
        issuer_cnpj=provider_cnpj,
        issuer_name=provider_name,
        issuer_ie=compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:IM") or compare_nfse_text(root, ".//nfse:infDPS/nfse:prest/nfse:IM"),
        issuer_city_code=compare_nfse_text(root, ".//nfse:infDPS/nfse:cLocEmi"),
        recipient_cnpj=taker_cnpj,
        recipient_name=taker_name,
        recipient_ie="",
        total_value=compare_format_float(total_doc),
        date_doc=compare_date_to_sped(issue_datetime),
        date_entry=compare_date_to_sped(issue_datetime),
        model="NFSe",
        freight_mode="9",
        total_doc=total_doc,
        total_discount=0.0,
        total_products=total_doc,
        total_freight=0.0,
        total_insurance=0.0,
        total_other=0.0,
        total_bc_icms=0.0,
        total_icms=0.0,
        total_bc_st=0.0,
        total_st=0.0,
        total_ipi=0.0,
        total_pis=item.vl_pis,
        total_cofins=item.vl_cofins,
        file_path=str(file_path),
        items=[item],
        issuer_address=compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:enderNac/nfse:xLgr"),
        issuer_number=compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:enderNac/nfse:nro"),
        issuer_complement=compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:enderNac/nfse:xCpl"),
        issuer_district=compare_nfse_text(root, ".//nfse:infNFSe/nfse:emit/nfse:enderNac/nfse:xBairro"),
    )


def parse_compare_xml_file(file_path: Path) -> CompareXmlInvoice | None:
    try:
        root = ET.parse(file_path).getroot()
    except ET.ParseError:
        return None
    if root.tag.endswith("NFSe") or root.find(".//nfse:infNFSe", COMPARE_NS_NFSE) is not None:
        return parse_compare_nfse_file(root, file_path)
    if root.find(".//nfe:infNFe", COMPARE_NS_NFE) is None:
        return None
    key = compare_extract_key(root)
    if not key:
        return None

    issue_datetime = compare_xml_text(root, ".//nfe:ide/nfe:dhEmi") or compare_xml_text(root, ".//nfe:ide/nfe:dEmi")
    entry_datetime = (
        compare_xml_text(root, ".//nfe:ide/nfe:dhSaiEnt")
        or compare_xml_text(root, ".//nfe:ide/nfe:dSaiEnt")
        or issue_datetime
    )

    items: list[CompareXmlItem] = []
    for det in root.findall(".//nfe:det", COMPARE_NS_NFE):
        cst_icms, vl_bc_icms, aliq_icms, vl_icms, vl_bc_st, aliq_st, vl_st = compare_extract_icms(det)
        vl_ipi = compare_extract_ipi(det)
        cst_pis, vl_bc_pis, aliq_pis, vl_pis = compare_extract_pis_cofins(det, "PIS")
        cst_cofins, vl_bc_cofins, aliq_cofins, vl_cofins = compare_extract_pis_cofins(det, "COFINS")
        items.append(
            CompareXmlItem(
                item_no=compare_clean(det.attrib.get("nItem", "")),
                code=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:cProd")),
                ean=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:cEAN")),
                description=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:xProd")),
                unit=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:uCom")),
                ncm=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:NCM")),
                cest=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:CEST")),
                c_classtrib=compare_sanitize(
                    compare_xml_text(det, "./nfe:prod/nfe:cClassTrib")
                    or compare_xml_text(det, "./nfe:imposto/nfe:IBSCBS/nfe:cClassTrib")
                ),
                cfop=compare_sanitize(compare_xml_text(det, "./nfe:prod/nfe:CFOP")),
                quantity=compare_to_float(compare_xml_text(det, "./nfe:prod/nfe:qCom")),
                value=compare_to_float(compare_xml_text(det, "./nfe:prod/nfe:vProd")),
                discount=compare_to_float(compare_xml_text(det, "./nfe:prod/nfe:vDesc")),
                cst_icms=cst_icms,
                vl_bc_icms=vl_bc_icms,
                aliq_icms=aliq_icms,
                vl_icms=vl_icms,
                vl_bc_icms_st=vl_bc_st,
                aliq_icms_st=aliq_st,
                vl_icms_st=vl_st,
                vl_ipi=vl_ipi,
                cst_pis=cst_pis,
                vl_bc_pis=vl_bc_pis,
                aliq_pis=aliq_pis,
                vl_pis=vl_pis,
                cst_cofins=cst_cofins,
                vl_bc_cofins=vl_bc_cofins,
                aliq_cofins=aliq_cofins,
                vl_cofins=vl_cofins,
            )
        )

    total_doc = compare_to_float(
        compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vNF")
        or compare_xml_text(root, ".//nfe:total/nfe:vNFTot")
    )
    return CompareXmlInvoice(
        key=key,
        number=compare_xml_text(root, ".//nfe:ide/nfe:nNF"),
        series=compare_xml_text(root, ".//nfe:ide/nfe:serie"),
        issue_date=issue_datetime,
        issuer_cnpj=compare_xml_text(root, ".//nfe:emit/nfe:CNPJ"),
        issuer_name=compare_xml_text(root, ".//nfe:emit/nfe:xNome"),
        issuer_ie=compare_xml_text(root, ".//nfe:emit/nfe:IE"),
        issuer_city_code=compare_xml_text(root, ".//nfe:emit/nfe:enderEmit/nfe:cMun"),
        recipient_cnpj=compare_xml_text(root, ".//nfe:dest/nfe:CNPJ") or compare_xml_text(root, ".//nfe:dest/nfe:CPF"),
        recipient_name=compare_xml_text(root, ".//nfe:dest/nfe:xNome"),
        recipient_ie=compare_xml_text(root, ".//nfe:dest/nfe:IE"),
        total_value=compare_format_float(total_doc),
        date_doc=compare_date_to_sped(issue_datetime),
        date_entry=compare_date_to_sped(entry_datetime) or compare_date_to_sped(issue_datetime),
        model=compare_xml_text(root, ".//nfe:ide/nfe:mod") or "55",
        freight_mode=compare_xml_text(root, ".//nfe:transp/nfe:modFrete") or "9",
        total_doc=total_doc,
        total_discount=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vDesc")),
        total_products=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vProd")),
        total_freight=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vFrete")),
        total_insurance=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vSeg")),
        total_other=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vOutro")),
        total_bc_icms=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vBC")),
        total_icms=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vICMS")),
        total_bc_st=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vBCST")),
        total_st=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vST")),
        total_ipi=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vIPI")),
        total_pis=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vPIS")),
        total_cofins=compare_to_float(compare_xml_text(root, ".//nfe:total/nfe:ICMSTot/nfe:vCOFINS")),
        file_path=str(file_path),
        items=items,
        issuer_address=compare_xml_text(root, ".//nfe:emit/nfe:enderEmit/nfe:xLgr"),
        issuer_number=compare_xml_text(root, ".//nfe:emit/nfe:enderEmit/nfe:nro"),
        issuer_complement=compare_xml_text(root, ".//nfe:emit/nfe:enderEmit/nfe:xCpl"),
        issuer_district=compare_xml_text(root, ".//nfe:emit/nfe:enderEmit/nfe:xBairro"),
    )


def describe_compare_ignored_xml(xml_path: Path) -> str:
    try:
        root = ET.parse(xml_path).getroot()
    except ET.ParseError as exc:
        return f"XML invalido: {exc}"
    tag = root.tag.split("}", 1)[-1]
    xml_text = ""
    try:
        xml_text = xml_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        try:
            xml_text = xml_path.read_text(encoding="latin-1", errors="ignore")
        except Exception:
            xml_text = ""
    upper_text = xml_text.upper()
    if "PROCEVENTONFE" in tag.upper() or "<EVENTO" in upper_text:
        return "Evento de NF-e/NFC-e, nao e documento fiscal de lancamento."
    if "CANCEL" in upper_text:
        return "XML de cancelamento/evento, nao e documento fiscal de lancamento."
    if "INUT" in tag.upper() or "INUTILIZ" in upper_text:
        return "XML de inutilizacao, nao e documento fiscal de lancamento."
    if root.find(".//nfe:infNFe", COMPARE_NS_NFE) is not None:
        return "NF-e/NFC-e sem chave de acesso valida ou estrutura incompleta."
    if root.find(".//nfse:infNFSe", COMPARE_NS_NFSE) is not None or tag.upper().endswith("NFSE"):
        return "NFS-e sem identificacao valida ou estrutura incompleta."
    return f"Tipo de XML nao suportado para comparacao: {tag or 'raiz desconhecida'}."


def extract_xml_cancellation_event(xml_path: Path) -> dict[str, str] | None:
    try:
        root = ET.parse(xml_path).getroot()
    except ET.ParseError:
        return None
    tag = root.tag.split("}", 1)[-1].upper()
    event_type = compare_xml_text(root, ".//nfe:infEvento/nfe:tpEvento")
    event_description = compare_xml_text(root, ".//nfe:infEvento/nfe:detEvento/nfe:descEvento")
    protocol_status = (
        compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:cStat")
        or compare_xml_text(root, ".//nfe:retEvento/nfe:infEvento/nfe:cStat")
        or compare_xml_text(root, ".//nfe:infProt/nfe:cStat")
    )
    protocol_reason = (
        compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:xMotivo")
        or compare_xml_text(root, ".//nfe:retEvento/nfe:infEvento/nfe:xMotivo")
        or compare_xml_text(root, ".//nfe:infProt/nfe:xMotivo")
    )
    key = compare_xml_text(root, ".//nfe:infEvento/nfe:chNFe")
    if not key:
        key = compare_xml_text(root, ".//nfe:infCanc/nfe:chNFe")
    if not key:
        key = compare_xml_text(root, ".//nfe:retEvento/nfe:infEvento/nfe:chNFe")
    if not key:
        key = compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:chNFe")
    if not key:
        key = compare_xml_text(root, ".//nfe:infProt/nfe:chNFe")
    if not key:
        key = compare_extract_key(root)
    normalized_description = normalize_text(event_description)
    normalized_reason = normalize_text(protocol_reason)
    is_cancellation = (
        event_type == "110111"
        or protocol_status in {"101", "151", "155"}
        or "CANCEL" in tag
        or "CANCELAMENTO" in normalized_description
        or "CANCELAMENTO" in normalized_reason
    )
    if not is_cancellation or len(key) != 44 or not key.isdigit():
        return None
    reason = (
        protocol_reason
        or event_description
        or "Evento de cancelamento localizado"
    )
    protocol = (
        compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:nProt")
        or compare_xml_text(root, ".//nfe:retEvento/nfe:infEvento/nfe:nProt")
        or compare_xml_text(root, ".//nfe:infProt/nfe:nProt")
    )
    event_date = (
        compare_xml_text(root, ".//nfe:protNFe/nfe:infProt/nfe:dhRecbto")
        or compare_xml_text(root, ".//nfe:retEvento/nfe:infEvento/nfe:dhRegEvento")
        or compare_xml_text(root, ".//nfe:infEvento/nfe:dhEvento")
    )
    return {
        "key": key,
        "status": protocol_status,
        "reason": reason,
        "protocol": protocol,
        "event_date": event_date,
        "file_path": str(xml_path),
    }


def collect_xml_cancellation_events(xml_files: list[Path]) -> dict[str, dict[str, str]]:
    cancellation_events: dict[str, dict[str, str]] = {}
    for xml_path in xml_files:
        event = extract_xml_cancellation_event(xml_path)
        if event is not None:
            cancellation_events[event["key"]] = event
    return cancellation_events


def collect_compare_ignored_xml_rows(xml_folder: Path) -> list[list[object]]:
    rows: list[list[object]] = []
    xml_paths = sorted(xml_folder.rglob("*.xml")) if xml_folder.exists() else []
    for xml_path in xml_paths:
        cancellation_event = extract_xml_cancellation_event(xml_path)
        if cancellation_event is not None:
            rows.append([xml_path.name, f"NF-e/NFC-e cancelada: {cancellation_event.get('reason', '')}", str(xml_path)])
            continue
        if parse_compare_xml_file(xml_path) is not None:
            continue
        rows.append([xml_path.name, describe_compare_ignored_xml(xml_path), str(xml_path)])
    return rows
