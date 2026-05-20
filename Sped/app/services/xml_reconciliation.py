from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree as ET

from app.models import CompareXmlItem
from app.parsers.compare_xml import (
    compare_extract_icms,
    compare_extract_ipi,
    compare_extract_pis_cofins,
    compare_xml_text,
    parse_compare_xml_file,
)
from app.parsers.sped_fiscal_parser import read_sped_file
from app.parsers.sped_parser import (
    first_non_empty,
    get_field,
    normalize_cfop,
    normalize_document_key,
    normalize_sped_line,
    parse_decimal,
    parse_rate,
    read_sped_plain_lines,
)
from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.services.analysis_utils import (
    calculate_abc_curve_labels,
    filter_details_by_operation_scope,
    infer_sped_period_label,
)
from app.services.path_selection import get_xml_worker_count
from app.services.product_import import (
    collect_cest_values,
    collect_xml_files,
    get_xml_local_name,
    local_name_text,
    ncm_matches_filters,
    normalize_ncm,
    parse_nfce_xml_items,
    read_sped_contrib_detailed_rows,
    summarize_pis_cofins_analysis,
)
from app.services.tax_rules import (
    build_pis_cofins_side_values,
    normalize_cst_icms_for_sped,
    normalize_operation_type,
    normalize_tax_code,
    normalize_text,
)

def set_field(fields: list[str], index: int, value: str) -> None:
    while len(fields) <= index:
        fields.append("")
    fields[index] = value

def display_text(value: object, default: str = "") -> str:
    if value is None:
        return default
    return str(value)

def map_xml_cfop_to_entry_cfops(cfop: str) -> set[str]:
    normalized_cfop = normalize_cfop(cfop)
    if normalized_cfop in {"1401", "1403"}:
        return {normalized_cfop}
    if normalized_cfop in {"5401", "5403", "5405"}:
        # Na entrada o SPED pode estar em 1401 ou 1403, mesmo quando o XML veio
        # com CFOP de saida 5401/5403/5405. O casamento final ainda exige chave + item.
        return {"1401", "1403"}
    return set()

def parse_nfe_xml_st_items(file_path: Path) -> tuple[str, dict[tuple[str, str], list[dict[str, object]]]]:
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

    items: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    if not document_key:
        return "", items

    for det in root.findall(".//nfe:det", namespace):
        xml_cfop = det.findtext("nfe:prod/nfe:CFOP", default="", namespaces=namespace)
        target_cfops = map_xml_cfop_to_entry_cfops(xml_cfop)
        if not target_cfops:
            continue

        icms_group = det.find("nfe:imposto/nfe:ICMS", namespace)
        if icms_group is None:
            continue

        base_icms_st = Decimal("0")
        icms_st_rate = Decimal("0")
        icms_st_value = Decimal("0")
        for element in icms_group.iter():
            local_name = get_xml_local_name(element.tag)
            text = (element.text or "").strip()
            if not text:
                continue
            if local_name in {"vBCST", "vBCSTRet"}:
                base_icms_st = parse_decimal(text)
            elif local_name == "pICMSST":
                icms_st_rate = parse_decimal(text)
            elif local_name in {"vICMSST", "vICMSSTRet"}:
                icms_st_value = parse_decimal(text)

        item_number = (det.attrib.get("nItem", "") or "").strip()
        code = first_non_empty(
            det.findtext("nfe:prod/nfe:cEANTrib", default="", namespaces=namespace),
            det.findtext("nfe:prod/nfe:cProd", default="", namespaces=namespace),
        )
        xml_codes = {
            normalize_document_key(det.findtext("nfe:prod/nfe:cEANTrib", default="", namespaces=namespace)),
            normalize_document_key(det.findtext("nfe:prod/nfe:cEAN", default="", namespaces=namespace)),
            normalize_document_key(det.findtext("nfe:prod/nfe:cProd", default="", namespaces=namespace)),
        }
        xml_codes.discard("")
        if not item_number and not code:
            continue

        for target_cfop in target_cfops:
            items[(item_number, target_cfop)].append(
                {
                    "code": str(code).strip(),
                    "xml_codes": xml_codes,
                    "xml_cfop": normalize_cfop(xml_cfop),
                    "cfop": target_cfop,
                    "base_icms_st": base_icms_st,
                    "icms_st_rate": icms_st_rate,
                    "icms_st_value": icms_st_value,
                }
            )

    return document_key, items

def collect_xml_candidate_document_keys(detailed_sales: list[dict[str, object]]) -> set[str]:
    candidate_keys: set[str] = set()
    for item in detailed_sales:
        if str(item.get("operation_type", "")).strip() != "Entrada":
            continue
        original_cfop = str(item.get("original_cfop", item.get("cfop", ""))).strip()
        if original_cfop not in {"1401", "1403"}:
            continue
        document_key = normalize_document_key(str(item.get("document_key", "")))
        if document_key:
            candidate_keys.add(document_key)
    return candidate_keys

def read_xml_document_key(file_path: Path) -> str:
    try:
        for _, element in ET.iterparse(file_path, events=("end",)):
            local_name = get_xml_local_name(element.tag)
            if local_name == "chNFe":
                text = (element.text or "").strip()
                if text:
                    return normalize_document_key(text)
            if local_name == "infNFe":
                inf_nfe_id = element.attrib.get("Id", "")
                if inf_nfe_id.startswith("NFe"):
                    inf_nfe_id = inf_nfe_id[3:]
                normalized = normalize_document_key(inf_nfe_id)
                if normalized:
                    return normalized
        return ""
    except ET.ParseError:
        return ""

def build_pis_cofins_period_comparison_rows(
    sped_paths: list[Path],
    operation_type: str,
    xml_sources: list[Path] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[str], list[dict[str, object]]]:
    if not sped_paths:
        raise ValueError("Selecione ao menos um arquivo SPED Contribuicoes para processar.")
    if len(sped_paths) > 12:
        raise ValueError("A consulta aceita no maximo 12 arquivos SPED por vez.")

    normalized_operation = normalize_operation_type(operation_type)
    normalized_xml_sources = [path for path in dict.fromkeys(xml_sources or []) if path.exists()]
    period_labels: list[str] = []
    comparison_rows: list[dict[str, object]] = []
    used_labels: set[str] = set()

    total_files = len(sped_paths)
    for file_index, sped_path in enumerate(sped_paths, start=1):
        if not sped_path.exists():
            raise ValueError(f"O arquivo SPED nao foi encontrado: {sped_path}")
        if progress_callback:
            progress_callback(file_index - 1, total_files, f"Lendo {normalized_operation.lower()}: {sped_path.name}")

        detailed_rows = [
            item
            for item in read_sped_contrib_detailed_rows(sped_path)
            if normalize_operation_type(str(item.get("operation_type", "")).strip()) == normalized_operation
        ]
        document_rows = scan_sped_contrib_c100_documents(sped_path, normalized_operation)
        if normalized_xml_sources:
            rebuilt_rows = rebuild_sped_contrib_items_from_xml(document_rows, detailed_rows, normalized_operation, normalized_xml_sources)
            if rebuilt_rows:
                detailed_rows.extend(rebuilt_rows)
        label_base = infer_sped_period_label(sped_path, detailed_rows)
        label = label_base
        suffix = 2
        while label in used_labels:
            label = f"{label_base} ({suffix})"
            suffix += 1
        used_labels.add(label)
        period_labels.append(label)

        grouped_items: dict[str, list[dict[str, object]]] = defaultdict(list)
        for item in detailed_rows:
            code = str(item.get("code", "")).strip() or str(item.get("description", "")).strip()
            if code:
                grouped_items[code].append(item)

        for code, items in grouped_items.items():
            cfops = sorted({str(item.get("cfop", "")).strip() for item in items if str(item.get("cfop", "")).strip()})
            csts_pis = sorted({normalize_tax_code(item.get("cst_pis", ""), 2) for item in items if str(item.get("cst_pis", "")).strip()})
            csts_cofins = sorted({normalize_tax_code(item.get("cst_cofins", ""), 2) for item in items if str(item.get("cst_cofins", "")).strip()})
            descriptions = sorted({str(item.get("description", "")).strip() for item in items if str(item.get("description", "")).strip()})
            ncms = sorted({str(item.get("ncm", "")).strip() for item in items if str(item.get("ncm", "")).strip()})
            cests = collect_cest_values(items)
            document_keys = {
                normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
                for item in items
                if normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
            }
            launch_details = [dict(item) for item in sorted(items, key=lambda current: (str(current.get("document_date", "")), str(current.get("document_number", "")), str(current.get("item_number", ""))))]
            summary = summarize_pis_cofins_analysis(launch_details, normalized_operation)
            pis_base = sum((Decimal(item.get("base_pis", Decimal("0"))) for item in items), Decimal("0"))
            cofins_base = sum((Decimal(item.get("base_cofins", Decimal("0"))) for item in items), Decimal("0"))
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
                    "cst_pis": " | ".join(csts_pis),
                    "cst_cofins": " | ".join(csts_cofins),
                    "suppliers": " | ".join(sorted({str(item.get("participant_name", "")).strip() for item in items if str(item.get("participant_name", "")).strip()})),
                    "supplier_count": len({(str(item.get("participant_name", "")).strip(), str(item.get("participant_tax_id", "")).strip()) for item in items if str(item.get("participant_name", "")).strip() or str(item.get("participant_tax_id", "")).strip()}),
                    "quantity": sum((Decimal(item.get("quantity", Decimal("0"))) for item in items), Decimal("0")),
                    "sale_value": sum((Decimal(item.get("sale_value", Decimal("0"))) for item in items), Decimal("0")),
                    "base_pis": pis_base,
                    "base_cofins": cofins_base,
                    "pis_value": Decimal(summary["pis_value"]),
                    "cofins_value": Decimal(summary["cofins_value"]),
                    "aliquota_pis": next((Decimal(item.get("aliquota_pis", Decimal("0"))) for item in launch_details if Decimal(item.get("aliquota_pis", Decimal("0"))) > 0), Decimal("0")),
                    "aliquota_cofins": next((Decimal(item.get("aliquota_cofins", Decimal("0"))) for item in launch_details if Decimal(item.get("aliquota_cofins", Decimal("0"))) > 0), Decimal("0")),
                    "document_count": len(document_keys),
                    "launch_count": len(items),
                    "status": str(summary["status"]),
                    "launch_details": launch_details,
                }
            )

        if progress_callback:
            progress_callback(file_index, total_files, f"{normalized_operation}s processadas: {sped_path.name}")

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
            str(item.get("code", "")),
            str(item.get("period", "")),
            str(item.get("cfop", "")),
            str(item.get("cst_pis", "")),
            str(item.get("cst_cofins", "")),
        )
    )
    return period_labels, comparison_rows

def load_nfce_xml_items_for_index(
    file_path: Path,
    candidate_document_keys: set[str] | None = None,
    ncm_filters: set[str] | None = None,
) -> list[dict[str, object]]:
    if candidate_document_keys:
        document_key_hint = read_xml_document_key(file_path)
        if not document_key_hint or document_key_hint not in candidate_document_keys:
            return []

    items = parse_nfce_xml_items(file_path)
    if not items:
        return []

    filtered_items: list[dict[str, object]] = []
    for item in items:
        document_key = normalize_document_key(str(item.get("chave_acesso", "")))
        if not document_key:
            continue
        if candidate_document_keys and document_key not in candidate_document_keys:
            continue
        if ncm_filters and not ncm_matches_filters(str(item.get("ncm", "")), ncm_filters):
            continue
        filtered_items.append(item)
    return filtered_items

def export_nfce_items_by_ncm(
    xml_sources: list[Path],
    ncm_filters: set[str],
    output_path: Path,
) -> tuple[int, int, int]:
    xml_files = collect_xml_files(xml_sources)
    if not xml_files:
        raise ValueError("Nenhum arquivo XML foi encontrado nos caminhos selecionados.")

    filtered_items: list[dict[str, object]] = []
    skipped_files: list[list[object]] = []

    worker_count = get_xml_worker_count(len(xml_files))
    if worker_count == 1:
        file_entries = []
        for file_path in xml_files:
            try:
                items = load_nfce_xml_items_for_index(file_path, ncm_filters=ncm_filters)
            except ET.ParseError as exc:
                skipped_files.append([str(file_path), f"XML invalido: {exc}"])
                continue
            except Exception as exc:
                skipped_files.append([str(file_path), str(exc)])
                continue
            file_entries.append((file_path, items))
    else:
        file_entries = []
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_to_path = {
                executor.submit(load_nfce_xml_items_for_index, file_path, None, ncm_filters): file_path
                for file_path in xml_files
            }
            for future in as_completed(future_to_path):
                file_path = future_to_path[future]
                try:
                    items = future.result()
                except ET.ParseError as exc:
                    skipped_files.append([str(file_path), f"XML invalido: {exc}"])
                    continue
                except Exception as exc:
                    skipped_files.append([str(file_path), str(exc)])
                    continue
                file_entries.append((file_path, items))

    for _, items in sorted(file_entries, key=lambda current: str(current[0])):
        filtered_items.extend(items)

    if not filtered_items:
        raise ValueError("Nenhum item de NFC-e foi encontrado para o(s) NCM(s) informado(s).")

    headers = [
        "Arquivo XML",
        "Chave de Acesso",
        "Numero NFC-e",
        "Serie",
        "Data Emissao",
        "CNPJ Emitente",
        "Emitente",
        "Item",
        "Codigo",
        "EAN",
        "Descricao",
        "NCM",
        "CFOP",
        "Unidade",
        "Quantidade",
        "Valor Unitario",
        "Valor Produto",
        "Valor Desconto",
        "Valor Frete",
        "Valor Outros",
        "Valor Total Tributos",
        "Orig ICMS",
        "CST ICMS",
        "CSOSN ICMS",
        "Base ICMS",
        "Aliquota ICMS",
        "Valor ICMS",
        "Base ICMS ST",
        "Aliquota ICMS ST",
        "Valor ICMS ST",
        "CST PIS",
        "Base PIS",
        "Aliquota PIS",
        "Valor PIS",
        "CST COFINS",
        "Base COFINS",
        "Aliquota COFINS",
        "Valor COFINS",
        "CST IPI",
        "Base IPI",
        "Aliquota IPI",
        "Valor IPI",
    ]
    item_rows = [
        [
            item["arquivo_xml"],
            item["chave_acesso"],
            item["numero_nfce"],
            item["serie"],
            item["data_emissao"],
            item["cnpj_emitente"],
            item["emitente"],
            item["item"],
            item["codigo"],
            item["ean"],
            item["descricao"],
            item["ncm"],
            item["cfop"],
            item["unidade"],
            item["quantidade"],
            item["valor_unitario"],
            item["valor_produto"],
            item["valor_desconto"],
            item["valor_frete"],
            item["valor_outros"],
            item["valor_total_tributos"],
            item["orig_icms"],
            item["cst_icms"],
            item["csosn_icms"],
            item["base_icms"],
            item["aliquota_icms"],
            item["valor_icms"],
            item["base_icms_st"],
            item["aliquota_icms_st"],
            item["valor_icms_st"],
            item["cst_pis"],
            item["base_pis"],
            item["aliquota_pis"],
            item["valor_pis"],
            item["cst_cofins"],
            item["base_cofins"],
            item["aliquota_cofins"],
            item["valor_cofins"],
            item["cst_ipi"],
            item["base_ipi"],
            item["aliquota_ipi"],
            item["valor_ipi"],
        ]
        for item in filtered_items
    ]

    resumo_por_ncm: dict[str, dict[str, object]] = {}
    for item in filtered_items:
        ncm = str(item["ncm"])
        bucket = resumo_por_ncm.setdefault(
            ncm,
            {
                "ncm": ncm,
                "quantidade_itens": 0,
                "valor_produto": Decimal("0"),
                "valor_icms": Decimal("0"),
                "valor_icms_st": Decimal("0"),
                "valor_pis": Decimal("0"),
                "valor_cofins": Decimal("0"),
                "valor_ipi": Decimal("0"),
                "valor_total_tributos": Decimal("0"),
            },
        )
        bucket["quantidade_itens"] = int(bucket["quantidade_itens"]) + 1
        for field_name in ("valor_produto", "valor_icms", "valor_icms_st", "valor_pis", "valor_cofins", "valor_ipi", "valor_total_tributos"):
            bucket[field_name] = bucket[field_name] + item[field_name]

    resumo_headers = [
        "NCM",
        "Quantidade Itens",
        "Valor Produto",
        "Valor ICMS",
        "Valor ICMS ST",
        "Valor PIS",
        "Valor COFINS",
        "Valor IPI",
        "Valor Total Tributos",
    ]
    resumo_rows = [
        [
            row["ncm"],
            row["quantidade_itens"],
            row["valor_produto"],
            row["valor_icms"],
            row["valor_icms_st"],
            row["valor_pis"],
            row["valor_cofins"],
            row["valor_ipi"],
            row["valor_total_tributos"],
        ]
        for row in sorted(resumo_por_ncm.values(), key=lambda current: str(current["ncm"]))
    ]

    sheets = [("Itens NFCe", headers, item_rows), ("Resumo NCM", resumo_headers, resumo_rows)]
    if skipped_files:
        sheets.append(("XMLs Ignorados", ["Arquivo XML", "Motivo"], skipped_files))

    write_simple_excel_workbook(output_path, sheets)
    return len(xml_files), len(filtered_items), len(skipped_files)

def compose_xml_icms_cst_for_sped(xml_item: dict[str, object], fallback: str = "") -> str:
    orig = "".join(char for char in str(xml_item.get("orig_icms", "")).strip() if char.isdigit())
    cst = "".join(char for char in str(xml_item.get("cst_icms", "")).strip() if char.isdigit())
    csosn = "".join(char for char in str(xml_item.get("csosn_icms", "")).strip() if char.isdigit())
    suffix = csosn or cst
    if suffix:
        return normalize_cst_icms_for_sped(f"{orig}{suffix}")
    return normalize_cst_icms_for_sped(fallback)

def build_xml_fiscal_item_index(
    xml_sources: list[Path],
    candidate_document_keys: set[str] | None = None,
    ncm_filters: set[str] | None = None,
) -> dict[str, list[dict[str, object]]]:
    xml_files = collect_xml_files(xml_sources)
    if not xml_files:
        return {}

    xml_index: dict[str, list[dict[str, object]]] = defaultdict(list)
    worker_count = get_xml_worker_count(len(xml_files))
    if worker_count == 1:
        indexed_entries = []
        for file_path in xml_files:
            try:
                items = load_nfce_xml_items_for_index(file_path, candidate_document_keys, ncm_filters)
            except (ET.ParseError, UnicodeDecodeError):
                continue
            indexed_entries.append((file_path, items))
    else:
        indexed_entries = []
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_to_path = {
                executor.submit(load_nfce_xml_items_for_index, file_path, candidate_document_keys, ncm_filters): file_path
                for file_path in xml_files
            }
            for future in as_completed(future_to_path):
                file_path = future_to_path[future]
                try:
                    items = future.result()
                except (ET.ParseError, UnicodeDecodeError):
                    continue
                indexed_entries.append((file_path, items))

    for _, items in sorted(indexed_entries, key=lambda current: str(current[0])):
        for item in items:
            document_key = normalize_document_key(str(item.get("chave_acesso", "")))
            if document_key:
                xml_index[document_key].append(item)
    return xml_index

def build_xml_fiscal_identity_index(xml_index: dict[str, list[dict[str, object]]]) -> dict[tuple[str, str, str], list[dict[str, object]]]:
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
    return xml_index_by_identity

def scan_sped_contrib_c100_documents(
    sped_path: Path,
    operation_scope: str,
) -> dict[tuple[str, str], dict[str, object]]:
    documents: dict[tuple[str, str], dict[str, object]] = {}
    participants: dict[str, dict[str, str]] = {}
    scoped_operations = {
        str(item["operation_type"])
        for item in filter_details_by_operation_scope(
            [{"operation_type": "Entrada"}, {"operation_type": "Saida"}],
            operation_scope,
        )
    }

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|"):
            continue
        fields = normalize_sped_line(line)
        register = get_field(fields, 1)
        if register == "0150":
            participant_code = get_field(fields, 2)
            participants[participant_code] = {
                "name": get_field(fields, 3),
                "tax_id": first_non_empty(get_field(fields, 5), get_field(fields, 6)),
            }
            continue
        if register != "C100":
            continue
        operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
        if operation_type not in scoped_operations:
            continue
        document_number = get_field(fields, 8)
        document_key = normalize_document_key(get_field(fields, 9))
        if not document_key:
            continue
        participant_code = get_field(fields, 4)
        participant = participants.get(participant_code, {})
        documents[(operation_type, document_key)] = {
            "operation_type": operation_type,
            "document_number": document_number,
            "document_key": document_key,
            "document_date": first_non_empty(get_field(fields, 10), get_field(fields, 11)),
            "document_series": get_field(fields, 7),
            "participant_code": participant_code,
            "participant_name": participant.get("name", ""),
            "participant_tax_id": participant.get("tax_id", ""),
            "cod_mod": get_field(fields, 5),
            "cod_sit": get_field(fields, 6),
        }
    return documents

def rebuild_sped_contrib_items_from_xml(
    document_rows: dict[tuple[str, str], dict[str, object]],
    detailed_rows: list[dict[str, object]],
    operation_type: str,
    xml_sources: list[Path],
) -> list[dict[str, object]]:
    document_keys_with_items = {
        normalize_document_key(str(item.get("document_key", "")))
        for item in detailed_rows
        if normalize_document_key(str(item.get("document_key", "")))
    }
    candidate_document_keys = {
        document_key
        for (row_operation, document_key), document_meta in document_rows.items()
        if row_operation == operation_type and str(document_meta.get("cod_mod", "")).strip() in {"55", "65"} and document_key not in document_keys_with_items
    }
    if not candidate_document_keys:
        return []

    xml_index = build_xml_fiscal_item_index(xml_sources, candidate_document_keys=candidate_document_keys)
    if not xml_index:
        return []
    xml_index_by_identity = build_xml_fiscal_identity_index(xml_index)

    rebuilt_rows: list[dict[str, object]] = []
    for (row_operation, document_key), document_meta in document_rows.items():
        if row_operation != operation_type:
            continue
        if str(document_meta.get("cod_mod", "")).strip() not in {"55", "65"}:
            continue
        if document_key in document_keys_with_items:
            continue
        xml_items = xml_index.get(document_key, [])
        if not xml_items:
            xml_items = xml_index_by_identity.get(
                (
                    str(document_meta.get("cod_mod", "")).strip(),
                    str(document_meta.get("document_number", "")).strip(),
                    str(document_meta.get("document_series", "")).strip(),
                ),
                [],
            )
        if not xml_items:
            continue
        for xml_item in xml_items:
            rebuilt_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(document_meta.get("document_number", "")).strip() or str(xml_item.get("numero_nfce", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(document_meta.get("document_date", "")).strip() or str(xml_item.get("data_emissao", "")).strip(),
                    "document_series": str(document_meta.get("document_series", "")).strip() or str(xml_item.get("serie", "")).strip(),
                    "document_model": str(document_meta.get("cod_mod", "")).strip() or str(xml_item.get("modelo", "")).strip() or "55",
                    "participant_code": str(document_meta.get("participant_code", "")).strip(),
                    "participant_name": str(document_meta.get("participant_name", "")).strip() or str(xml_item.get("emitente", "")).strip(),
                    "participant_tax_id": str(document_meta.get("participant_tax_id", "")).strip() or str(xml_item.get("cnpj_emitente", "")).strip(),
                    "item_number": str(xml_item.get("item", "")).strip(),
                    "code": str(xml_item.get("codigo", "")).strip() or str(xml_item.get("descricao", "")).strip(),
                    "description": str(xml_item.get("descricao", "")).strip(),
                    "ncm": str(xml_item.get("ncm", "")).strip(),
                    "cest": str(xml_item.get("cest", "")).strip(),
                    "cfop": str(xml_item.get("cfop", "")).strip(),
                    "quantity": Decimal(xml_item.get("quantidade", Decimal("0"))),
                    "sale_value": Decimal(xml_item.get("valor_operacao", xml_item.get("valor_produto", Decimal("0")))),
                    "cst_pis": normalize_tax_code(xml_item.get("cst_pis", ""), 2),
                    "base_pis": Decimal(xml_item.get("base_pis", Decimal("0"))),
                    "aliquota_pis": Decimal(xml_item.get("aliquota_pis", Decimal("0"))),
                    "pis_value": Decimal(xml_item.get("valor_pis", Decimal("0"))),
                    "cst_cofins": normalize_tax_code(xml_item.get("cst_cofins", ""), 2),
                    "base_cofins": Decimal(xml_item.get("base_cofins", Decimal("0"))),
                    "aliquota_cofins": Decimal(xml_item.get("aliquota_cofins", Decimal("0"))),
                    "cofins_value": Decimal(xml_item.get("valor_cofins", Decimal("0"))),
                    "xml_rebuilt": True,
                }
            )
    return rebuilt_rows

def scan_sped_c100_documents(
    sped_path: Path,
    operation_scope: str,
) -> dict[tuple[str, str], dict[str, object]]:
    documents: dict[tuple[str, str], dict[str, object]] = {}
    scoped_operations = {
        str(item["operation_type"])
        for item in filter_details_by_operation_scope(
            [{"operation_type": "Entrada"}, {"operation_type": "Saida"}],
            operation_scope,
        )
    }

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|C100|"):
            continue
        fields = normalize_sped_line(line)
        operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
        if operation_type not in scoped_operations:
            continue
        document_number = get_field(fields, 8)
        document_key = normalize_document_key(get_field(fields, 9))
        if not document_key:
            continue
        documents[(operation_type, document_key)] = {
            "operation_type": operation_type,
            "document_number": document_number,
            "document_key": document_key,
            "document_date": get_field(fields, 10),
            "document_series": get_field(fields, 7),
            "participant_code": get_field(fields, 4),
            "cod_mod": get_field(fields, 5),
            "cod_sit": get_field(fields, 6),
        }
    return documents

def read_c100_c190_fallback_rows(sped_path: Path, operation_type: str, cod_mod: str = "65") -> list[dict[str, object]]:
    fallback_rows: list[dict[str, object]] = []
    current_document: dict[str, object] | None = None

    for line in read_sped_plain_lines(sped_path):
        if not line.startswith("|"):
            continue
        fields = normalize_sped_line(line)
        register = get_field(fields, 1)

        if register == "C100":
            line_operation_type = "Entrada" if get_field(fields, 2) == "0" else "Saida" if get_field(fields, 2) == "1" else ""
            document_key = normalize_document_key(get_field(fields, 9))
            current_document = None
            if line_operation_type != operation_type:
                continue
            if str(get_field(fields, 5)).strip() != str(cod_mod).strip():
                continue
            if not document_key:
                continue
            current_document = {
                "operation_type": line_operation_type,
                "document_number": get_field(fields, 8),
                "document_key": document_key,
                "document_date": get_field(fields, 10),
                "document_series": get_field(fields, 7),
                "document_model": get_field(fields, 5),
                "participant_code": get_field(fields, 4),
                "participant_name": "",
                "participant_tax_id": "",
            }
            continue

        if register == "C190" and current_document:
            document_model = str(current_document.get("document_model", "")).strip()
            document_number = str(current_document.get("document_number", "")).strip()
            document_key = str(current_document.get("document_key", "")).strip()
            fallback_rows.append(
                {
                    **current_document,
                    "item_number": "",
                    "code": f"DOC-{document_number or document_key[:12]}",
                    "description": f"Documento modelo {document_model or cod_mod} sem itens C170 - fallback C190",
                    "ncm": "",
                    "cest": "",
                    "cst_icms": get_field(fields, 2),
                    "cfop": get_field(fields, 3),
                    "quantity": Decimal("0"),
                    "icms_rate": parse_rate(get_field(fields, 4)) or Decimal("0"),
                    "sale_value": parse_decimal(get_field(fields, 5)),
                    "base_icms": parse_decimal(get_field(fields, 6)),
                    "icms_value": parse_decimal(get_field(fields, 7)),
                    "base_icms_st": parse_decimal(get_field(fields, 8)),
                    "icms_st_rate": Decimal("0"),
                    "icms_st_value": parse_decimal(get_field(fields, 9)),
                    "ipi_value": parse_decimal(get_field(fields, 11)),
                    "fallback_source": "C100/C190",
                }
            )
            continue

    return fallback_rows

def allocate_decimal_proportionally(
    weights: list[Decimal],
    target_total: Decimal,
) -> list[Decimal]:
    if not weights:
        return []
    quantized_target = target_total.quantize(Decimal("0.01"))
    positive_weights = [max(weight, Decimal("0")) for weight in weights]
    total_weight = sum(positive_weights, Decimal("0"))
    if total_weight <= 0:
        allocations = [Decimal("0.00")] * len(weights)
        allocations[-1] = quantized_target
        return allocations

    allocations: list[Decimal] = []
    allocated_total = Decimal("0.00")
    for index, weight in enumerate(positive_weights):
        if index == len(positive_weights) - 1:
            value = (quantized_target - allocated_total).quantize(Decimal("0.01"))
        else:
            value = (quantized_target * weight / total_weight).quantize(Decimal("0.01"))
            allocated_total += value
        allocations.append(value)
    return allocations

def normalize_xml_rebuilt_items_with_fallback(
    rebuilt_items: list[dict[str, object]],
    fallback_rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    if not rebuilt_items or not fallback_rows:
        return rebuilt_items

    fallback_groups: dict[tuple[str, str, Decimal], dict[str, Decimal]] = {}
    fallback_groups_by_cfop_cst: dict[tuple[str, str], list[tuple[tuple[str, str, Decimal], dict[str, Decimal]]]] = defaultdict(list)
    for fallback_row in fallback_rows:
        key = (
            normalize_cst_icms_for_sped(str(fallback_row.get("cst_icms", ""))),
            str(fallback_row.get("cfop", "")).strip(),
            Decimal(fallback_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01")),
        )
        bucket = fallback_groups.setdefault(
            key,
            {
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_icms_st": Decimal("0"),
                "icms_st_value": Decimal("0"),
                "ipi_value": Decimal("0"),
            },
        )
        bucket["base_icms"] += Decimal(fallback_row.get("base_icms", Decimal("0")))
        bucket["icms_value"] += Decimal(fallback_row.get("icms_value", Decimal("0")))
        bucket["base_icms_st"] += Decimal(fallback_row.get("base_icms_st", Decimal("0")))
        bucket["icms_st_value"] += Decimal(fallback_row.get("icms_st_value", Decimal("0")))
        bucket["ipi_value"] += Decimal(fallback_row.get("ipi_value", Decimal("0")))
    for key, totals in fallback_groups.items():
        fallback_groups_by_cfop_cst[(key[0], key[1])].append((key, totals))

    rebuilt_groups: dict[tuple[str, str, Decimal], list[dict[str, object]]] = defaultdict(list)
    for rebuilt_item in rebuilt_items:
        key = (
            normalize_cst_icms_for_sped(str(rebuilt_item.get("cst_icms", ""))),
            str(rebuilt_item.get("cfop", "")).strip(),
            Decimal(rebuilt_item.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01")),
        )
        rebuilt_groups[key].append(rebuilt_item)

    for key, grouped_items in rebuilt_groups.items():
        fallback_group_key = key
        fallback_totals = fallback_groups.get(fallback_group_key)
        if not fallback_totals:
            fallback_candidates = fallback_groups_by_cfop_cst.get((key[0], key[1]), [])
            if len(fallback_candidates) == 1:
                fallback_group_key, fallback_totals = fallback_candidates[0]
        if not fallback_totals:
            continue
        weights = [Decimal(item.get("sale_value", Decimal("0"))) for item in grouped_items]
        for field_name in ("base_icms", "icms_value", "base_icms_st", "icms_st_value", "ipi_value"):
            allocated_values = allocate_decimal_proportionally(weights, fallback_totals[field_name])
            for item, value in zip(grouped_items, allocated_values):
                item[field_name] = value
                item["icms_rate"] = fallback_group_key[2]
    return rebuilt_items

def find_matching_xml_fiscal_item(
    detail: dict[str, object],
    xml_items: list[dict[str, object]],
    used_indexes: set[int],
) -> tuple[int, dict[str, object]] | tuple[None, None]:
    item_number = str(detail.get("item_number", "")).strip()
    code = normalize_document_key(str(detail.get("code", "")))
    description = normalize_text(str(detail.get("description", "")))
    ncm = normalize_ncm(str(detail.get("ncm", "")))

    if item_number:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            if str(xml_item.get("item", "")).strip() == item_number:
                return index, xml_item

    if code:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            xml_codes = {
                normalize_document_key(str(xml_item.get("codigo", ""))),
                normalize_document_key(str(xml_item.get("ean", ""))),
            }
            xml_codes.discard("")
            if code in xml_codes:
                return index, xml_item

    if description and ncm:
        for index, xml_item in enumerate(xml_items):
            if index in used_indexes:
                continue
            if (
                normalize_text(str(xml_item.get("descricao", ""))) == description
                and normalize_ncm(str(xml_item.get("ncm", ""))) == ncm
            ):
                return index, xml_item

    return None, None

def apply_xml_fiscal_adjustments_to_details(
    detailed_sales: list[dict[str, object]],
    sped_documents: dict[tuple[str, str], dict[str, object]],
    xml_sources: list[Path],
    ncm_filters: set[str],
) -> tuple[list[dict[str, object]], set[tuple[str, str]], list[dict[str, object]]]:
    if not xml_sources or not ncm_filters or not sped_documents:
        return list(detailed_sales), set(), []

    candidate_document_keys = {document_key for _, document_key in sped_documents}
    xml_index = build_xml_fiscal_item_index(xml_sources, candidate_document_keys, ncm_filters)
    if not xml_index:
        return list(detailed_sales), set(), []

    affected_documents: set[tuple[str, str]] = set()
    adjustment_rows: list[dict[str, object]] = []
    xml_rebuilt_details: list[dict[str, object]] = []

    for document_identity, document_meta in sped_documents.items():
        xml_items = xml_index.get(str(document_meta.get("document_key", "")), [])
        if not xml_items:
            continue
        affected_documents.add(document_identity)
        for xml_item in xml_items:
            rebuilt_item = {
                "operation_type": str(document_meta.get("operation_type", "")),
                "document_number": str(document_meta.get("document_number", "")),
                "document_key": str(document_meta.get("document_key", "")),
                "document_date": str(document_meta.get("document_date", "")),
                "item_number": str(xml_item.get("item", "")),
                "code": str(xml_item.get("codigo", "")),
                "description": str(xml_item.get("descricao", "")),
                "ncm": str(xml_item.get("ncm", "")),
                "cst_icms": compose_xml_icms_cst_for_sped(xml_item),
                "cfop": str(xml_item.get("cfop", "")),
                "quantity": xml_item.get("quantidade", Decimal("0")),
                "icms_rate": xml_item.get("aliquota_icms", Decimal("0")),
                "sale_value": xml_item.get("valor_produto", Decimal("0")),
                "base_icms": xml_item.get("base_icms", Decimal("0")),
                "icms_value": xml_item.get("valor_icms", Decimal("0")),
                "base_icms_st": xml_item.get("base_icms_st", Decimal("0")),
                "icms_st_rate": xml_item.get("aliquota_icms_st", Decimal("0")),
                "icms_st_value": xml_item.get("valor_icms_st", Decimal("0")),
                "ipi_value": xml_item.get("valor_ipi", Decimal("0")),
            }
            xml_rebuilt_details.append(rebuilt_item)
            adjustment_rows.append(
                {
                    "operation_type": rebuilt_item["operation_type"],
                    "document_number": rebuilt_item["document_number"],
                    "document_key": rebuilt_item["document_key"],
                    "item_number": rebuilt_item["item_number"],
                    "code": rebuilt_item["code"],
                    "ncm_original": "",
                    "ncm_xml": rebuilt_item["ncm"],
                    "cfop_original": "",
                    "cfop_xml": rebuilt_item["cfop"],
                    "cst_original": "",
                    "cst_xml": rebuilt_item["cst_icms"],
                    "vl_item_original": Decimal("0"),
                    "vl_item_xml": rebuilt_item["sale_value"],
                    "vl_bc_original": Decimal("0"),
                    "vl_bc_xml": rebuilt_item["base_icms"],
                    "aliq_original": Decimal("0"),
                    "aliq_xml": rebuilt_item["icms_rate"],
                    "vl_icms_original": Decimal("0"),
                    "vl_icms_xml": rebuilt_item["icms_value"],
                    "vl_bc_st_original": Decimal("0"),
                    "vl_bc_st_xml": rebuilt_item["base_icms_st"],
                    "vl_icms_st_original": Decimal("0"),
                    "vl_icms_st_xml": rebuilt_item["icms_st_value"],
                    "vl_ipi_original": Decimal("0"),
                    "vl_ipi_xml": rebuilt_item["ipi_value"],
                }
            )

    preserved_sales = [
        dict(item)
        for item in detailed_sales
        if (
            str(item.get("operation_type", "")),
            normalize_document_key(str(item.get("document_key", ""))),
        ) not in affected_documents
    ]
    preserved_sales.extend(xml_rebuilt_details)
    return preserved_sales, affected_documents, adjustment_rows

def build_xml_st_index(
    xml_sources: list[Path],
    candidate_document_keys: set[str] | None = None,
) -> dict[str, dict[tuple[str, str], list[dict[str, object]]]]:
    if not xml_sources:
        return {}

    xml_files: list[Path] = []
    for xml_source in xml_sources:
        if not xml_source.exists():
            raise ValueError(f"O caminho de XML informado nao existe: {xml_source}")
        if xml_source.is_file():
            xml_files.append(xml_source)
        else:
            xml_files.extend(sorted(xml_source.rglob("*.xml")))

    xml_index: dict[str, dict[tuple[str, str], list[dict[str, object]]]] = {}
    for file_path in xml_files:
        if file_path.suffix.lower() != ".xml":
            continue
        if candidate_document_keys:
            document_key_hint = read_xml_document_key(file_path)
            if not document_key_hint or document_key_hint not in candidate_document_keys:
                continue
        try:
            document_key, items = parse_nfe_xml_st_items(file_path)
        except (ET.ParseError, UnicodeDecodeError):
            continue
        if document_key and items:
            document_bucket = xml_index.setdefault(document_key, {})
            for item_key, item_values in items.items():
                document_bucket.setdefault(item_key, []).extend(item_values)
    return xml_index

def apply_xml_st_adjustments_to_details(
    detailed_sales: list[dict[str, object]],
    xml_sources: list[Path],
) -> tuple[list[dict[str, object]], set[tuple[str, str]], list[dict[str, object]], list[dict[str, object]]]:
    candidate_document_keys = collect_xml_candidate_document_keys(detailed_sales)
    xml_index = build_xml_st_index(xml_sources, candidate_document_keys)
    if not xml_index:
        missing_rows = [
            {
                "operation_type": str(item.get("operation_type", "")).strip(),
                "document_number": str(item.get("document_number", "")).strip(),
                "document_key": normalize_document_key(str(item.get("document_key", ""))),
                "document_date": str(item.get("document_date", "")).strip(),
                "item_number": str(item.get("item_number", "")).strip(),
                "code": str(item.get("code", "")).strip(),
                "description": str(item.get("description", "")).strip(),
                "cfop_original": str(item.get("original_cfop", item.get("cfop", ""))).strip(),
                "cfop_ajustado": str(item.get("cfop", "")).strip(),
                "missing_reason": "Nenhum XML selecionado com chave correspondente.",
            }
            for item in detailed_sales
            if str(item.get("operation_type", "")).strip() == "Entrada"
            and str(item.get("original_cfop", item.get("cfop", ""))).strip() in {"1401", "1403"}
        ]
        return list(detailed_sales), set(), [], missing_rows

    adjusted_sales: list[dict[str, object]] = []
    affected_documents: set[tuple[str, str]] = set()
    adjustment_rows: list[dict[str, object]] = []
    missing_rows: list[dict[str, object]] = []
    fallback_usage: dict[str, set[tuple[str, str, str]]] = defaultdict(set)

    for item in detailed_sales:
        rebuilt_item = dict(item)
        operation_type = str(rebuilt_item.get("operation_type", "")).strip()
        document_key = normalize_document_key(str(rebuilt_item.get("document_key", "")))
        cfop = str(rebuilt_item.get("cfop", "")).strip()
        original_cfop = str(rebuilt_item.get("original_cfop", cfop)).strip()
        item_number = str(rebuilt_item.get("item_number", "")).strip()
        target_cfop = original_cfop if original_cfop in {"1401", "1403"} else cfop
        if operation_type != "Entrada" or target_cfop not in {"1401", "1403"} or not document_key:
            adjusted_sales.append(rebuilt_item)
            continue

        document_items = xml_index.get(document_key)
        if not document_items:
            missing_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(rebuilt_item.get("document_number", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(rebuilt_item.get("document_date", "")).strip(),
                    "item_number": item_number,
                    "code": str(rebuilt_item.get("code", "")).strip(),
                    "description": str(rebuilt_item.get("description", "")).strip(),
                    "cfop_original": target_cfop,
                    "cfop_ajustado": cfop,
                    "missing_reason": "Chave do SPED nao encontrada nos XMLs selecionados.",
                }
            )
            adjusted_sales.append(rebuilt_item)
            continue

        code = normalize_document_key(str(rebuilt_item.get("code", "")).strip())
        xml_candidates = document_items.get((item_number, target_cfop), [])
        xml_item = None

        if xml_candidates:
            if code:
                xml_item = next(
                    (
                        candidate
                        for candidate in xml_candidates
                        if code in candidate.get("xml_codes", set())
                    ),
                    None,
                )
            if xml_item is None and len(xml_candidates) == 1:
                xml_item = xml_candidates[0]
            if xml_item is not None:
                fallback_usage[document_key].add((item_number, target_cfop, str(xml_item.get("code", "")).strip()))

        if xml_item is None and code:
            used_keys = fallback_usage[document_key]
            for candidate_key, candidate_items in document_items.items():
                if candidate_key[1] != target_cfop:
                    continue
                for candidate_item in candidate_items:
                    usage_key = (candidate_key[0], candidate_key[1], str(candidate_item.get("code", "")).strip())
                    if usage_key in used_keys:
                        continue
                    if code in candidate_item.get("xml_codes", set()):
                        xml_item = candidate_item
                        used_keys.add(usage_key)
                        break
                if xml_item is not None:
                    break

        if xml_item is None:
            missing_rows.append(
                {
                    "operation_type": operation_type,
                    "document_number": str(rebuilt_item.get("document_number", "")).strip(),
                    "document_key": document_key,
                    "document_date": str(rebuilt_item.get("document_date", "")).strip(),
                    "item_number": item_number,
                    "code": str(rebuilt_item.get("code", "")).strip(),
                    "description": str(rebuilt_item.get("description", "")).strip(),
                    "cfop_original": target_cfop,
                    "cfop_ajustado": cfop,
                    "missing_reason": "Chave encontrada, mas o item/CFOP 1401/1403 nao foi localizado no XML.",
                }
            )
            adjusted_sales.append(rebuilt_item)
            continue

        original_base_icms_st = rebuilt_item.get("base_icms_st", Decimal("0"))
        original_icms_st_rate = rebuilt_item.get("icms_st_rate", Decimal("0"))
        original_icms_st_value = rebuilt_item.get("icms_st_value", Decimal("0"))
        rebuilt_item["base_icms_st"] = xml_item["base_icms_st"]
        rebuilt_item["icms_st_rate"] = xml_item["icms_st_rate"]
        rebuilt_item["icms_st_value"] = xml_item["icms_st_value"]
        adjustment_rows.append(
            {
                "operation_type": operation_type,
                "document_number": str(rebuilt_item.get("document_number", "")),
                "document_key": document_key,
                "item_number": item_number,
                "code": str(rebuilt_item.get("code", "")).strip(),
                "cfop_sped": target_cfop,
                "cfop_ajustado": cfop,
                "cfop_xml": str(xml_item.get("xml_cfop", "")),
                "base_icms_st_original": original_base_icms_st,
                "icms_st_rate_original": original_icms_st_rate,
                "icms_st_value_original": original_icms_st_value,
                "base_icms_st_ajustado": xml_item["base_icms_st"],
                "icms_st_rate_ajustado": xml_item["icms_st_rate"],
                "icms_st_value_ajustado": xml_item["icms_st_value"],
            }
        )
        adjusted_sales.append(rebuilt_item)
        affected_documents.add((operation_type, document_key or str(rebuilt_item.get("document_number", ""))))

    return adjusted_sales, affected_documents, adjustment_rows, missing_rows
