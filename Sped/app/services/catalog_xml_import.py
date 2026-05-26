from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from app.parsers.compare_xml import describe_compare_ignored_xml, parse_compare_xml_file
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.services.path_selection import parse_selected_paths


def _normalize_name(value: object, fallback: str) -> str:
    text = str(value or "").strip()
    return text or fallback


def _collect_xml_files(raw_sources: str) -> list[Path]:
    files: list[Path] = []
    for source_path in parse_selected_paths(raw_sources):
        if source_path.is_file() and source_path.suffix.lower() == ".xml":
            files.append(source_path)
            continue
        if source_path.is_dir():
            files.extend(sorted(source_path.rglob("*.xml")))
    unique: list[Path] = []
    seen: set[str] = set()
    for xml_path in files:
        key = str(xml_path.resolve()).lower() if xml_path.exists() else str(xml_path).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(xml_path)
    return unique


def _digits_only(value: object) -> str:
    return "".join(char for char in str(value or "") if char.isdigit())


NUMERIC_COMPARE_FIELDS = {
    "aliquota_icms",
    "aliquota_pis",
    "aliquota_cofins",
    "bc_st",
    "valor_icms_st",
    "aliquota_icms_st",
}


def _normalize_text(value: object) -> str:
    return str(value or "").strip()


def _normalize_decimal_text(value: object) -> str:
    text = _normalize_text(value).replace(" ", "")
    if not text:
        return ""
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return str(Decimal(text).normalize())
    except (InvalidOperation, ValueError):
        return text


def _values_equal(field_name: str, left: object, right: object) -> bool:
    if field_name in NUMERIC_COMPARE_FIELDS:
        return _normalize_decimal_text(left) == _normalize_decimal_text(right)
    return _normalize_text(left) == _normalize_text(right)


def _change_reason(field_name: str, left: object, right: object) -> str:
    if field_name in NUMERIC_COMPARE_FIELDS:
        return "valor numerico diferente"
    return "texto diferente"


UPDATABLE_PRODUCT_FIELDS = [
    "descricao",
    "ncm",
    "cest",
    "c_classtrib",
    "cst_icms",
    "aliquota_icms",
    "cst_pis",
    "cst_cofins",
    "aliquota_pis",
    "aliquota_cofins",
    "bc_st",
    "valor_icms_st",
    "aliquota_icms_st",
]


@dataclass
class CatalogImportPreviewRow:
    xml_file: str
    xml_file_path: str
    empresa: str
    empresa_ie: str
    fornecedor: str
    fornecedor_ie: str
    regime_tributario: str
    codigo_fornecedor: str
    descricao: str
    ean: str
    ncm: str
    cest: str
    c_classtrib: str
    cst_icms: str
    aliquota_icms: str
    bc_st: str
    valor_icms_st: str
    aliquota_icms_st: str
    cst_pis: str
    cst_cofins: str
    aliquota_pis: str
    aliquota_cofins: str
    classificacao: str
    key_token: str
    company_exists: bool = False
    supplier_exists: bool = False
    exists: bool = False


def build_catalog_import_preview(
    repository: MysqlCadastroRepository,
    environment: str,
    source_value: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CatalogImportPreviewRow], dict[str, int]]:
    repository.ensure_schema()
    stats: dict[str, int] = {
        "xml_total": 0,
        "xml_valid": 0,
        "xml_ignored": 0,
        "products_previewed": 0,
        "products_skipped_without_key": 0,
    }
    ignored_rows: list[list[str]] = []
    rows: list[CatalogImportPreviewRow] = []
    xml_files = _collect_xml_files(source_value)
    stats["xml_total"] = len(xml_files)
    for file_index, xml_file in enumerate(xml_files, start=1):
        if progress_callback is not None:
            progress_callback(file_index, len(xml_files), f"Lendo XML {file_index}/{len(xml_files)}: {xml_file.name}")
        invoice = parse_compare_xml_file(xml_file)
        if invoice is None:
            stats["xml_ignored"] += 1
            ignored_rows.append([xml_file.name, str(xml_file), describe_compare_ignored_xml(xml_file)])
            continue
        if not invoice.items:
            stats["xml_ignored"] += 1
            ignored_rows.append([xml_file.name, str(xml_file), "XML sem itens para importacao de produtos."])
            continue
        stats["xml_valid"] += 1
        company_name = _normalize_name(invoice.recipient_name, "Empresa sem nome")
        supplier_name = _normalize_name(invoice.issuer_name, "Fornecedor sem nome")
        for item in invoice.items:
            icms_source = str(getattr(item, "icms_code_source", "") or "").strip().upper()
            regime_tributario = "SIMPLES_NACIONAL" if icms_source == "CSOSN" else "LUCRO_REAL_PRESUMIDO"
            ean = _digits_only(getattr(item, "ean", ""))
            code = str(getattr(item, "code", "") or "").strip()
            if not ean and not code:
                stats["products_skipped_without_key"] += 1
                continue
            cst_pis = str(getattr(item, "cst_pis", "") or "").strip()
            cst_cofins = str(getattr(item, "cst_cofins", "") or "").strip()
            aliquota_pis = str(getattr(item, "aliq_pis", "") or "0")
            aliquota_cofins = str(getattr(item, "aliq_cofins", "") or "0")
            rows.append(
                CatalogImportPreviewRow(
                    xml_file=xml_file.name,
                    xml_file_path=str(xml_file),
                    empresa=company_name,
                    empresa_ie=_digits_only(invoice.recipient_ie),
                    fornecedor=supplier_name,
                    fornecedor_ie=_digits_only(invoice.issuer_ie),
                    regime_tributario=regime_tributario,
                    codigo_fornecedor=code or ean,
                    descricao=str(getattr(item, "description", "") or "").strip(),
                    ean=ean,
                    ncm=str(getattr(item, "ncm", "") or "").strip(),
                    cest=str(getattr(item, "cest", "") or "").strip(),
                    c_classtrib=str(getattr(item, "c_classtrib", "") or "").strip(),
                    cst_icms=str(getattr(item, "cst_icms", "") or "").strip(),
                    aliquota_icms=str(getattr(item, "aliq_icms", "") or "0"),
                    bc_st=str(getattr(item, "vl_bc_icms_st", "") or "0"),
                    valor_icms_st=str(getattr(item, "vl_icms_st", "") or "0"),
                    aliquota_icms_st=str(getattr(item, "aliq_icms_st", "") or "0"),
                    cst_pis=cst_pis,
                    cst_cofins=cst_cofins,
                    aliquota_pis=aliquota_pis,
                    aliquota_cofins=aliquota_cofins,
                    classificacao="Revenda",
                    key_token=f"{_digits_only(invoice.recipient_cnpj)}|{_digits_only(invoice.issuer_cnpj)}|{code or ean}",
                )
            )
    stats["products_previewed"] = len(rows)

    company_cache: dict[tuple[str, str], int | None] = {}
    supplier_cache: dict[tuple[int, str, str], int | None] = {}
    for row in rows:
        company_key = (row.empresa.upper(), _digits_only(row.key_token.split("|", 2)[0]))
        if company_key not in company_cache:
            company_cache[company_key] = repository.find_company_id(environment, row.empresa, company_key[1])
        company_id = company_cache[company_key]
        if company_id is None:
            row.company_exists = False
            row.exists = False
            continue
        row.company_exists = True
        supplier_cnpj = _digits_only(row.key_token.split("|", 2)[1])
        supplier_key = (company_id, row.fornecedor.upper(), supplier_cnpj)
        if (company_id, row.fornecedor.upper()) not in supplier_cache:
            supplier_cache[(company_id, row.fornecedor.upper())] = repository.find_supplier_id(company_id, row.fornecedor, supplier_cnpj)
        supplier_id = supplier_cache[(company_id, row.fornecedor.upper())]
        if supplier_id is None:
            row.supplier_exists = False
            row.exists = False
            continue
        row.supplier_exists = True
        row.exists = repository.supplier_product_exists(
            supplier_id,
            row.ean,
            row.codigo_fornecedor,
        )
    stats["ignored_rows"] = ignored_rows  # type: ignore[assignment]
    return rows, stats


def import_catalogs_from_preview(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
    allow_update_existing: bool,
    progress_callback: Callable[[int, int, str], None] | None = None,
    update_fields: set[str] | None = None,
) -> dict[str, int]:
    repository.ensure_schema()
    default_type_id = repository.ensure_product_type(environment, "Revenda", "Classificacao padrao para importacao XML")
    stats: dict[str, int] = {
        "rows_total": len(preview_rows),
        "companies_processed": 0,
        "suppliers_processed": 0,
        "products_created": 0,
        "products_updated": 0,
        "products_skipped_existing": 0,
        "default_type_id": default_type_id,
    }
    skipped_rows: list[list[str]] = []
    company_cache: dict[tuple[str, str, str], int] = {}
    supplier_cache: dict[tuple[int, str], int] = {}
    selected_fields = set(update_fields or UPDATABLE_PRODUCT_FIELDS)
    for index, row in enumerate(preview_rows, start=1):
        if progress_callback is not None:
            progress_callback(index, len(preview_rows), f"Importando {row.xml_file} ({index}/{len(preview_rows)})")
        recipient_cnpj, issuer_cnpj, _key = row.key_token.split("|", 2)
        company_key = (row.empresa.upper(), recipient_cnpj, row.empresa_ie)
        if company_key not in company_cache:
            company_cache[company_key] = repository.ensure_company(environment, row.empresa, recipient_cnpj, row.empresa_ie)
            stats["companies_processed"] += 1
        company_id = company_cache[company_key]
        supplier_key = (company_id, row.fornecedor.upper())
        if supplier_key not in supplier_cache:
            supplier_cache[supplier_key] = repository.ensure_supplier(
                company_id,
                row.fornecedor,
                issuer_cnpj,
                row.fornecedor_ie,
                row.regime_tributario,
            )
            stats["suppliers_processed"] += 1
        supplier_id = supplier_cache[supplier_key]
        payload = {
            "tipo_produto_id": default_type_id,
            "codigo_fornecedor": row.codigo_fornecedor,
            "codigo_empresa": "",
            "descricao": row.descricao,
            "ean": row.ean,
            "ncm": row.ncm,
            "cest": row.cest,
            "c_classtrib": row.c_classtrib,
            "c_benef": "",
            "cst_icms": row.cst_icms,
            "aliquota_icms": row.aliquota_icms,
            "cst_ipi": "",
            "aliquota_ipi": "0",
            "cst_pis": row.cst_pis,
            "cst_cofins": row.cst_cofins,
            "aliquota_pis": row.aliquota_pis,
            "aliquota_cofins": row.aliquota_cofins,
            "cst_pis_cofins": f"{row.cst_pis}/{row.cst_cofins}".strip("/"),
            "aliquota_pis_cofins": row.aliquota_pis,
            "bc_st": row.bc_st,
            "mva": "0",
            "valor_icms_st": row.valor_icms_st,
            "aliquota_icms_st": row.aliquota_icms_st,
        }
        existing_product = repository.find_supplier_product_by_key(supplier_id, row.ean, row.codigo_fornecedor)
        if existing_product and allow_update_existing:
            for field_name in UPDATABLE_PRODUCT_FIELDS:
                if field_name not in selected_fields or _values_equal(field_name, existing_product.get(field_name, ""), payload.get(field_name, "")):
                    payload[field_name] = existing_product.get(field_name, payload[field_name])
        _product_id, existed = repository.upsert_supplier_product_by_ean(supplier_id, payload, allow_update_existing=allow_update_existing)
        if existed and not allow_update_existing:
            stats["products_skipped_existing"] += 1
            skipped_rows.append([row.xml_file, row.codigo_fornecedor, row.ean, row.descricao, "Ignorado", "Produto ja existe e atualizacao de existentes nao foi autorizada."])
            continue
        if existed:
            stats["products_updated"] += 1
            changed_fields = [
                field_name
                for field_name in UPDATABLE_PRODUCT_FIELDS
                if _values_equal((field_name), existing_product.get(field_name, "") if existing_product else "", payload.get(field_name, ""))
                is False
            ] if existing_product else []
            if not changed_fields:
                skipped_rows.append([row.xml_file, row.codigo_fornecedor, row.ean, row.descricao, "Sem alteracao", "Item existente sem diferenca real nos campos selecionados."])
        else:
            stats["products_created"] += 1
    stats["skipped_rows"] = skipped_rows  # type: ignore[assignment]
    return stats


def build_catalog_import_change_rows(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
) -> list[list[str]]:
    changes: list[list[str]] = []
    company_cache: dict[tuple[str, str], int | None] = {}
    supplier_cache: dict[tuple[int, str, str], int | None] = {}
    for row in preview_rows:
        recipient_cnpj, issuer_cnpj, _ = row.key_token.split("|", 2)
        company_key = (row.empresa.upper(), recipient_cnpj)
        if company_key not in company_cache:
            company_cache[company_key] = repository.find_company_id(environment, row.empresa, recipient_cnpj)
        company_id = company_cache[company_key]
        if not company_id:
            continue
        supplier_key = (company_id, row.fornecedor.upper(), issuer_cnpj)
        if supplier_key not in supplier_cache:
            supplier_cache[supplier_key] = repository.find_supplier_id(company_id, row.fornecedor, issuer_cnpj)
        supplier_id = supplier_cache[supplier_key]
        if not supplier_id:
            continue
        existing = repository.find_supplier_product_by_key(supplier_id, row.ean, row.codigo_fornecedor)
        if not existing:
            continue
        current_map = {
            "descricao": str(existing.get("descricao", "")),
            "ncm": str(existing.get("ncm", "")),
            "cest": str(existing.get("cest", "")),
            "c_classtrib": str(existing.get("c_classtrib", "")),
            "cst_icms": str(existing.get("cst_icms", "")),
            "aliquota_icms": str(existing.get("aliquota_icms", "")),
        }
        new_map = {
            "descricao": row.descricao,
            "ncm": row.ncm,
            "cest": row.cest,
            "c_classtrib": row.c_classtrib,
            "cst_icms": row.cst_icms,
            "aliquota_icms": row.aliquota_icms,
        }
        for field_name, old_value in current_map.items():
            new_value = new_map[field_name]
            if not _values_equal(field_name, old_value, new_value):
                changes.append([row.xml_file, row.codigo_fornecedor, field_name, str(old_value), str(new_value), _change_reason(field_name, old_value, new_value)])
    return changes


def build_catalog_import_duplicate_rows(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
) -> list[list[str]]:
    rows: list[list[str]] = []
    company_cache: dict[tuple[str, str], int | None] = {}
    supplier_cache: dict[tuple[int, str, str], int | None] = {}
    for row in preview_rows:
        recipient_cnpj, issuer_cnpj, _ = row.key_token.split("|", 2)
        company_key = (row.empresa.upper(), recipient_cnpj)
        if company_key not in company_cache:
            company_cache[company_key] = repository.find_company_id(environment, row.empresa, recipient_cnpj)
        company_id = company_cache[company_key]
        if not company_id:
            continue
        supplier_key = (company_id, row.fornecedor.upper(), issuer_cnpj)
        if supplier_key not in supplier_cache:
            supplier_cache[supplier_key] = repository.find_supplier_id(company_id, row.fornecedor, issuer_cnpj)
        supplier_id = supplier_cache[supplier_key]
        if not supplier_id:
            continue
        duplicates = repository.find_supplier_product_duplicates(supplier_id, row.ean, row.codigo_fornecedor)
        for dup in duplicates:
            rows.append(
                [
                    row.xml_file,
                    str(supplier_id),
                    row.codigo_fornecedor,
                    row.ean,
                    str(dup.get("id", "")),
                    str(dup.get("descricao", "")),
                    "chave duplicada no cadastro",
                ]
            )
    unique = []
    seen = set()
    for item in rows:
        key = tuple(item)
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    return unique
