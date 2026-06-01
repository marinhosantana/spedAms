from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from app.parsers.compare_xml import describe_compare_ignored_xml, parse_compare_xml_file
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.services.path_selection import parse_selected_paths


MAX_SQL_IN_ITEMS = 800
MAX_IMPORT_REVIEW_DETAIL_ROWS = 2000


def _chunks(values: list[object], size: int = MAX_SQL_IN_ITEMS) -> list[list[object]]:
    return [values[index:index + size] for index in range(0, len(values), size)]


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
    "reducao_bc_icms",
    "aliquota_ipi",
    "aliquota_pis",
    "aliquota_cofins",
    "aliquota_pis_cofins",
    "bc_st",
    "mva",
    "valor_icms_st",
    "aliquota_icms_st",
}


NUMERIC_UPDATE_FIELDS = set(NUMERIC_COMPARE_FIELDS)


TEXT_DIGITS_UPDATE_FIELDS = {
    "ncm",
    "cest",
    "cfop_saida_fornecedor",
}


TEXT_LIMITS_UPDATE_FIELDS = {
    "chave_nfe_origem": 44,
    "descricao": 255,
    "origem_entrada": 4,
    "c_classtrib": 20,
    "c_benef": 20,
    "cst_icms": 4,
    "cst_ipi": 4,
    "cst_pis": 4,
    "cst_cofins": 4,
    "cst_pis_cofins": 20,
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


def _placeholders(values: object) -> str:
    return ", ".join(["%s"] * len(values))


def _mark_preview_existing_rows(
    repository: MysqlCadastroRepository,
    environment: str,
    rows: list[CatalogImportPreviewRow],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> None:
    if not rows:
        return
    if progress_callback is not None:
        progress_callback(0, 3, "Conferindo cadastros existentes...")

    company_cnpjs = {_digits_only(row.key_token.split("|", 2)[0]) for row in rows}
    company_cnpjs.discard("")
    company_names = {row.empresa.upper() for row in rows if row.empresa.strip()}
    company_by_cnpj: dict[str, int] = {}
    company_by_name: dict[str, int] = {}

    connection = repository.get_connection()
    try:
        cursor = connection.cursor(dictionary=True)
        company_conditions: list[str] = []
        company_params: list[object] = [environment]
        if company_cnpjs:
            company_conditions.append(f"cnpj IN ({_placeholders(company_cnpjs)})")
            company_params.extend(sorted(company_cnpjs))
        if company_names:
            company_conditions.append(f"UPPER(nome) IN ({_placeholders(company_names)})")
            company_params.extend(sorted(company_names))
        if company_conditions:
            cursor.execute(
                f"""
                SELECT id, cnpj, UPPER(nome) AS nome_key
                FROM cad_empresas
                WHERE ambiente = %s
                  AND ({' OR '.join(company_conditions)})
                """,
                company_params,
            )
            for company in cursor.fetchall():
                company_id = int(company["id"])
                cnpj = _digits_only(company.get("cnpj", ""))
                name = str(company.get("nome_key", "") or "").strip().upper()
                if cnpj and cnpj not in company_by_cnpj:
                    company_by_cnpj[cnpj] = company_id
                if name and name not in company_by_name:
                    company_by_name[name] = company_id

        company_id_by_row: dict[int, int] = {}
        for row in rows:
            company_cnpj = _digits_only(row.key_token.split("|", 2)[0])
            company_id = company_by_cnpj.get(company_cnpj) or company_by_name.get(row.empresa.upper())
            if company_id:
                row.company_exists = True
                company_id_by_row[id(row)] = company_id
            else:
                row.company_exists = False
                row.supplier_exists = False
                row.exists = False

        if progress_callback is not None:
            progress_callback(1, 3, "Conferindo fornecedores existentes...")

        company_ids = set(company_id_by_row.values())
        supplier_cnpjs = {_digits_only(row.key_token.split("|", 2)[1]) for row in rows if id(row) in company_id_by_row}
        supplier_cnpjs.discard("")
        supplier_names = {row.fornecedor.upper() for row in rows if id(row) in company_id_by_row and row.fornecedor.strip()}
        supplier_by_key: dict[tuple[int, str], int] = {}
        if company_ids and (supplier_cnpjs or supplier_names):
            supplier_conditions: list[str] = []
            supplier_params: list[object] = sorted(company_ids)
            if supplier_cnpjs:
                supplier_conditions.append(f"cnpj IN ({_placeholders(supplier_cnpjs)})")
                supplier_params.extend(sorted(supplier_cnpjs))
            if supplier_names:
                supplier_conditions.append(f"UPPER(nome) IN ({_placeholders(supplier_names)})")
                supplier_params.extend(sorted(supplier_names))
            cursor.execute(
                f"""
                SELECT id, empresa_id, cnpj, UPPER(nome) AS nome_key
                FROM cad_fornecedores
                WHERE empresa_id IN ({_placeholders(company_ids)})
                  AND ({' OR '.join(supplier_conditions)})
                """,
                supplier_params,
            )
            for supplier in cursor.fetchall():
                supplier_id = int(supplier["id"])
                company_id = int(supplier["empresa_id"])
                cnpj = _digits_only(supplier.get("cnpj", ""))
                name = str(supplier.get("nome_key", "") or "").strip().upper()
                if cnpj:
                    supplier_by_key.setdefault((company_id, f"cnpj:{cnpj}"), supplier_id)
                if name:
                    supplier_by_key.setdefault((company_id, f"name:{name}"), supplier_id)

        supplier_id_by_row: dict[int, int] = {}
        for row in rows:
            company_id = company_id_by_row.get(id(row))
            if not company_id:
                continue
            supplier_cnpj = _digits_only(row.key_token.split("|", 2)[1])
            supplier_id = supplier_by_key.get((company_id, f"cnpj:{supplier_cnpj}")) or supplier_by_key.get((company_id, f"name:{row.fornecedor.upper()}"))
            if supplier_id:
                row.supplier_exists = True
                supplier_id_by_row[id(row)] = supplier_id
            else:
                row.supplier_exists = False
                row.exists = False

        if progress_callback is not None:
            progress_callback(2, 3, "Conferindo produtos existentes...")

        supplier_ids = set(supplier_id_by_row.values())
        eans = {row.ean for row in rows if id(row) in supplier_id_by_row and row.ean}
        codes = {row.codigo_fornecedor for row in rows if id(row) in supplier_id_by_row and row.codigo_fornecedor}
        existing_products: set[tuple[int, str, str]] = set()
        if supplier_ids and (eans or codes):
            product_conditions: list[str] = []
            product_params: list[object] = sorted(supplier_ids)
            if eans:
                product_conditions.append(f"ean IN ({_placeholders(eans)})")
                product_params.extend(sorted(eans))
            if codes:
                product_conditions.append(f"codigo_fornecedor IN ({_placeholders(codes)})")
                product_params.extend(sorted(codes))
            cursor.execute(
                f"""
                SELECT fornecedor_id, ean, codigo_fornecedor
                FROM cad_produtos_fornecedor
                WHERE fornecedor_id IN ({_placeholders(supplier_ids)})
                  AND ({' OR '.join(product_conditions)})
                """,
                product_params,
            )
            for product in cursor.fetchall():
                existing_products.add(
                    (
                        int(product["fornecedor_id"]),
                        _digits_only(product.get("ean", "")),
                        str(product.get("codigo_fornecedor", "") or "").strip(),
                    )
                )

        for row in rows:
            supplier_id = supplier_id_by_row.get(id(row))
            if not supplier_id:
                continue
            row.exists = (supplier_id, row.ean, row.codigo_fornecedor) in existing_products
            if not row.exists and row.ean:
                row.exists = any(item[0] == supplier_id and item[1] == row.ean for item in existing_products)
            if not row.exists and row.codigo_fornecedor:
                row.exists = any(item[0] == supplier_id and item[2] == row.codigo_fornecedor for item in existing_products)
    finally:
        connection.close()

    if progress_callback is not None:
        progress_callback(3, 3, "Conferencia de cadastros concluida.")


def _load_existing_products_for_preview(
    repository: MysqlCadastroRepository,
    environment: str,
    rows: list[CatalogImportPreviewRow],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[
    list[tuple[CatalogImportPreviewRow, int]],
    dict[tuple[int, str], list[dict[str, object]]],
    dict[tuple[int, str], list[dict[str, object]]],
]:
    if not rows:
        return [], {}, {}
    if progress_callback is not None:
        progress_callback(0, 4, "Localizando empresas da previa...")

    company_cnpjs = {_digits_only(row.key_token.split("|", 2)[0]) for row in rows}
    company_cnpjs.discard("")
    company_names = {row.empresa.upper() for row in rows if row.empresa.strip()}
    company_by_cnpj: dict[str, int] = {}
    company_by_name: dict[str, int] = {}

    connection = repository.get_connection()
    try:
        cursor = connection.cursor(dictionary=True)
        company_conditions: list[str] = []
        company_params: list[object] = [environment]
        if company_cnpjs:
            company_conditions.append(f"cnpj IN ({_placeholders(company_cnpjs)})")
            company_params.extend(sorted(company_cnpjs))
        if company_names:
            company_conditions.append(f"UPPER(nome) IN ({_placeholders(company_names)})")
            company_params.extend(sorted(company_names))
        if company_conditions:
            cursor.execute(
                f"""
                SELECT id, cnpj, UPPER(nome) AS nome_key
                FROM cad_empresas
                WHERE ambiente = %s
                  AND ({' OR '.join(company_conditions)})
                """,
                company_params,
            )
            for company in cursor.fetchall():
                company_id = int(company["id"])
                cnpj = _digits_only(company.get("cnpj", ""))
                name = str(company.get("nome_key", "") or "").strip().upper()
                if cnpj:
                    company_by_cnpj.setdefault(cnpj, company_id)
                if name:
                    company_by_name.setdefault(name, company_id)

        company_id_by_row: dict[int, int] = {}
        for row in rows:
            company_cnpj = _digits_only(row.key_token.split("|", 2)[0])
            company_id = company_by_cnpj.get(company_cnpj) or company_by_name.get(row.empresa.upper())
            if company_id:
                company_id_by_row[id(row)] = company_id

        if progress_callback is not None:
            progress_callback(1, 4, "Localizando fornecedores da previa...")

        company_ids = set(company_id_by_row.values())
        supplier_cnpjs = {_digits_only(row.key_token.split("|", 2)[1]) for row in rows if id(row) in company_id_by_row}
        supplier_cnpjs.discard("")
        supplier_names = {row.fornecedor.upper() for row in rows if id(row) in company_id_by_row and row.fornecedor.strip()}
        supplier_by_key: dict[tuple[int, str], int] = {}
        if company_ids and (supplier_cnpjs or supplier_names):
            supplier_conditions: list[str] = []
            supplier_params: list[object] = sorted(company_ids)
            if supplier_cnpjs:
                supplier_conditions.append(f"cnpj IN ({_placeholders(supplier_cnpjs)})")
                supplier_params.extend(sorted(supplier_cnpjs))
            if supplier_names:
                supplier_conditions.append(f"UPPER(nome) IN ({_placeholders(supplier_names)})")
                supplier_params.extend(sorted(supplier_names))
            cursor.execute(
                f"""
                SELECT id, empresa_id, cnpj, UPPER(nome) AS nome_key
                FROM cad_fornecedores
                WHERE empresa_id IN ({_placeholders(company_ids)})
                  AND ({' OR '.join(supplier_conditions)})
                """,
                supplier_params,
            )
            for supplier in cursor.fetchall():
                supplier_id = int(supplier["id"])
                company_id = int(supplier["empresa_id"])
                cnpj = _digits_only(supplier.get("cnpj", ""))
                name = str(supplier.get("nome_key", "") or "").strip().upper()
                if cnpj:
                    supplier_by_key.setdefault((company_id, f"cnpj:{cnpj}"), supplier_id)
                if name:
                    supplier_by_key.setdefault((company_id, f"name:{name}"), supplier_id)

        resolved_rows: list[tuple[CatalogImportPreviewRow, int]] = []
        for row in rows:
            company_id = company_id_by_row.get(id(row))
            if not company_id:
                continue
            supplier_cnpj = _digits_only(row.key_token.split("|", 2)[1])
            supplier_id = supplier_by_key.get((company_id, f"cnpj:{supplier_cnpj}")) or supplier_by_key.get((company_id, f"name:{row.fornecedor.upper()}"))
            if supplier_id:
                resolved_rows.append((row, supplier_id))

        if progress_callback is not None:
            progress_callback(2, 4, "Carregando produtos existentes em lote...")

        supplier_ids = {supplier_id for _row, supplier_id in resolved_rows}
        eans = {row.ean for row, _supplier_id in resolved_rows if row.ean}
        codes = {row.codigo_fornecedor for row, _supplier_id in resolved_rows if row.codigo_fornecedor}
        products_by_ean: dict[tuple[int, str], list[dict[str, object]]] = {}
        products_by_code: dict[tuple[int, str], list[dict[str, object]]] = {}
        if supplier_ids and (eans or codes):
            loaded_product_ids: set[int] = set()

            def add_loaded_products(product_rows: list[dict[str, object]]) -> None:
                for product_row in product_rows:
                    product = dict(product_row)
                    product_id = int(product.get("id") or 0)
                    if product_id in loaded_product_ids:
                        continue
                    loaded_product_ids.add(product_id)
                    supplier_id = int(product["fornecedor_id"])
                    ean = repository._only_digits(product.get("ean", ""))
                    code = str(product.get("codigo_fornecedor", "") or "").strip()
                    if ean:
                        products_by_ean.setdefault((supplier_id, ean), []).append(product)
                    if code:
                        products_by_code.setdefault((supplier_id, code), []).append(product)

            supplier_id_list = sorted(supplier_ids)
            for supplier_chunk in _chunks(supplier_id_list):
                for ean_chunk in _chunks(sorted(eans)):
                    cursor.execute(
                        f"""
                        SELECT *
                        FROM cad_produtos_fornecedor
                        WHERE fornecedor_id IN ({_placeholders(supplier_chunk)})
                          AND ean IN ({_placeholders(ean_chunk)})
                        ORDER BY id
                        """,
                        [*supplier_chunk, *ean_chunk],
                    )
                    add_loaded_products([dict(row) for row in cursor.fetchall()])
                for code_chunk in _chunks(sorted(codes)):
                    cursor.execute(
                        f"""
                        SELECT *
                        FROM cad_produtos_fornecedor
                        WHERE fornecedor_id IN ({_placeholders(supplier_chunk)})
                          AND codigo_fornecedor IN ({_placeholders(code_chunk)})
                        ORDER BY id
                        """,
                        [*supplier_chunk, *code_chunk],
                    )
                    add_loaded_products([dict(row) for row in cursor.fetchall()])
    finally:
        connection.close()

    if progress_callback is not None:
        progress_callback(3, 4, "Produtos existentes carregados.")
    return resolved_rows, products_by_ean, products_by_code


def collect_catalog_import_existing_product_ids(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[int]:
    resolved_rows, products_by_ean, products_by_code = _load_existing_products_for_preview(
        repository,
        environment,
        preview_rows,
        progress_callback,
    )
    product_ids: set[int] = set()
    for row, supplier_id in resolved_rows:
        products = products_by_ean.get((supplier_id, row.ean), []) if row.ean else []
        if not products and row.codigo_fornecedor:
            products = products_by_code.get((supplier_id, row.codigo_fornecedor), [])
        for product in products:
            product_id = int(product.get("id") or 0)
            if product_id:
                product_ids.add(product_id)
    return sorted(product_ids)


def build_catalog_import_existing_review(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
    update_fields: set[str] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[list[str]], list[list[str]], list[int]]:
    selected_fields = set(update_fields or UPDATABLE_PRODUCT_FIELDS)
    resolved_rows, products_by_ean, products_by_code = _load_existing_products_for_preview(
        repository,
        environment,
        preview_rows,
        progress_callback,
    )
    changes: list[list[str]] = []
    duplicate_rows: list[list[str]] = []
    product_ids: set[int] = set()
    seen_duplicates: set[tuple[object, ...]] = set()
    hidden_changes = 0
    hidden_duplicates = 0
    total = len(resolved_rows)
    for index, (row, supplier_id) in enumerate(resolved_rows, start=1):
        if progress_callback is not None and (index == 1 or index == total or index % 100 == 0):
            progress_callback(index, total, f"Conferindo existentes ({index}/{total})")
        products = products_by_ean.get((supplier_id, row.ean), []) if row.ean else []
        if not products and row.codigo_fornecedor:
            products = products_by_code.get((supplier_id, row.codigo_fornecedor), [])
        if not products:
            continue
        for product in products:
            product_id = int(product.get("id") or 0)
            if product_id:
                product_ids.add(product_id)
        duplicate_products = products
        if len(duplicate_products) <= 1 and row.codigo_fornecedor:
            duplicate_products = products_by_code.get((supplier_id, row.codigo_fornecedor), [])
        if len(duplicate_products) > 1:
            for dup in duplicate_products:
                duplicate_item = (
                    row.xml_file,
                    str(supplier_id),
                    row.codigo_fornecedor,
                    row.ean,
                    str(dup.get("id", "")),
                    str(dup.get("descricao", "")),
                    "chave duplicada no cadastro",
                )
                if duplicate_item not in seen_duplicates:
                    seen_duplicates.add(duplicate_item)
                    if len(duplicate_rows) < MAX_IMPORT_REVIEW_DETAIL_ROWS:
                        duplicate_rows.append(list(duplicate_item))
                    else:
                        hidden_duplicates += 1
        existing = products[0]
        new_map = _payload_from_preview_row(row)
        for field_name in UPDATABLE_PRODUCT_FIELDS:
            if field_name not in selected_fields:
                continue
            old_value = existing.get(field_name, "")
            new_value = new_map.get(field_name, "")
            if _should_ignore_blank_text_update(field_name, old_value, new_value):
                continue
            if not _values_equal(field_name, old_value, new_value):
                if len(changes) < MAX_IMPORT_REVIEW_DETAIL_ROWS:
                    changes.append([row.xml_file, row.codigo_fornecedor, field_name, str(old_value), str(new_value), _change_reason(field_name, old_value, new_value)])
                else:
                    hidden_changes += 1
    if hidden_duplicates:
        duplicate_rows.append(["...", "", "", "", "", "", f"{hidden_duplicates} duplicidade(s) omitida(s) para preservar desempenho da tela."])
    if hidden_changes:
        changes.append(["...", "", "", "", "", f"{hidden_changes} mudanca(s) omitida(s) para preservar desempenho da tela."])
    return changes, duplicate_rows, sorted(product_ids)


def _supplier_product_values(repository: MysqlCadastroRepository, supplier_id: int, data: dict[str, object]) -> tuple[object, ...]:
    cst_pis = repository._trim_text(data.get("cst_pis", data.get("cst_pis_cofins", "")), 4)
    cst_cofins = repository._trim_text(data.get("cst_cofins", data.get("cst_pis_cofins", "")), 4)
    aliquota_pis = repository._decimal_text(data.get("aliquota_pis", data.get("aliquota_pis_cofins", "")))
    aliquota_cofins = repository._decimal_text(data.get("aliquota_cofins", data.get("aliquota_pis_cofins", "")))
    cst_pis_cofins = repository._trim_text(data.get("cst_pis_cofins", ""), 20)
    aliquota_pis_cofins = repository._decimal_text(data.get("aliquota_pis_cofins", aliquota_pis))
    normalized_ean = repository._only_digits(data.get("ean", ""))[:30]
    return (
        supplier_id,
        int(data.get("tipo_produto_id") or 0) or None,
        repository._trim_text(data.get("codigo_fornecedor", ""), 80),
        repository._trim_text(data.get("codigo_empresa", ""), 80),
        repository._trim_text(data.get("chave_nfe_origem", ""), 44),
        repository._trim_text(data.get("descricao", ""), 255),
        normalized_ean,
        normalized_ean or None,
        repository._only_digits(data.get("ncm", ""))[:20],
        repository._only_digits(data.get("cest", ""))[:20],
        repository._trim_text(data.get("origem_entrada", ""), 4),
        repository._only_digits(data.get("cfop_saida_fornecedor", ""))[:10],
        repository._only_digits(data.get("cfop_entrada", ""))[:10],
        repository._only_digits(data.get("cfop_saida", ""))[:10],
        repository._trim_text(data.get("origem_saida", ""), 4),
        repository._trim_text(data.get("c_classtrib", ""), 20),
        repository._trim_text(data.get("c_benef", ""), 20),
        repository._trim_text(data.get("cst_icms", ""), 4),
        repository._decimal_text(data.get("reducao_bc_icms", "")),
        repository._decimal_text(data.get("aliquota_icms", "")),
        repository._trim_text(data.get("cst_ipi", ""), 4),
        repository._decimal_text(data.get("aliquota_ipi", "")),
        cst_pis_cofins,
        aliquota_pis_cofins,
        cst_pis,
        cst_cofins,
        aliquota_pis,
        aliquota_cofins,
        repository._decimal_text(data.get("bc_st", "")),
        repository._decimal_text(data.get("mva", "")),
        repository._decimal_text(data.get("valor_icms_st", "")),
        repository._decimal_text(data.get("aliquota_icms_st", "")),
    )


def _payload_from_preview_row(row: CatalogImportPreviewRow) -> dict[str, object]:
    cfop_xml = str(row.cfop_saida_fornecedor or "").strip()
    origem_xml = str(row.origem_entrada or "").strip()
    return {
        "tipo_produto_id": None,
        "chave_nfe_origem": row.chave_nfe_origem,
        "codigo_fornecedor": row.codigo_fornecedor,
        "codigo_empresa": "",
        "descricao": row.descricao,
        "ean": row.ean,
        "ncm": row.ncm,
        "cest": row.cest,
        "origem_entrada": origem_xml,
        "cfop_saida_fornecedor": cfop_xml,
        "cfop_entrada": "",
        "cfop_saida": "",
        "origem_saida": origem_xml,
        "c_classtrib": row.c_classtrib,
        "c_benef": "",
        "cst_icms": row.cst_icms,
        "reducao_bc_icms": row.reducao_bc_icms,
        "aliquota_icms": row.aliquota_icms,
        "cst_ipi": "",
        "aliquota_ipi": "0",
        "cst_pis": row.cst_pis,
        "cst_cofins": row.cst_cofins,
        "aliquota_pis": row.aliquota_pis,
        "aliquota_cofins": row.aliquota_cofins,
        "cst_pis_cofins": "",
        "aliquota_pis_cofins": row.aliquota_pis,
        "bc_st": row.bc_st,
        "mva": row.mva,
        "valor_icms_st": row.valor_icms_st,
        "aliquota_icms_st": row.aliquota_icms_st,
    }


def _normalize_product_update_value(repository: MysqlCadastroRepository, field_name: str, value: object) -> object:
    if field_name in NUMERIC_UPDATE_FIELDS:
        return repository._decimal_text(value)
    if field_name in TEXT_DIGITS_UPDATE_FIELDS:
        return repository._only_digits(value)[:10 if field_name.startswith("cfop") else 20]
    if field_name in TEXT_LIMITS_UPDATE_FIELDS:
        return repository._trim_text(value, TEXT_LIMITS_UPDATE_FIELDS[field_name])
    return value


def _should_ignore_blank_text_update(field_name: str, current_value: object, new_value: object) -> bool:
    if field_name in NUMERIC_UPDATE_FIELDS:
        return False
    return bool(str(current_value or "").strip()) and not str(new_value or "").strip()


def _change_reason(field_name: str, left: object, right: object) -> str:
    if field_name in NUMERIC_COMPARE_FIELDS:
        return "valor numerico diferente"
    return "texto diferente"


UPDATABLE_PRODUCT_FIELDS = [
    "descricao",
    "chave_nfe_origem",
    "ncm",
    "cest",
    "origem_entrada",
    "cfop_saida_fornecedor",
    "c_classtrib",
    "c_benef",
    "cst_icms",
    "reducao_bc_icms",
    "aliquota_icms",
    "cst_ipi",
    "aliquota_ipi",
    "cst_pis",
    "cst_cofins",
    "cst_pis_cofins",
    "aliquota_pis_cofins",
    "aliquota_pis",
    "aliquota_cofins",
    "bc_st",
    "mva",
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
    fornecedor_uf: str
    regime_tributario: str
    chave_nfe_origem: str
    codigo_fornecedor: str
    descricao: str
    ean: str
    ncm: str
    cest: str
    c_classtrib: str
    cst_icms: str
    cfop_saida_fornecedor: str
    origem_entrada: str
    reducao_bc_icms: str
    aliquota_icms: str
    bc_st: str
    valor_icms_st: str
    aliquota_icms_st: str
    mva: str
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
                    fornecedor_uf=str(getattr(invoice, "issuer_uf", "") or "").strip().upper()[:2],
                    regime_tributario=regime_tributario,
                    chave_nfe_origem=str(invoice.key or "").strip(),
                    codigo_fornecedor=code or ean,
                    descricao=str(getattr(item, "description", "") or "").strip(),
                    ean=ean,
                    ncm=str(getattr(item, "ncm", "") or "").strip(),
                    cest=str(getattr(item, "cest", "") or "").strip(),
                    c_classtrib=str(getattr(item, "c_classtrib", "") or "").strip(),
                    cst_icms=str(getattr(item, "cst_icms", "") or "").strip(),
                    cfop_saida_fornecedor=str(getattr(item, "cfop", "") or "").strip(),
                    origem_entrada=str(getattr(item, "orig_icms", "") or "").strip(),
                    reducao_bc_icms=str(getattr(item, "reducao_bc_icms", "") or "0"),
                    aliquota_icms=str(getattr(item, "aliq_icms", "") or "0"),
                    bc_st=str(getattr(item, "vl_bc_icms_st", "") or "0"),
                    valor_icms_st=str(getattr(item, "vl_icms_st", "") or "0"),
                    aliquota_icms_st=str(getattr(item, "aliq_icms_st", "") or "0"),
                    mva=str(getattr(item, "mva", "") or "0"),
                    cst_pis=cst_pis,
                    cst_cofins=cst_cofins,
                    aliquota_pis=aliquota_pis,
                    aliquota_cofins=aliquota_cofins,
                    classificacao="",
                    key_token=f"{_digits_only(invoice.recipient_cnpj)}|{_digits_only(invoice.issuer_cnpj)}|{code or ean}",
                )
            )
    stats["products_previewed"] = len(rows)

    _mark_preview_existing_rows(repository, environment, rows, progress_callback)
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
    stats: dict[str, int] = {
        "rows_total": len(preview_rows),
        "companies_processed": 0,
        "suppliers_processed": 0,
        "products_created": 0,
        "products_updated": 0,
        "products_skipped_existing": 0,
    }
    skipped_rows: list[list[str]] = []
    hidden_skipped_rows = 0

    def add_skipped_row(row_data: list[str]) -> None:
        nonlocal hidden_skipped_rows
        if len(skipped_rows) < MAX_IMPORT_REVIEW_DETAIL_ROWS:
            skipped_rows.append(row_data)
        else:
            hidden_skipped_rows += 1

    company_cache: dict[tuple[str, str, str], int] = {}
    supplier_cache: dict[tuple[int, str, str], int] = {}
    selected_fields = set(update_fields or UPDATABLE_PRODUCT_FIELDS)
    resolved_rows: list[tuple[CatalogImportPreviewRow, int, dict[str, object]]] = []
    for index, row in enumerate(preview_rows, start=1):
        if progress_callback is not None and (index == 1 or index == len(preview_rows) or index % 50 == 0):
            progress_callback(index, len(preview_rows), f"Preparando cadastros {row.xml_file} ({index}/{len(preview_rows)})")
        recipient_cnpj, issuer_cnpj, _key = row.key_token.split("|", 2)
        company_key = (row.empresa.upper(), recipient_cnpj, row.empresa_ie)
        if company_key not in company_cache:
            company_cache[company_key] = repository.ensure_company(environment, row.empresa, recipient_cnpj, row.empresa_ie)
            stats["companies_processed"] += 1
        company_id = company_cache[company_key]
        supplier_key = (company_id, row.fornecedor.upper(), issuer_cnpj)
        if supplier_key not in supplier_cache:
            supplier_cache[supplier_key] = repository.ensure_supplier(
                company_id,
                row.fornecedor,
                issuer_cnpj,
                row.fornecedor_ie,
                row.fornecedor_uf,
                row.regime_tributario,
            )
            stats["suppliers_processed"] += 1
        supplier_id = supplier_cache[supplier_key]
        resolved_rows.append((row, supplier_id, _payload_from_preview_row(row)))

    supplier_ids = {supplier_id for _row, supplier_id, _payload in resolved_rows}
    eans = {str(payload.get("ean", "") or "").strip() for _row, _supplier_id, payload in resolved_rows if str(payload.get("ean", "") or "").strip()}
    codes = {str(payload.get("codigo_fornecedor", "") or "").strip() for _row, _supplier_id, payload in resolved_rows if str(payload.get("codigo_fornecedor", "") or "").strip()}
    existing_by_id: dict[int, dict[str, object]] = {}
    existing_by_ean: dict[tuple[int, str], int] = {}
    existing_by_code: dict[tuple[int, str], int] = {}
    if supplier_ids and (eans or codes):
        connection = repository.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            loaded_product_ids: set[int] = set()

            def add_existing_products(product_rows: list[dict[str, object]]) -> None:
                for existing in product_rows:
                    product = dict(existing)
                    product_id = int(product["id"])
                    if product_id in loaded_product_ids:
                        continue
                    loaded_product_ids.add(product_id)
                    supplier_id = int(product["fornecedor_id"])
                    existing_by_id[product_id] = product
                    existing_ean = repository._only_digits(product.get("ean", ""))
                    existing_code = str(product.get("codigo_fornecedor", "") or "").strip()
                    if existing_ean:
                        existing_by_ean.setdefault((supplier_id, existing_ean), product_id)
                    if existing_code:
                        existing_by_code.setdefault((supplier_id, existing_code), product_id)

            for supplier_chunk in _chunks(sorted(supplier_ids)):
                for ean_chunk in _chunks(sorted(eans)):
                    cursor.execute(
                        f"""
                        SELECT *
                        FROM cad_produtos_fornecedor
                        WHERE fornecedor_id IN ({_placeholders(supplier_chunk)})
                          AND ean IN ({_placeholders(ean_chunk)})
                        """,
                        [*supplier_chunk, *ean_chunk],
                    )
                    add_existing_products([dict(row) for row in cursor.fetchall()])
                for code_chunk in _chunks(sorted(codes)):
                    cursor.execute(
                        f"""
                        SELECT *
                        FROM cad_produtos_fornecedor
                        WHERE fornecedor_id IN ({_placeholders(supplier_chunk)})
                          AND codigo_fornecedor IN ({_placeholders(code_chunk)})
                        """,
                        [*supplier_chunk, *code_chunk],
                    )
                    add_existing_products([dict(row) for row in cursor.fetchall()])
        finally:
            connection.close()

    insert_values: list[tuple[object, ...]] = []
    update_values: list[tuple[int, dict[str, object]]] = []
    staged_by_ean = dict(existing_by_ean)
    staged_by_code = dict(existing_by_code)
    for index, (row, supplier_id, payload) in enumerate(resolved_rows, start=1):
        if progress_callback is not None and (index == 1 or index == len(resolved_rows) or index % 50 == 0):
            progress_callback(index, len(resolved_rows), f"Preparando produtos ({index}/{len(resolved_rows)})")
        normalized_ean = repository._only_digits(payload.get("ean", ""))[:30]
        normalized_code = repository._trim_text(payload.get("codigo_fornecedor", ""), 80)
        existing_id = staged_by_ean.get((supplier_id, normalized_ean)) if normalized_ean else None
        if existing_id is None and normalized_code:
            existing_id = staged_by_code.get((supplier_id, normalized_code))
        existing_product = existing_by_id.get(existing_id or 0)
        if existing_id is not None and existing_product is None:
            add_skipped_row([row.xml_file, row.codigo_fornecedor, row.ean, row.descricao, "Sem alteracao", "Produto repetido na previa; item ja preparado para cadastro."])
            continue
        if existing_id and not allow_update_existing:
            stats["products_skipped_existing"] += 1
            add_skipped_row([row.xml_file, row.codigo_fornecedor, row.ean, row.descricao, "Ignorado", "Produto ja existe e atualizacao de existentes nao foi autorizada."])
            continue
        if existing_product and allow_update_existing:
            changed_values: dict[str, object] = {}
            for field_name in UPDATABLE_PRODUCT_FIELDS:
                if field_name not in selected_fields:
                    continue
                current_value = existing_product.get(field_name, "")
                new_value = payload.get(field_name, "")
                if _should_ignore_blank_text_update(field_name, current_value, new_value):
                    continue
                if _values_equal(field_name, current_value, new_value):
                    continue
                changed_values[field_name] = _normalize_product_update_value(repository, field_name, new_value)
            if not changed_values:
                add_skipped_row([row.xml_file, row.codigo_fornecedor, row.ean, row.descricao, "Sem alteracao", "Item existente sem diferenca real nos campos selecionados."])
                continue
            update_values.append((int(existing_id), changed_values))
            stats["products_updated"] += 1
        else:
            values = _supplier_product_values(repository, supplier_id, payload)
            insert_values.append(values)
            stats["products_created"] += 1
            fake_id = -stats["products_created"]
            if normalized_ean:
                staged_by_ean[(supplier_id, normalized_ean)] = fake_id
            if normalized_code:
                staged_by_code[(supplier_id, normalized_code)] = fake_id

    if insert_values or update_values:
        if progress_callback is not None:
            progress_callback(0, len(insert_values) + len(update_values), "Gravando produtos no banco...")
        connection = repository.get_connection()
        try:
            cursor = connection.cursor()
            if update_values:
                update_groups: dict[tuple[str, ...], list[tuple[object, ...]]] = {}
                for product_id, changed_values in update_values:
                    field_names = tuple(changed_values)
                    update_groups.setdefault(field_names, []).append((*changed_values.values(), product_id))
                for field_names, grouped_values in update_groups.items():
                    assignments = ", ".join(f"{field_name} = %s" for field_name in field_names)
                    cursor.executemany(
                        f"UPDATE cad_produtos_fornecedor SET {assignments} WHERE id = %s",
                        grouped_values,
                    )
            if insert_values:
                cursor.executemany(
                    """
                    INSERT INTO cad_produtos_fornecedor (
                        fornecedor_id, tipo_produto_id, codigo_fornecedor, codigo_empresa, chave_nfe_origem, descricao, ean, ean_unico, ncm, cest, origem_entrada, cfop_saida_fornecedor, cfop_entrada, cfop_saida, origem_saida,
                        c_classtrib, c_benef, cst_icms, reducao_bc_icms, aliquota_icms, cst_ipi, aliquota_ipi,
                        cst_pis_cofins, aliquota_pis_cofins, cst_pis, cst_cofins, aliquota_pis, aliquota_cofins,
                        bc_st, mva, valor_icms_st, aliquota_icms_st
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s
                    )
                    """,
                    insert_values,
                )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()
    if hidden_skipped_rows:
        add_skipped_row(["...", "", "", "", "Resumo", f"{hidden_skipped_rows} linha(s) omitida(s) para preservar desempenho da tela."])
    stats["skipped_rows"] = skipped_rows  # type: ignore[assignment]
    return stats


def build_catalog_import_change_rows(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
    update_fields: set[str] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[list[str]]:
    changes: list[list[str]] = []
    selected_fields = set(update_fields or UPDATABLE_PRODUCT_FIELDS)
    resolved_rows, products_by_ean, products_by_code = _load_existing_products_for_preview(
        repository,
        environment,
        preview_rows,
        progress_callback,
    )
    for index, (row, supplier_id) in enumerate(resolved_rows, start=1):
        if progress_callback is not None and (index == 1 or index == len(preview_rows) or index % 50 == 0):
            progress_callback(index, len(preview_rows), f"Conferindo mudancas ({index}/{len(preview_rows)})")
        products = products_by_ean.get((supplier_id, row.ean), []) if row.ean else []
        if not products and row.codigo_fornecedor:
            products = products_by_code.get((supplier_id, row.codigo_fornecedor), [])
        if not products:
            continue
        existing = products[0]
        new_map = _payload_from_preview_row(row)
        for field_name in UPDATABLE_PRODUCT_FIELDS:
            if field_name not in selected_fields:
                continue
            old_value = existing.get(field_name, "")
            new_value = new_map.get(field_name, "")
            if _should_ignore_blank_text_update(field_name, old_value, new_value):
                continue
            if not _values_equal(field_name, old_value, new_value):
                changes.append([row.xml_file, row.codigo_fornecedor, field_name, str(old_value), str(new_value), _change_reason(field_name, old_value, new_value)])
    return changes


def build_catalog_import_duplicate_rows(
    repository: MysqlCadastroRepository,
    environment: str,
    preview_rows: list[CatalogImportPreviewRow],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[list[str]]:
    rows: list[list[str]] = []
    resolved_rows, products_by_ean, products_by_code = _load_existing_products_for_preview(
        repository,
        environment,
        preview_rows,
        progress_callback,
    )
    for index, (row, supplier_id) in enumerate(resolved_rows, start=1):
        if progress_callback is not None and (index == 1 or index == len(preview_rows) or index % 50 == 0):
            progress_callback(index, len(preview_rows), f"Conferindo duplicidades ({index}/{len(preview_rows)})")
        duplicates = products_by_ean.get((supplier_id, row.ean), []) if row.ean else []
        if len(duplicates) <= 1 and row.codigo_fornecedor:
            duplicates = products_by_code.get((supplier_id, row.codigo_fornecedor), [])
        if len(duplicates) <= 1:
            continue
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
