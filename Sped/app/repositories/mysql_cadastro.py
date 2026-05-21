from __future__ import annotations

import datetime as dt
import json
import sys
import traceback
from pathlib import Path

from app.config import MYSQL_CONNECTION_TIMEOUT_SECONDS, MYSQL_DEFAULT_CONFIG
from app.services.tax_rules import normalize_tax_code

try:
    import mysql.connector
    from mysql.connector import Error as MySQLError
    MYSQL_IMPORT_ERROR = ""
except Exception:
    mysql = None
    MYSQL_IMPORT_ERROR = traceback.format_exc()

    class MySQLError(Exception):
        pass


class MysqlCadastroRepository:
    def __init__(self, config_path: Path, schema_path: Path, default_config: dict[str, str] | None = None) -> None:
        self.config_path = config_path
        self.schema_path = schema_path
        self.default_config = dict(default_config or MYSQL_DEFAULT_CONFIG)

    def mysql_available(self) -> bool:
        return "mysql" in globals() and getattr(mysql, "connector", None) is not None

    def ensure_driver(self) -> None:
        if not self.mysql_available():
            runtime_python = sys.executable
            import_error = MYSQL_IMPORT_ERROR.strip()
            details = f"Python em uso: {runtime_python}"
            if import_error:
                details = f"{details}\n\nDetalhe da importacao:\n{import_error}"
            raise RuntimeError(
                "Driver MySQL nao encontrado. Instale o pacote mysql-connector-python no interpretador que esta executando o sistema.\n\n"
                + details
            )

    def load_config(self) -> dict[str, str]:
        if not self.config_path.exists():
            self.save_config(self.default_config)
            return dict(self.default_config)
        try:
            loaded = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            loaded = {}
        config = dict(self.default_config)
        if isinstance(loaded, dict):
            for key in self.default_config:
                value = loaded.get(key, self.default_config[key])
                config[key] = str(value or self.default_config[key])
        return config

    def save_config(self, config: dict[str, str]) -> None:
        payload = {key: str(config.get(key, self.default_config[key])) for key in self.default_config}
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self.config_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")

    def _build_connection_config(self, include_database: bool = True) -> dict[str, object]:
        config = self.load_config()
        connection_config: dict[str, object] = {
            "host": config["host"],
            "port": int(str(config["port"] or "3306")),
            "user": config["user"],
            "password": config["password"],
            "connection_timeout": MYSQL_CONNECTION_TIMEOUT_SECONDS,
        }
        if include_database:
            connection_config["database"] = config["database"]
        return connection_config

    def get_connection(self, include_database: bool = True):
        self.ensure_driver()
        return mysql.connector.connect(**self._build_connection_config(include_database=include_database))

    def get_table_columns(self, table_name: str) -> set[str]:
        database_name = self.load_config()["database"]
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT COLUMN_NAME
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = %s
                """,
                (database_name, table_name),
            )
            return {str(row[0]).strip() for row in cursor.fetchall()}
        finally:
            connection.close()

    def test_connection(self) -> None:
        connection = self.get_connection(include_database=False)
        connection.close()

    def ensure_schema(self) -> None:
        schema_sql = self.schema_path.read_text(encoding="utf-8")
        database_name = self.load_config()["database"]
        connection = self.get_connection(include_database=False)
        try:
            cursor = connection.cursor()
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{database_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            cursor.execute(f"USE `{database_name}`")
            for statement in [chunk.strip() for chunk in schema_sql.split(";") if chunk.strip()]:
                cursor.execute(statement)
            self.ensure_schema_migrations(cursor, database_name)
            connection.commit()
        finally:
            connection.close()

    def ensure_schema_migrations(self, cursor, database_name: str) -> None:
        def column_exists(table_name: str, column_name: str) -> bool:
            cursor.execute(
                """
                SELECT 1
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = %s
                  AND COLUMN_NAME = %s
                LIMIT 1
                """,
                (database_name, table_name, column_name),
            )
            return cursor.fetchone() is not None

        def normalize_existing_icms_cst_values() -> None:
            if not column_exists("produtos_empresa", "cst_icms_entrada") or not column_exists("produtos_empresa", "cst_icms_saida"):
                return
            cursor.execute("SELECT id, cst_icms_entrada, cst_icms_saida FROM produtos_empresa")
            updates: list[tuple[str, str, int]] = []
            for product_id, cst_entrada, cst_saida in cursor.fetchall():
                normalized_entrada = normalize_tax_code(cst_entrada, 3)
                normalized_saida = normalize_tax_code(cst_saida, 3)
                current_entrada = str(cst_entrada or "").strip()
                current_saida = str(cst_saida or "").strip()
                if normalized_entrada != current_entrada or normalized_saida != current_saida:
                    updates.append((normalized_entrada, normalized_saida, int(product_id)))
            if updates:
                cursor.executemany(
                    """
                    UPDATE produtos_empresa
                    SET cst_icms_entrada = %s,
                        cst_icms_saida = %s
                    WHERE id = %s
                    """,
                    updates,
                )

        if not column_exists("empresas", "ativo"):
            cursor.execute("ALTER TABLE empresas ADD COLUMN ativo TINYINT(1) NOT NULL DEFAULT 1")

        if not column_exists("produtos_empresa", "codigo_origem"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN codigo_origem VARCHAR(80) NOT NULL DEFAULT ''")

        if not column_exists("produtos_empresa", "ativo"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN ativo TINYINT(1) NOT NULL DEFAULT 1")

        if not column_exists("produtos_empresa", "cst_icms_entrada"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_icms_entrada VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "cst_icms_saida"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_icms_saida VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "cst_pis_entrada"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_pis_entrada VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "cst_pis_saida"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_pis_saida VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "cst_cofins_entrada"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_cofins_entrada VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "cst_cofins_saida"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN cst_cofins_saida VARCHAR(4) NOT NULL DEFAULT ''")
        if not column_exists("produtos_empresa", "tipo_produto"):
            cursor.execute("ALTER TABLE produtos_empresa ADD COLUMN tipo_produto VARCHAR(30) NOT NULL DEFAULT 'Revenda'")

        normalize_existing_icms_cst_values()

    def find_company_id_by_tax_id(self, tax_id: str) -> int | None:
        normalized_tax_id = "".join(char for char in str(tax_id or "") if char.isdigit())
        if not normalized_tax_id:
            return None
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT id
                FROM empresas
                WHERE REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') = %s
                LIMIT 1
                """,
                (normalized_tax_id,),
            )
            row = cursor.fetchone()
            return int(row[0]) if row else None
        finally:
            connection.close()

    def get_sped_archive_by_hash(self, environment: str, file_hash_sha256: str) -> dict[str, object] | None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT *
                FROM sped_arquivos
                WHERE ambiente = %s
                  AND arquivo_hash_sha256 = %s
                LIMIT 1
                """,
                (environment, file_hash_sha256),
            )
            row = cursor.fetchone()
            return dict(row) if row else None
        finally:
            connection.close()

    def ensure_sped_profile(self, environment: str, profile_name: str, company_id: int | None = None, description: str = "") -> int:
        normalized_name = str(profile_name or "").strip() or "SPED sem perfil"
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT id
                FROM sped_perfis
                WHERE ambiente = %s
                  AND nome = %s
                LIMIT 1
                """,
                (environment, normalized_name),
            )
            row = cursor.fetchone()
            if row:
                profile_id = int(row[0])
                if company_id:
                    cursor.execute(
                        """
                        UPDATE sped_perfis
                        SET empresa_id = COALESCE(empresa_id, %s),
                            ativo = 1
                        WHERE id = %s
                        """,
                        (company_id, profile_id),
                    )
                    connection.commit()
                return profile_id

            cursor.execute(
                """
                INSERT INTO sped_perfis (
                    empresa_id,
                    ambiente,
                    nome,
                    descricao,
                    ativo
                ) VALUES (%s, %s, %s, %s, 1)
                """,
                (company_id, environment, normalized_name, description),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def save_sped_archive(self, data: dict[str, object]) -> int:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT id
                FROM sped_arquivos
                WHERE ambiente = %s
                  AND arquivo_hash_sha256 = %s
                LIMIT 1
                """,
                (data["ambiente"], data["arquivo_hash_sha256"]),
            )
            row = cursor.fetchone()
            if row:
                return int(row[0])

            cursor.execute(
                """
                INSERT INTO sped_arquivos (
                    perfil_id,
                    empresa_id,
                    ambiente,
                    tipo_sped,
                    periodo_inicio,
                    periodo_fim,
                    empresa_cnpj,
                    arquivo_nome_original,
                    arquivo_hash_sha256,
                    arquivo_tamanho,
                    caminho_arquivo_original,
                    caminho_arquivo_arquivado,
                    observacao
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    data.get("perfil_id"),
                    data.get("empresa_id"),
                    data["ambiente"],
                    data.get("tipo_sped", ""),
                    data.get("periodo_inicio") or None,
                    data.get("periodo_fim") or None,
                    data.get("empresa_cnpj", ""),
                    data["arquivo_nome_original"],
                    data["arquivo_hash_sha256"],
                    data["arquivo_tamanho"],
                    data["caminho_arquivo_original"],
                    data["caminho_arquivo_arquivado"],
                    data.get("observacao", ""),
                ),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def replace_sped_extracted_data(
        self,
        sped_archive_id: int,
        products: list[object],
        detailed_items: list[dict[str, object]],
        c190_rows: list[dict[str, object]],
    ) -> None:
        def decimal_value(value: object, fallback: str = "0") -> str:
            return str(value if value not in (None, "") else fallback)

        document_keys: set[tuple[str, str, str, str, str, str, str, str, str, str]] = set()
        for item in detailed_items:
            document_keys.add(
                (
                    str(item.get("operation_type", "")),
                    str(item.get("document_number", "")),
                    str(item.get("document_key", "")),
                    str(item.get("document_date", "")),
                    str(item.get("document_series", "")),
                    str(item.get("document_model", "")),
                    str(item.get("participant_code", "")),
                    str(item.get("participant_name", "")),
                    str(item.get("participant_tax_id", "")),
                    str(item.get("document_tax_id", "")),
                )
            )
        for row in c190_rows:
            document_keys.add(
                (
                    str(row.get("operation_type", "")),
                    str(row.get("document_number", "")),
                    str(row.get("document_key", "")),
                    str(row.get("document_date", "")),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                )
            )

        product_values = [
            (
                sped_archive_id,
                str(getattr(product, "code", "")),
                str(getattr(product, "description", "")),
                str(getattr(product, "ncm", "")),
                str(getattr(product, "cest", "")),
                str(getattr(product, "cst_icms", "")),
                decimal_value(getattr(product, "icms_rate", "0")),
            )
            for product in products
        ]
        document_values = [
            (
                sped_archive_id,
                operation_type,
                document_number,
                document_key,
                document_date,
                document_series,
                document_model,
                participant_code,
                participant_name,
                participant_tax_id or document_tax_id,
            )
            for (
                operation_type,
                document_number,
                document_key,
                document_date,
                document_series,
                document_model,
                participant_code,
                participant_name,
                participant_tax_id,
                document_tax_id,
            ) in sorted(document_keys)
        ]
        item_values = [
            (
                sped_archive_id,
                str(item.get("operation_type", "")),
                str(item.get("document_number", "")),
                str(item.get("document_key", "")),
                str(item.get("document_date", "")),
                str(item.get("document_series", "")),
                str(item.get("document_model", "")),
                str(item.get("participant_code", "")),
                str(item.get("participant_name", "")),
                str(item.get("participant_tax_id", "") or item.get("document_tax_id", "")),
                str(item.get("item_number", "")),
                str(item.get("code", "")),
                str(item.get("description", "")),
                str(item.get("ncm", "")),
                str(item.get("cest", "")),
                str(item.get("cst_icms", "")),
                str(item.get("cfop", "")),
                decimal_value(item.get("quantity")),
                decimal_value(item.get("sale_value")),
                decimal_value(item.get("discount_value")),
                decimal_value(item.get("base_icms")),
                decimal_value(item.get("icms_rate")),
                decimal_value(item.get("icms_value")),
                decimal_value(item.get("base_icms_st")),
                decimal_value(item.get("icms_st_rate")),
                decimal_value(item.get("icms_st_value")),
                decimal_value(item.get("base_ipi")),
                decimal_value(item.get("ipi_rate")),
                decimal_value(item.get("ipi_value")),
            )
            for item in detailed_items
        ]
        c190_values = [
            (
                sped_archive_id,
                str(row.get("operation_type", "")),
                str(row.get("document_number", "")),
                str(row.get("document_key", "")),
                str(row.get("document_date", "")),
                str(row.get("cst_icms", "")),
                str(row.get("cfop", "")),
                decimal_value(row.get("icms_rate")),
                decimal_value(row.get("total_operation_value")),
                decimal_value(row.get("base_icms")),
                decimal_value(row.get("icms_value")),
                decimal_value(row.get("base_icms_st")),
                decimal_value(row.get("icms_st_value")),
                decimal_value(row.get("reduction_value")),
                decimal_value(row.get("ipi_value")),
            )
            for row in c190_rows
        ]

        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            for table_name in ("sped_resumos_c190", "sped_itens_c170", "sped_documentos", "sped_produtos_0200"):
                cursor.execute(f"DELETE FROM {table_name} WHERE sped_arquivo_id = %s", (sped_archive_id,))

            if product_values:
                cursor.executemany(
                    """
                    INSERT INTO sped_produtos_0200 (
                        sped_arquivo_id, codigo, descricao, ncm, cest, cst_icms, aliquota_icms
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """,
                    product_values,
                )
            if document_values:
                cursor.executemany(
                    """
                    INSERT INTO sped_documentos (
                        sped_arquivo_id, tipo_operacao, numero_documento, chave_documento, data_documento,
                        serie_documento, modelo_documento, participante_codigo, participante_nome, participante_cnpj
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    document_values,
                )
            if item_values:
                cursor.executemany(
                    """
                    INSERT INTO sped_itens_c170 (
                        sped_arquivo_id, tipo_operacao, numero_documento, chave_documento, data_documento,
                        serie_documento, modelo_documento, participante_codigo, participante_nome, participante_cnpj,
                        numero_item, codigo_produto, descricao_produto, ncm, cest, cst_icms, cfop,
                        quantidade, valor_operacao, valor_desconto, base_icms, aliquota_icms, valor_icms,
                        base_icms_st, aliquota_icms_st, valor_icms_st, base_ipi, aliquota_ipi, valor_ipi
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s
                    )
                    """,
                    item_values,
                )
            if c190_values:
                cursor.executemany(
                    """
                    INSERT INTO sped_resumos_c190 (
                        sped_arquivo_id, tipo_operacao, numero_documento, chave_documento, data_documento,
                        cst_icms, cfop, aliquota_icms, valor_operacao, base_icms, valor_icms,
                        base_icms_st, valor_icms_st, valor_reducao, valor_ipi
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    c190_values,
                )
            connection.commit()
        finally:
            connection.close()

    def list_sped_profiles(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    p.id,
                    p.nome,
                    p.ambiente,
                    p.empresa_id,
                    COALESCE(e.razao_social, '') AS empresa_nome,
                    COALESCE(e.cnpj, '') AS empresa_cnpj,
                    COUNT(a.id) AS total_arquivos,
                    COALESCE(SUM(a.arquivo_tamanho), 0) AS total_bytes,
                    MIN(a.periodo_inicio) AS periodo_inicio,
                    MAX(a.periodo_fim) AS periodo_fim,
                    MAX(a.created_at) AS ultimo_arquivo_em
                FROM sped_perfis p
                LEFT JOIN empresas e ON e.id = p.empresa_id
                LEFT JOIN sped_arquivos a ON a.perfil_id = p.id
                WHERE p.ambiente = %s
                  AND p.ativo = 1
                GROUP BY p.id, p.nome, p.ambiente, p.empresa_id, e.razao_social, e.cnpj
                ORDER BY ultimo_arquivo_em DESC, p.nome
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_sped_archives(self, environment: str, profile_id: int | None = None) -> list[dict[str, object]]:
        filters = ["a.ambiente = %s"]
        params: list[object] = [environment]
        if profile_id:
            filters.append("a.perfil_id = %s")
            params.append(profile_id)
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                f"""
                SELECT
                    a.id,
                    a.perfil_id,
                    COALESCE(p.nome, '') AS perfil_nome,
                    a.empresa_id,
                    COALESCE(e.razao_social, '') AS empresa_nome,
                    a.ambiente,
                    a.tipo_sped,
                    a.periodo_inicio,
                    a.periodo_fim,
                    a.empresa_cnpj,
                    a.arquivo_nome_original,
                    a.arquivo_hash_sha256,
                    a.arquivo_tamanho,
                    a.caminho_arquivo_original,
                    a.caminho_arquivo_arquivado,
                    a.created_at,
                    COALESCE(produtos.total, 0) AS total_produtos,
                    COALESCE(documentos.total, 0) AS total_documentos,
                    COALESCE(itens.total, 0) AS total_itens,
                    COALESCE(c190.total, 0) AS total_c190
                FROM sped_arquivos a
                LEFT JOIN sped_perfis p ON p.id = a.perfil_id
                LEFT JOIN empresas e ON e.id = a.empresa_id
                LEFT JOIN (
                    SELECT sped_arquivo_id, COUNT(*) AS total
                    FROM sped_produtos_0200
                    GROUP BY sped_arquivo_id
                ) produtos ON produtos.sped_arquivo_id = a.id
                LEFT JOIN (
                    SELECT sped_arquivo_id, COUNT(*) AS total
                    FROM sped_documentos
                    GROUP BY sped_arquivo_id
                ) documentos ON documentos.sped_arquivo_id = a.id
                LEFT JOIN (
                    SELECT sped_arquivo_id, COUNT(*) AS total
                    FROM sped_itens_c170
                    GROUP BY sped_arquivo_id
                ) itens ON itens.sped_arquivo_id = a.id
                LEFT JOIN (
                    SELECT sped_arquivo_id, COUNT(*) AS total
                    FROM sped_resumos_c190
                    GROUP BY sped_arquivo_id
                ) c190 ON c190.sped_arquivo_id = a.id
                WHERE {" AND ".join(filters)}
                ORDER BY a.created_at DESC, a.id DESC
                """,
                tuple(params),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_companies(self) -> list[dict[str, object]]:
        columns = self.get_table_columns("empresas")
        has_ativo = "ativo" in columns
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                f"""
                SELECT
                    id,
                    razao_social,
                    nome_fantasia,
                    cnpj,
                    inscricao_estadual
                    {", ativo" if has_ativo else ""}
                FROM empresas
                ORDER BY razao_social, id
                """
            )
            rows = list(cursor.fetchall())
            if not has_ativo:
                for row in rows:
                    row["ativo"] = 1
            return rows
        finally:
            connection.close()

    def save_company(self, company_id: int | None, data: dict[str, str]) -> int:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            values = (
                data["razao_social"],
                data["nome_fantasia"],
                data["cnpj"],
                data["inscricao_estadual"],
                int(data.get("ativo", 1)),
            )
            if company_id:
                cursor.execute(
                    """
                    UPDATE empresas
                    SET razao_social = %s,
                        nome_fantasia = %s,
                        cnpj = %s,
                        inscricao_estadual = %s,
                        ativo = %s
                    WHERE id = %s
                    """,
                    (*values, company_id),
                )
                saved_id = company_id
            else:
                cursor.execute(
                    """
                    INSERT INTO empresas (
                        razao_social,
                        nome_fantasia,
                        cnpj,
                        inscricao_estadual,
                        ativo
                    ) VALUES (%s, %s, %s, %s, %s)
                    """,
                    values,
                )
                saved_id = int(cursor.lastrowid)
            connection.commit()
            return saved_id
        finally:
            connection.close()

    def deactivate_company(self, company_id: int) -> None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute("UPDATE empresas SET ativo = 0 WHERE id = %s", (company_id,))
            connection.commit()
        finally:
            connection.close()

    def reactivate_company(self, company_id: int) -> None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute("UPDATE empresas SET ativo = 1 WHERE id = %s", (company_id,))
            connection.commit()
        finally:
            connection.close()

    def list_products(self, company_id: int) -> list[dict[str, object]]:
        columns = self.get_table_columns("produtos_empresa")
        has_codigo_origem = "codigo_origem" in columns
        has_ativo = "ativo" in columns
        has_cst_icms_entrada = "cst_icms_entrada" in columns
        has_cst_icms_saida = "cst_icms_saida" in columns
        has_cst_pis_entrada = "cst_pis_entrada" in columns
        has_cst_pis_saida = "cst_pis_saida" in columns
        has_cst_cofins_entrada = "cst_cofins_entrada" in columns
        has_cst_cofins_saida = "cst_cofins_saida" in columns
        has_tipo_produto = "tipo_produto" in columns
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                f"""
                SELECT
                    id,
                    codigo,
                    {"codigo_origem," if has_codigo_origem else "'' AS codigo_origem,"}
                    descricao,
                    ncm,
                    unidade,
                    {"cst_icms_entrada," if has_cst_icms_entrada else "'' AS cst_icms_entrada,"}
                    {"cst_icms_saida," if has_cst_icms_saida else "'' AS cst_icms_saida,"}
                    {"cst_pis_entrada," if has_cst_pis_entrada else "'' AS cst_pis_entrada,"}
                    {"cst_pis_saida," if has_cst_pis_saida else "'' AS cst_pis_saida,"}
                    {"cst_cofins_entrada," if has_cst_cofins_entrada else "'' AS cst_cofins_entrada,"}
                    {"cst_cofins_saida," if has_cst_cofins_saida else "'' AS cst_cofins_saida,"}
                    {"tipo_produto," if has_tipo_produto else "'Revenda' AS tipo_produto,"}
                    icms_entrada,
                    icms_saida,
                    pis_entrada,
                    pis_saida,
                    cofins_entrada,
                    cofins_saida
                    {", ativo" if has_ativo else ""}
                FROM produtos_empresa
                WHERE empresa_id = %s
                ORDER BY descricao, codigo, id
                """,
                (company_id,),
            )
            rows = list(cursor.fetchall())
            for row in rows:
                if not has_codigo_origem:
                    row["codigo_origem"] = ""
                row["cst_icms_entrada"] = normalize_tax_code(row.get("cst_icms_entrada", "") if has_cst_icms_entrada else "", 3)
                row["cst_icms_saida"] = normalize_tax_code(row.get("cst_icms_saida", "") if has_cst_icms_saida else "", 3)
                row["cst_pis_entrada"] = normalize_tax_code(row.get("cst_pis_entrada", "") if has_cst_pis_entrada else "", 2)
                row["cst_pis_saida"] = normalize_tax_code(row.get("cst_pis_saida", "") if has_cst_pis_saida else "", 2)
                row["cst_cofins_entrada"] = normalize_tax_code(row.get("cst_cofins_entrada", "") if has_cst_cofins_entrada else "", 2)
                row["cst_cofins_saida"] = normalize_tax_code(row.get("cst_cofins_saida", "") if has_cst_cofins_saida else "", 2)
                if not has_tipo_produto:
                    row["tipo_produto"] = "Revenda"
                if not has_ativo:
                    row["ativo"] = 1
            return rows
        finally:
            connection.close()

    def save_product(self, product_id: int | None, company_id: int, data: dict[str, object]) -> int:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            normalized_cst_icms_entrada = normalize_tax_code(data.get("cst_icms_entrada", ""), 3)
            normalized_cst_icms_saida = normalize_tax_code(data.get("cst_icms_saida", ""), 3)
            values = (
                company_id,
                data["codigo_origem"],
                data["descricao"],
                data["ncm"],
                data["unidade"],
                normalized_cst_icms_entrada,
                normalized_cst_icms_saida,
                data["cst_pis_entrada"],
                data["cst_pis_saida"],
                data["cst_cofins_entrada"],
                data["cst_cofins_saida"],
                data["tipo_produto"],
                data["icms_entrada"],
                data["icms_saida"],
                data["pis_entrada"],
                data["pis_saida"],
                data["cofins_entrada"],
                data["cofins_saida"],
                int(data.get("ativo", 1)),
            )
            if product_id:
                cursor.execute(
                    """
                    UPDATE produtos_empresa
                    SET empresa_id = %s,
                        codigo_origem = %s,
                        descricao = %s,
                        ncm = %s,
                        unidade = %s,
                        cst_icms_entrada = %s,
                        cst_icms_saida = %s,
                        cst_pis_entrada = %s,
                        cst_pis_saida = %s,
                        cst_cofins_entrada = %s,
                        cst_cofins_saida = %s,
                        tipo_produto = %s,
                        icms_entrada = %s,
                        icms_saida = %s,
                        pis_entrada = %s,
                        pis_saida = %s,
                        cofins_entrada = %s,
                        cofins_saida = %s,
                        ativo = %s
                    WHERE id = %s
                    """,
                    (*values, product_id),
                )
                saved_id = product_id
            else:
                temporary_code = f"TMP-{company_id}-{dt.datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                cursor.execute(
                    """
                    INSERT INTO produtos_empresa (
                        empresa_id,
                        codigo,
                        codigo_origem,
                        descricao,
                        ncm,
                        unidade,
                        cst_icms_entrada,
                        cst_icms_saida,
                        cst_pis_entrada,
                        cst_pis_saida,
                        cst_cofins_entrada,
                        cst_cofins_saida,
                        tipo_produto,
                        icms_entrada,
                        icms_saida,
                        pis_entrada,
                        pis_saida,
                        cofins_entrada,
                        cofins_saida,
                        ativo
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        company_id,
                        temporary_code,
                        data["codigo_origem"],
                        data["descricao"],
                        data["ncm"],
                        data["unidade"],
                        normalized_cst_icms_entrada,
                        normalized_cst_icms_saida,
                        data["cst_pis_entrada"],
                        data["cst_pis_saida"],
                        data["cst_cofins_entrada"],
                        data["cst_cofins_saida"],
                        data["tipo_produto"],
                        data["icms_entrada"],
                        data["icms_saida"],
                        data["pis_entrada"],
                        data["pis_saida"],
                        data["cofins_entrada"],
                        data["cofins_saida"],
                        int(data.get("ativo", 1)),
                    ),
                )
                saved_id = int(cursor.lastrowid)
                cursor.execute("UPDATE produtos_empresa SET codigo = %s WHERE id = %s", (str(saved_id), saved_id))
            connection.commit()
            return saved_id
        finally:
            connection.close()

    def deactivate_product(self, product_id: int) -> None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute("UPDATE produtos_empresa SET ativo = 0 WHERE id = %s", (product_id,))
            connection.commit()
        finally:
            connection.close()

    def reactivate_product(self, product_id: int) -> None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute("UPDATE produtos_empresa SET ativo = 1 WHERE id = %s", (product_id,))
            connection.commit()
        finally:
            connection.close()

