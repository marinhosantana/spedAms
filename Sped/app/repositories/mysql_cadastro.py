from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

from app.config import MYSQL_CONNECTION_TIMEOUT_SECONDS, MYSQL_DEFAULT_CONFIG
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

        if column_exists("sped_perfis", "id"):
            if not column_exists("sped_perfis", "empresa_nome_sped"):
                cursor.execute("ALTER TABLE sped_perfis ADD COLUMN empresa_nome_sped VARCHAR(255) NOT NULL DEFAULT '' AFTER nome")
            if not column_exists("sped_perfis", "empresa_cnpj_sped"):
                cursor.execute("ALTER TABLE sped_perfis ADD COLUMN empresa_cnpj_sped VARCHAR(20) NOT NULL DEFAULT '' AFTER empresa_nome_sped")

        if column_exists("sped_arquivos", "id") and not column_exists("sped_arquivos", "empresa_nome_sped"):
            cursor.execute("ALTER TABLE sped_arquivos ADD COLUMN empresa_nome_sped VARCHAR(255) NOT NULL DEFAULT '' AFTER periodo_fim")

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

    def ensure_sped_profile(
        self,
        environment: str,
        profile_name: str,
        description: str = "",
        company_name: str = "",
        company_tax_id: str = "",
    ) -> int:
        normalized_name = str(profile_name or "").strip() or "SPED sem perfil"
        normalized_company_name = str(company_name or "").strip()
        normalized_company_tax_id = "".join(char for char in str(company_tax_id or "") if char.isdigit())
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_company_tax_id:
                cursor.execute(
                    """
                    SELECT id
                    FROM sped_perfis
                    WHERE ambiente = %s
                      AND empresa_cnpj_sped = %s
                    LIMIT 1
                    """,
                    (environment, normalized_company_tax_id),
                )
            else:
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
                cursor.execute(
                    """
                    UPDATE sped_perfis
                    SET nome = CASE WHEN nome = '' THEN %s ELSE nome END,
                        empresa_nome_sped = CASE WHEN empresa_nome_sped = '' THEN %s ELSE empresa_nome_sped END,
                        empresa_cnpj_sped = CASE WHEN empresa_cnpj_sped = '' THEN %s ELSE empresa_cnpj_sped END,
                        ativo = 1
                    WHERE id = %s
                    """,
                    (normalized_name, normalized_company_name, normalized_company_tax_id, profile_id),
                )
                connection.commit()
                return profile_id

            cursor.execute(
                """
                INSERT INTO sped_perfis (
                    ambiente,
                    nome,
                    empresa_nome_sped,
                    empresa_cnpj_sped,
                    descricao,
                    ativo
                ) VALUES (%s, %s, %s, %s, %s, 1)
                """,
                (environment, normalized_name, normalized_company_name, normalized_company_tax_id, description),
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
                    ambiente,
                    tipo_sped,
                    periodo_inicio,
                    periodo_fim,
                    empresa_nome_sped,
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
                    data["ambiente"],
                    data.get("tipo_sped", ""),
                    data.get("periodo_inicio") or None,
                    data.get("periodo_fim") or None,
                    data.get("empresa_nome_sped", ""),
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
                    p.empresa_nome_sped,
                    p.empresa_cnpj_sped,
                    p.ambiente,
                    p.empresa_nome_sped AS empresa_nome,
                    p.empresa_cnpj_sped AS empresa_cnpj,
                    COUNT(a.id) AS total_arquivos,
                    COALESCE(SUM(a.arquivo_tamanho), 0) AS total_bytes,
                    MIN(a.periodo_inicio) AS periodo_inicio,
                    MAX(a.periodo_fim) AS periodo_fim,
                    MAX(a.created_at) AS ultimo_arquivo_em
                FROM sped_perfis p
                LEFT JOIN sped_arquivos a ON a.perfil_id = p.id
                WHERE p.ambiente = %s
                  AND p.ativo = 1
                GROUP BY p.id, p.nome, p.empresa_nome_sped, p.empresa_cnpj_sped, p.ambiente
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
                    a.empresa_nome_sped AS empresa_nome,
                    a.ambiente,
                    a.tipo_sped,
                    a.periodo_inicio,
                    a.periodo_fim,
                    a.empresa_nome_sped,
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

