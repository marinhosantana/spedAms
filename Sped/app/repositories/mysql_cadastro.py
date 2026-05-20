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
    def __init__(self, config_path: Path, schema_path: Path) -> None:
        self.config_path = config_path
        self.schema_path = schema_path

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
            self.save_config(MYSQL_DEFAULT_CONFIG)
            return dict(MYSQL_DEFAULT_CONFIG)
        try:
            loaded = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            loaded = {}
        config = dict(MYSQL_DEFAULT_CONFIG)
        if isinstance(loaded, dict):
            for key in MYSQL_DEFAULT_CONFIG:
                value = loaded.get(key, MYSQL_DEFAULT_CONFIG[key])
                config[key] = str(value or MYSQL_DEFAULT_CONFIG[key])
        return config

    def save_config(self, config: dict[str, str]) -> None:
        payload = {key: str(config.get(key, MYSQL_DEFAULT_CONFIG[key])) for key in MYSQL_DEFAULT_CONFIG}
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

