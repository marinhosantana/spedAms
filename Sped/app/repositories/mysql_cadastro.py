from __future__ import annotations

import hashlib
import json
import secrets
import sys
import traceback
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
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

        def table_constraint_exists(table_name: str, constraint_name: str) -> bool:
            cursor.execute(
                """
                SELECT 1
                FROM information_schema.TABLE_CONSTRAINTS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = %s
                  AND CONSTRAINT_NAME = %s
                LIMIT 1
                """,
                (database_name, table_name, constraint_name),
            )
            return cursor.fetchone() is not None

        def index_exists(table_name: str, index_name: str) -> bool:
            cursor.execute(
                """
                SELECT 1
                FROM information_schema.STATISTICS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = %s
                  AND INDEX_NAME = %s
                LIMIT 1
                """,
                (database_name, table_name, index_name),
            )
            return cursor.fetchone() is not None

        if column_exists("sped_perfis", "id"):
            if not column_exists("sped_perfis", "empresa_nome_sped"):
                cursor.execute("ALTER TABLE sped_perfis ADD COLUMN empresa_nome_sped VARCHAR(255) NOT NULL DEFAULT '' AFTER nome")
            if not column_exists("sped_perfis", "empresa_cnpj_sped"):
                cursor.execute("ALTER TABLE sped_perfis ADD COLUMN empresa_cnpj_sped VARCHAR(20) NOT NULL DEFAULT '' AFTER empresa_nome_sped")

        if column_exists("sped_arquivos", "id") and not column_exists("sped_arquivos", "empresa_nome_sped"):
            cursor.execute("ALTER TABLE sped_arquivos ADD COLUMN empresa_nome_sped VARCHAR(255) NOT NULL DEFAULT '' AFTER periodo_fim")

        if column_exists("cad_tipos_produto", "id"):
            if not column_exists("cad_tipos_produto", "ambiente"):
                cursor.execute("ALTER TABLE cad_tipos_produto ADD COLUMN ambiente VARCHAR(10) NOT NULL DEFAULT 'dev' AFTER id")
            if table_constraint_exists("cad_tipos_produto", "fk_cad_tipos_produto_empresa"):
                cursor.execute("ALTER TABLE cad_tipos_produto DROP FOREIGN KEY fk_cad_tipos_produto_empresa")
            if index_exists("cad_tipos_produto", "uq_cad_tipos_produto_empresa_nome"):
                cursor.execute("ALTER TABLE cad_tipos_produto DROP INDEX uq_cad_tipos_produto_empresa_nome")
            if column_exists("cad_tipos_produto", "empresa_id"):
                cursor.execute("ALTER TABLE cad_tipos_produto DROP COLUMN empresa_id")
            if not index_exists("cad_tipos_produto", "idx_cad_tipos_produto_ambiente_nome"):
                cursor.execute("ALTER TABLE cad_tipos_produto ADD KEY idx_cad_tipos_produto_ambiente_nome (ambiente, nome)")

        if column_exists("cad_produtos_fornecedor", "cst_pis_cofins"):
            cursor.execute(
                """
                SELECT CHARACTER_MAXIMUM_LENGTH
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = 'cad_produtos_fornecedor'
                  AND COLUMN_NAME = 'cst_pis_cofins'
                LIMIT 1
                """,
                (database_name,),
            )
            row = cursor.fetchone()
            current_length = int(row[0]) if row and row[0] is not None else 0
            if current_length < 20:
                cursor.execute(
                    "ALTER TABLE cad_produtos_fornecedor MODIFY COLUMN cst_pis_cofins VARCHAR(20) NOT NULL DEFAULT ''"
                )
        if column_exists("cad_produtos_fornecedor", "id"):
            if not column_exists("cad_produtos_fornecedor", "cst_pis"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cst_pis VARCHAR(4) NOT NULL DEFAULT '' AFTER cst_pis_cofins")
            if not column_exists("cad_produtos_fornecedor", "cst_cofins"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cst_cofins VARCHAR(4) NOT NULL DEFAULT '' AFTER cst_pis")
            if not column_exists("cad_produtos_fornecedor", "aliquota_pis"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN aliquota_pis DECIMAL(10,4) NOT NULL DEFAULT 0.0000 AFTER aliquota_pis_cofins")
            if not column_exists("cad_produtos_fornecedor", "aliquota_cofins"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN aliquota_cofins DECIMAL(10,4) NOT NULL DEFAULT 0.0000 AFTER aliquota_pis")
            if index_exists("cad_produtos_fornecedor", "uq_cad_produtos_fornecedor_ean"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor DROP INDEX uq_cad_produtos_fornecedor_ean")
            if not column_exists("cad_produtos_fornecedor", "ean_unico"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN ean_unico VARCHAR(30) NULL AFTER ean")
            cursor.execute(
                """
                UPDATE cad_produtos_fornecedor
                SET ean_unico = CASE
                    WHEN TRIM(COALESCE(ean, '')) = '' THEN NULL
                    ELSE TRIM(ean)
                END
                """
            )
            if not index_exists("cad_produtos_fornecedor", "uq_cad_produtos_fornecedor_ean_unico"):
                cursor.execute(
                    """
                    SELECT fornecedor_id, ean_unico, COUNT(*) total
                    FROM cad_produtos_fornecedor
                    WHERE ean_unico IS NOT NULL AND TRIM(COALESCE(ean_unico, '')) <> ''
                    GROUP BY fornecedor_id, ean_unico
                    HAVING COUNT(*) > 1
                    LIMIT 1
                    """
                )
                has_duplicates = cursor.fetchone() is not None
                if not has_duplicates:
                    cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD UNIQUE KEY uq_cad_produtos_fornecedor_ean_unico (fornecedor_id, ean_unico)")
        if column_exists("cad_fornecedores", "id"):
            if not column_exists("cad_fornecedores", "inscricao_estadual"):
                cursor.execute("ALTER TABLE cad_fornecedores ADD COLUMN inscricao_estadual VARCHAR(30) NOT NULL DEFAULT '' AFTER cnpj")
            if not column_exists("cad_fornecedores", "uf"):
                cursor.execute("ALTER TABLE cad_fornecedores ADD COLUMN uf VARCHAR(2) NOT NULL DEFAULT '' AFTER inscricao_estadual")
            if not column_exists("cad_fornecedores", "regime_tributario"):
                cursor.execute(
                    "ALTER TABLE cad_fornecedores ADD COLUMN regime_tributario VARCHAR(30) NOT NULL DEFAULT 'LUCRO_REAL_PRESUMIDO' AFTER codigo"
                )
        if column_exists("cad_produtos_fornecedor", "id"):
            if not column_exists("cad_produtos_fornecedor", "cfop_entrada"):
                cursor.execute(
                    "ALTER TABLE cad_produtos_fornecedor ADD COLUMN cfop_entrada VARCHAR(10) NOT NULL DEFAULT '' AFTER cest"
                )
            if not column_exists("cad_produtos_fornecedor", "cfop_saida"):
                cursor.execute(
                    "ALTER TABLE cad_produtos_fornecedor ADD COLUMN cfop_saida VARCHAR(10) NOT NULL DEFAULT '' AFTER cfop_entrada"
                )
            if not column_exists("cad_produtos_fornecedor", "status_produto"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN status_produto VARCHAR(40) NOT NULL DEFAULT '' AFTER codigo_empresa")
            if not column_exists("cad_produtos_fornecedor", "origem_entrada"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN origem_entrada VARCHAR(4) NOT NULL DEFAULT '' AFTER cest")
            if not column_exists("cad_produtos_fornecedor", "cfop_saida_fornecedor"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cfop_saida_fornecedor VARCHAR(10) NOT NULL DEFAULT '' AFTER origem_entrada")
            if not column_exists("cad_produtos_fornecedor", "natureza_receita_entrada"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN natureza_receita_entrada VARCHAR(40) NOT NULL DEFAULT '' AFTER cfop_saida")
            if not column_exists("cad_produtos_fornecedor", "origem_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN origem_saida VARCHAR(4) NOT NULL DEFAULT '' AFTER c_benef")
            if not column_exists("cad_produtos_fornecedor", "cst_icms_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cst_icms_saida VARCHAR(4) NOT NULL DEFAULT '' AFTER origem_saida")
            if not column_exists("cad_produtos_fornecedor", "cfop_saida_empresa"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cfop_saida_empresa VARCHAR(10) NOT NULL DEFAULT '' AFTER cst_icms_saida")
            if not column_exists("cad_produtos_fornecedor", "aliquota_icms_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN aliquota_icms_saida DECIMAL(10,4) NOT NULL DEFAULT 0.0000 AFTER cfop_saida_empresa")
            if not column_exists("cad_produtos_fornecedor", "reducao_bc_icms"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN reducao_bc_icms DECIMAL(10,4) NOT NULL DEFAULT 0.0000 AFTER cst_icms")
            if not column_exists("cad_produtos_fornecedor", "cst_pis_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cst_pis_saida VARCHAR(4) NOT NULL DEFAULT '' AFTER cst_pis")
            if not column_exists("cad_produtos_fornecedor", "cst_cofins_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN cst_cofins_saida VARCHAR(4) NOT NULL DEFAULT '' AFTER cst_cofins")
            if not column_exists("cad_produtos_fornecedor", "natureza_receita_saida"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN natureza_receita_saida VARCHAR(40) NOT NULL DEFAULT '' AFTER cst_cofins_saida")
            if not column_exists("cad_produtos_fornecedor", "chave_nfe_origem"):
                cursor.execute("ALTER TABLE cad_produtos_fornecedor ADD COLUMN chave_nfe_origem VARCHAR(44) NOT NULL DEFAULT '' AFTER codigo_empresa")

        if column_exists("ncm", "id"):
            if index_exists("ncm", "uq_ncm_regra"):
                cursor.execute("ALTER TABLE ncm DROP INDEX uq_ncm_regra")
            if not column_exists("ncm", "cfop_entrada"):
                after_column = "base_legal_pis_cofins" if column_exists("ncm", "base_legal_pis_cofins") else "id"
                cursor.execute(f"ALTER TABLE ncm ADD COLUMN cfop_entrada VARCHAR(20) NOT NULL DEFAULT '' AFTER {after_column}")
            if not column_exists("ncm", "cfop_saida"):
                cursor.execute("ALTER TABLE ncm ADD COLUMN cfop_saida VARCHAR(20) NOT NULL DEFAULT '' AFTER cfop_entrada")
            if column_exists("ncm", "cfop"):
                cursor.execute(
                    """
                    UPDATE ncm
                    SET cfop_entrada = CASE
                            WHEN TRIM(COALESCE(cfop_entrada, '')) = '' THEN TRIM(SUBSTRING_INDEX(COALESCE(cfop, ''), '/', 1))
                            ELSE cfop_entrada
                        END,
                        cfop_saida = CASE
                            WHEN TRIM(COALESCE(cfop_saida, '')) = '' AND LOCATE('/', COALESCE(cfop, '')) > 0 THEN TRIM(SUBSTRING_INDEX(COALESCE(cfop, ''), '/', -1))
                            ELSE cfop_saida
                        END
                    """
                )
            if not index_exists("ncm", "uq_ncm_regra"):
                cursor.execute(
                    """
                    ALTER TABLE ncm
                    ADD UNIQUE KEY uq_ncm_regra (
                        ambiente, uf, regime_tributario, data_vigencia, ncm, cest,
                        cfop_entrada, cfop_saida, cst_csosn, codigo_beneficio_fiscal, ex_tipi
                    )
                    """
                )

    def _only_digits(self, value: object) -> str:
        return "".join(char for char in str(value or "") if char.isdigit())

    def _cfop_digits(self, value: object, max_length: int = 10) -> str:
        """Extrai dígitos de um CFOP, convertendo representação float ("1102.0" → "1102")."""
        text = str(value or "").strip()
        if not text:
            return ""
        try:
            fv = float(text)
            if fv == int(fv):
                text = str(int(fv))
        except (ValueError, TypeError, OverflowError):
            pass
        return "".join(c for c in text if c.isdigit())[:max_length]

    def _decimal_text(self, value: object) -> str:
        text = str(value or "").strip()
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        if not text:
            return "0"
        try:
            return str(Decimal(text))
        except InvalidOperation:
            return "0"

    def _date_text(self, value: object) -> str | None:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()
        text = str(value or "").strip()
        if not text:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt).date().isoformat()
            except ValueError:
                continue
        return None

    def _trim_text(self, value: object, max_length: int) -> str:
        return str(value or "").strip()[:max_length]

    def list_companies(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT id, ambiente, nome, cnpj, inscricao_estadual, observacao
                FROM cad_empresas
                WHERE ambiente = %s
                ORDER BY nome
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_companies_for_confronto(self, environment: str) -> list[dict[str, str]]:
        """
        Retorna lista de empresas para o combo do confronto.
        Tenta cad_empresas primeiro; se vazio, usa sped_perfis como fallback
        para que o confronto funcione mesmo sem produtos importados.
        Cada item: {"nome": str, "cnpj": str}.
        """
        rows = self.list_companies(environment)
        if rows:
            return [{"nome": str(r["nome"] or ""), "cnpj": str(r["cnpj"] or "")} for r in rows]

        # Fallback: empresas presentes nos perfis SPED
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT empresa_nome_sped AS nome, empresa_cnpj_sped AS cnpj
                FROM sped_perfis
                WHERE ambiente = %s
                  AND empresa_cnpj_sped != ''
                GROUP BY empresa_cnpj_sped
                ORDER BY empresa_nome_sped
                """,
                (environment,),
            )
            return [{"nome": str(r["nome"] or ""), "cnpj": str(r["cnpj"] or "")} for r in cursor.fetchall()]
        finally:
            connection.close()

    def save_company(self, environment: str, data: dict[str, object]) -> int:
        company_id = int(data.get("id") or 0)
        values = (
            environment,
            str(data.get("nome", "")).strip(),
            self._only_digits(data.get("cnpj", "")),
            str(data.get("inscricao_estadual", "")).strip(),
            str(data.get("observacao", "")).strip(),
        )
        if not values[1]:
            raise ValueError("Informe o nome da empresa.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if company_id:
                cursor.execute(
                    """
                    UPDATE cad_empresas
                    SET ambiente = %s, nome = %s, cnpj = %s, inscricao_estadual = %s, observacao = %s
                    WHERE id = %s
                    """,
                    (*values, company_id),
                )
                connection.commit()
                return company_id
            cursor.execute(
                """
                INSERT INTO cad_empresas (ambiente, nome, cnpj, inscricao_estadual, observacao)
                VALUES (%s, %s, %s, %s, %s)
                """,
                values,
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def delete_company(self, company_id: int) -> None:
        self._delete_by_id("cad_empresas", company_id)

    def list_suppliers(self, company_id: int) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT id, empresa_id, nome, cnpj, inscricao_estadual, uf, codigo, regime_tributario, observacao
                FROM cad_fornecedores
                WHERE empresa_id = %s
                ORDER BY nome
                """,
                (company_id,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_suppliers_catalog(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    f.id,
                    f.empresa_id,
                    e.nome AS empresa_nome,
                    f.nome,
                    f.cnpj,
                    f.inscricao_estadual,
                    f.uf,
                    f.codigo,
                    f.regime_tributario,
                    f.observacao
                FROM cad_fornecedores f
                INNER JOIN cad_empresas e ON e.id = f.empresa_id
                WHERE e.ambiente = %s
                ORDER BY e.nome, f.nome
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def save_supplier(self, company_id: int, data: dict[str, object]) -> int:
        supplier_id = int(data.get("id") or 0)
        values = (
            company_id,
            str(data.get("nome", "")).strip(),
            self._only_digits(data.get("cnpj", "")),
            str(data.get("inscricao_estadual", "")).strip(),
            self._trim_text(data.get("uf", ""), 2).upper(),
            str(data.get("codigo", "")).strip(),
            self._trim_text(data.get("regime_tributario", "LUCRO_REAL_PRESUMIDO"), 30) or "LUCRO_REAL_PRESUMIDO",
            str(data.get("observacao", "")).strip(),
        )
        if not company_id:
            raise ValueError("Selecione uma empresa para cadastrar fornecedor.")
        if not values[1]:
            raise ValueError("Informe o nome do fornecedor.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if supplier_id:
                cursor.execute(
                    """
                    UPDATE cad_fornecedores
                    SET empresa_id = %s, nome = %s, cnpj = %s, inscricao_estadual = %s, uf = %s, codigo = %s, regime_tributario = %s, observacao = %s
                    WHERE id = %s
                    """,
                    (*values, supplier_id),
                )
                connection.commit()
                return supplier_id
            cursor.execute(
                """
                INSERT INTO cad_fornecedores (empresa_id, nome, cnpj, inscricao_estadual, uf, codigo, regime_tributario, observacao)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                values,
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def delete_supplier(self, supplier_id: int) -> None:
        self._delete_by_id("cad_fornecedores", supplier_id)

    def list_product_types(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT id, ambiente, nome, descricao
                FROM cad_tipos_produto
                WHERE ambiente = %s
                ORDER BY nome
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_product_types_catalog(self, environment: str) -> list[dict[str, object]]:
        return self.list_product_types(environment)

    def save_product_type(self, environment: str, data: dict[str, object]) -> int:
        type_id = int(data.get("id") or 0)
        values = (environment, str(data.get("nome", "")).strip(), str(data.get("descricao", "")).strip())
        if not values[1]:
            raise ValueError("Informe o nome do tipo de produto.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if type_id:
                cursor.execute(
                    """
                    UPDATE cad_tipos_produto
                    SET ambiente = %s, nome = %s, descricao = %s
                    WHERE id = %s
                    """,
                    (*values, type_id),
                )
                connection.commit()
                return type_id
            cursor.execute(
                "INSERT INTO cad_tipos_produto (ambiente, nome, descricao) VALUES (%s, %s, %s)",
                values,
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def delete_product_type(self, type_id: int) -> None:
        self._delete_by_id("cad_tipos_produto", type_id)

    NCM_FIELDS = [
        "atividade",
        "regime_tributario",
        "uf",
        "data_vigencia",
        "ncm",
        "descricao",
        "cest",
        "aliquota_ipi",
        "cst_ipi",
        "ex_tipi",
        "cst_pis_cofins_entrada",
        "cst_pis_cofins_saida",
        "codigo_sped",
        "aliquota_pis",
        "aliquota_cofins",
        "base_legal_pis_cofins",
        "cfop_entrada",
        "cfop_saida",
        "cst_csosn",
        "ad_rem_icms",
        "aliquota_icms",
        "reducao_bc_icms",
        "reducao_bc_icms_st",
        "aliquota_icms_st",
        "aliquota_red_bc_icms",
        "mva",
        "fcp",
        "codigo_beneficio_fiscal",
        "antecipado",
        "percentual_diferimento",
        "percentual_isencao",
        "codigo_anp",
        "base_legal_icms",
    ]

    NCM_NUMERIC_FIELDS = {
        "aliquota_ipi",
        "aliquota_pis",
        "aliquota_cofins",
        "aliquota_icms",
        "reducao_bc_icms",
        "reducao_bc_icms_st",
        "aliquota_icms_st",
        "aliquota_red_bc_icms",
        "mva",
        "fcp",
        "percentual_diferimento",
        "percentual_isencao",
    }

    def _ncm_values(self, environment: str, data: dict[str, object]) -> tuple[object, ...]:
        values: list[object] = [environment]
        for field in self.NCM_FIELDS:
            value = data.get(field, "")
            if field == "data_vigencia":
                values.append(self._date_text(value))
            elif field == "ncm":
                values.append(self._only_digits(value)[:20])
            elif field == "uf":
                values.append(self._trim_text(value, 2).upper())
            elif field in self.NCM_NUMERIC_FIELDS:
                values.append(self._decimal_text(value))
            elif field in {"base_legal_pis_cofins", "base_legal_icms"}:
                values.append(str(value or "").strip())
            else:
                limits = {
                    "descricao": 255,
                    "atividade": 80,
                    "regime_tributario": 80,
                    "cest": 20,
                    "cst_ipi": 10,
                    "ex_tipi": 20,
                    "cst_pis_cofins_entrada": 20,
                    "cst_pis_cofins_saida": 20,
                    "codigo_sped": 40,
                    "cfop_entrada": 20,
                    "cfop_saida": 20,
                    "cst_csosn": 20,
                    "ad_rem_icms": 40,
                    "codigo_beneficio_fiscal": 40,
                    "antecipado": 10,
                    "codigo_anp": 40,
                }
                values.append(self._trim_text(value, limits.get(field, 80)))
        return tuple(values)

    def list_ncm_catalog(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT *
                FROM ncm
                WHERE ambiente = %s
                ORDER BY ncm, uf, regime_tributario, data_vigencia, id
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def save_ncm_rule(self, environment: str, data: dict[str, object]) -> int:
        rule_id = int(data.get("id") or 0)
        values = self._ncm_values(environment, data)
        if not values[5]:
            raise ValueError("Informe o NCM.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            assignments = ", ".join(f"{field} = %s" for field in ["ambiente", *self.NCM_FIELDS])
            if rule_id:
                cursor.execute(
                    f"UPDATE ncm SET {assignments} WHERE id = %s",
                    (*values, rule_id),
                )
                connection.commit()
                return rule_id
            placeholders = ", ".join(["%s"] * len(values))
            columns = ", ".join(["ambiente", *self.NCM_FIELDS])
            cursor.execute(f"INSERT INTO ncm ({columns}) VALUES ({placeholders})", values)
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def delete_ncm_rule(self, rule_id: int) -> None:
        self._delete_by_id("ncm", rule_id)

    def import_ncm_rules_from_excel(self, environment: str, excel_path: Path) -> dict[str, int]:
        from app.parsers.excel_parser import get_first_xlsx_sheet_name, read_xlsx_sheet_rows

        sheet_name = get_first_xlsx_sheet_name(excel_path)
        rows = read_xlsx_sheet_rows(excel_path, sheet_name)
        if len(rows) < 7:
            raise ValueError("Planilha de NCM sem linhas suficientes para importacao.")

        def row_value(row_index: int, column_index: int) -> str:
            row = rows[row_index] if row_index < len(rows) else []
            return str(row[column_index] if column_index < len(row) else "").strip()

        atividade = row_value(1, 1)
        regime = row_value(2, 1)
        uf = row_value(3, 1).upper()[:2]
        data_vigencia = row_value(4, 1)
        header_row = 6
        headers = [str(value or "").strip().upper() for value in rows[header_row - 1]]
        header_map = {
            "NCM": "ncm",
            "DESCRIÇÃO": "descricao",
            "CEST": "cest",
            "ALÍQUOTA IPI": "aliquota_ipi",
            "CST IPI": "cst_ipi",
            "EX": "ex_tipi",
            "CST PIS/COFINS ENTRADA": "cst_pis_cofins_entrada",
            "CST PIS/COFINS SAÍDA": "cst_pis_cofins_saida",
            "CÓDIGO SPED": "codigo_sped",
            "ALÍQUOTA PIS": "aliquota_pis",
            "ALÍQUOTA COFINS": "aliquota_cofins",
            "BASE LEGAL PIS/COFINS": "base_legal_pis_cofins",
            "CFOP": "cfop",
            "CST/CSOSN": "cst_csosn",
            "AD REM ICMS": "ad_rem_icms",
            "ALÍQUOTA ICMS": "aliquota_icms",
            "% RED. BASE DE CÁLCULO ICMS": "reducao_bc_icms",
            "% RED. BASE DE CÁLCULO ICMS ST": "reducao_bc_icms_st",
            "ALÍQUOTA ICMS ST": "aliquota_icms_st",
            "% ALÍQUOTA RED. BASE DE CÁLCULO ICMS": "aliquota_red_bc_icms",
            "MVA": "mva",
            "FCP": "fcp",
            "CÓD. BENEFÍCIO FISCAL": "codigo_beneficio_fiscal",
            "ANTECIPADO": "antecipado",
            "PERCENTUAL DIFERIMENTO": "percentual_diferimento",
            "PERCENTUAL ISENÇÃO": "percentual_isencao",
            "CÓDIGO ANP": "codigo_anp",
            "BASE LEGAL ICMS": "base_legal_icms",
        }
        fields_by_index = {index: header_map[header] for index, header in enumerate(headers) if header in header_map}
        stats = {"rows": 0, "inserted": 0, "updated": 0, "ignored": 0}
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            columns = ["ambiente", *self.NCM_FIELDS]
            placeholders = ", ".join(["%s"] * len(columns))
            update_assignments = ", ".join(f"{field} = VALUES({field})" for field in self.NCM_FIELDS)
            sql = (
                f"INSERT INTO ncm ({', '.join(columns)}) VALUES ({placeholders}) "
                f"ON DUPLICATE KEY UPDATE {update_assignments}"
            )
            values_list: list[tuple[object, ...]] = []
            for row in rows[header_row:]:
                payload = {
                    "atividade": atividade,
                    "regime_tributario": regime,
                    "uf": uf,
                    "data_vigencia": data_vigencia,
                }
                for index, field in fields_by_index.items():
                    raw_value = row[index] if index < len(row) else ""
                    if field == "cfop":
                        cfop_parts = [part.strip() for part in str(raw_value or "").split("/", 1)]
                        payload["cfop_entrada"] = cfop_parts[0] if cfop_parts else ""
                        payload["cfop_saida"] = cfop_parts[1] if len(cfop_parts) > 1 else ""
                    else:
                        payload[field] = raw_value
                if not self._only_digits(payload.get("ncm", "")):
                    stats["ignored"] += 1
                    continue
                values_list.append(self._ncm_values(environment, payload))
                stats["rows"] += 1
            if values_list:
                cursor.executemany(sql, values_list)
                connection.commit()
                stats["inserted"] = int(cursor.rowcount or 0)
            return stats
        finally:
            connection.close()

    def import_reviewed_products_from_excel(
        self,
        environment: str,
        company_id: int,
        excel_path: Path,
        progress_callback: object | None = None,
    ) -> dict[str, object]:
        from app.parsers.excel_parser import normalize_header, read_xlsx_sheet_rows
        from app.exporters.workbook_exporter import write_simple_excel_workbook

        def emit_progress(current: int, total: int, message: str) -> None:
            if progress_callback is None:
                return
            try:
                progress_callback(current, total, message)
            except Exception:
                pass

        emit_progress(0, 1, "Lendo aba BASE_COMPLETA...")
        rows = read_xlsx_sheet_rows(excel_path, "BASE_COMPLETA")
        if len(rows) < 2:
            raise ValueError('A aba "BASE_COMPLETA" nao possui linhas para importacao.')

        header_index = -1
        for index, row in enumerate(rows[:10]):
            normalized_headers = {normalize_header(value) for value in row}
            if {"FORNECEDOR", "CODFORN", "CLASSIFICACAO"}.issubset(normalized_headers):
                header_index = index
                break
        if header_index < 0:
            raise ValueError('Nao foi encontrado cabecalho valido na aba "BASE_COMPLETA".')

        headers = [normalize_header(value) for value in rows[header_index]]
        supplier_index = headers.index("FORNECEDOR") if "FORNECEDOR" in headers else -1
        supplier_code_index = headers.index("CODFORN") if "CODFORN" in headers else -1
        classification_index = headers.index("CLASSIFICACAO") if "CLASSIFICACAO" in headers else -1
        if not company_id:
            raise ValueError("Selecione a empresa que sera atualizada.")
        if supplier_index < 0 or supplier_code_index < 0:
            raise ValueError('As colunas "Fornecedor" e "Cod. Forn." sao obrigatorias para atualizar os produtos.')

        field_by_header = {
            "STATUS": "status_produto",
            "CODEMPRESA": "codigo_empresa",
            "DESCRICAO": "descricao",
            "EAN": "ean",
            "NCM": "ncm",
            "CEST": "cest",
            "ORIGEMENTRADA": "origem_entrada",
            "CSTICMSENTRADA": "cst_icms",
            "REDBCICMS": "reducao_bc_icms",
            "CFOPSAIDAFORNECEDOR": "cfop_saida_fornecedor",
            "ICMSENTRADA": "aliquota_icms",
            "CFOPENTRADAEMPRESA": "cfop_entrada",
            "CSTIPI": "cst_ipi",
            "IPI": "aliquota_ipi",
            "CSTPISENTRADA": "cst_pis",
            "CSTPISCOFINSENTRADAEMPRESA": "cst_pis_cofins",
            "PIS": "aliquota_pis",
            "CSTCOFINSENTRADA": "cst_cofins",
            "COFINS": "aliquota_cofins",
            "MVA": "mva",
            "VALORICMSST": "valor_icms_st",
            "CCLASSTRIB": "c_classtrib",
            "CBENEF": "c_benef",
            "ORIGEMSAIDA": "origem_saida",
            "CSTICMSSAIDA": "cst_icms_saida",
            "CFOPSAIDAEMPRESA": "cfop_saida_empresa",
            "ICMSSAIDA": "aliquota_icms_saida",
            "CSTPISSAIDA": "cst_pis_saida",
            "CSTCOFINSSAIDA": "cst_cofins_saida",
            "CHAVENFEORIGEM": "chave_nfe_origem",
        }
        fields_by_index: dict[int, str] = {}
        natureza_count = 0
        for index, header in enumerate(headers):
            if header == "NATUREZADARECEITA":
                natureza_count += 1
                fields_by_index[index] = "natureza_receita_entrada" if natureza_count == 1 else "natureza_receita_saida"
            elif header in field_by_header:
                fields_by_index[index] = field_by_header[header]

        data_rows = rows[header_index + 1 :]

        def row_value(row: list[str], column_index: int) -> str:
            return str(row[column_index] if column_index >= 0 and column_index < len(row) else "").strip()

        def match_key(supplier_name: object, supplier_code: object) -> tuple[str, str]:
            return (
                " ".join(str(supplier_name or "").split()).upper(),
                " ".join(str(supplier_code or "").split()).upper(),
            )

        emit_progress(0, len(data_rows), "Carregando produtos existentes...")
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    p.*,
                    f.nome AS fornecedor_nome,
                    e.nome AS empresa_nome,
                    COALESCE(t.nome, '') AS tipo_produto,
                    e.ambiente
                FROM cad_produtos_fornecedor p
                INNER JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                INNER JOIN cad_empresas e ON e.id = f.empresa_id
                LEFT JOIN cad_tipos_produto t ON t.id = p.tipo_produto_id
                WHERE e.ambiente = %s
                  AND e.id = %s
                """,
                (environment, company_id),
            )
            existing_products_by_key: dict[tuple[str, str], list[dict[str, object]]] = {}
            for existing_row in cursor.fetchall():
                key = match_key(existing_row.get("fornecedor_nome", ""), existing_row.get("codigo_fornecedor", ""))
                existing_products_by_key.setdefault(key, []).append(dict(existing_row))
        finally:
            connection.close()

        type_cache = {str(row.get("nome", "")).strip().upper(): int(row["id"]) for row in self.list_product_types(environment)}
        field_labels = {
            "status_produto": "Status",
            "codigo_fornecedor": "Cod. Forn.",
            "codigo_empresa": "Cod. Empresa",
            "descricao": "Descricao",
            "ean": "EAN",
            "ncm": "NCM",
            "cest": "CEST",
            "origem_entrada": "Origem (entrada)",
            "cst_icms": "CST ICMS (entrada)",
            "reducao_bc_icms": "% Red BC ICMS",
            "cfop_saida_fornecedor": "CFOP saida fornecedor",
            "aliquota_icms": "% ICMS (entrada)",
            "cfop_entrada": "CFOP entrada empresa",
            "cst_ipi": "CST IPI",
            "aliquota_ipi": "% IPI",
            "cst_pis": "CST PIS (entrada)",
            "cst_pis_cofins": "CST PIS_COFINS (ENTRADA EMPRESA)",
            "aliquota_pis": "% PIS",
            "cst_cofins": "CST COFINS (entrada)",
            "aliquota_cofins": "% COFINS",
            "natureza_receita_entrada": "Natureza da receita entrada",
            "mva": "MVA",
            "valor_icms_st": "Valor ICMS-ST",
            "c_classtrib": "cClassTrib",
            "c_benef": "cBenef",
            "origem_saida": "Origem (saida)",
            "cst_icms_saida": "CST ICMS (saida)",
            "cfop_saida_empresa": "CFOP saida empresa",
            "aliquota_icms_saida": "% ICMS (saida)",
            "cst_pis_saida": "CST PIS (saida)",
            "cst_cofins_saida": "CST COFINS (saida)",
            "natureza_receita_saida": "Natureza da receita saida",
            "chave_nfe_origem": "Chave NFe origem",
        }
        changed_rows: list[list[object]] = []
        stats: dict[str, object] = {
            "rows": 0,
            "updated": 0,
            "ignored": 0,
            "missing_products": 0,
            "duplicate_products": 0,
            "created_types": 0,
            "errors": 0,
            "error_messages": [],
            "changed_values": 0,
            "log_path": "",
        }
        total_data_rows = len(data_rows)
        for processed_index, (row_number, row) in enumerate(enumerate(data_rows, start=header_index + 2), start=1):
            emit_progress(processed_index - 1, total_data_rows, f"Processando linha {row_number}...")
            supplier_name = row_value(row, supplier_index)
            supplier_code = row_value(row, supplier_code_index)
            if not supplier_name or not supplier_code:
                stats["ignored"] = int(stats["ignored"]) + 1
                continue
            matching_products = existing_products_by_key.get(match_key(supplier_name, supplier_code), [])
            if not matching_products:
                stats["missing_products"] = int(stats["missing_products"]) + 1
                continue
            if len(matching_products) > 1:
                stats["duplicate_products"] = int(stats["duplicate_products"]) + 1
                error_messages = list(stats["error_messages"])
                if len(error_messages) < 10:
                    error_messages.append(
                        f"Linha {row_number}: chave duplicada para Fornecedor={supplier_name}, Cod. Forn.={supplier_code}"
                    )
                stats["error_messages"] = error_messages
                continue
            existing = matching_products[0]
            product_id = int(existing["id"])

            payload = dict(existing)
            payload["id"] = product_id
            row_changes: list[list[object]] = []
            for column_index, field in fields_by_index.items():
                new_value = row[column_index] if column_index < len(row) else ""
                old_value = existing.get(field, "")
                payload[field] = new_value
                if str(old_value or "").strip() != str(new_value or "").strip():
                    row_changes.append(
                        [
                            row_number,
                            product_id,
                            existing.get("empresa_nome", ""),
                            supplier_name,
                            supplier_code,
                            existing.get("descricao", ""),
                            field_labels.get(field, field),
                            old_value,
                            new_value,
                        ]
                    )

            classification = row[classification_index].strip() if classification_index >= 0 and classification_index < len(row) else ""
            existing_classification = str(existing.get("tipo_produto", "") or "").strip()
            if classification:
                cache_key = classification.upper()
                type_id = type_cache.get(cache_key)
                if type_id is None:
                    type_id = self.ensure_product_type(environment, classification)
                    type_cache[cache_key] = type_id
                    stats["created_types"] = int(stats["created_types"]) + 1
                payload["tipo_produto_id"] = type_id
            else:
                payload["tipo_produto_id"] = None
            if existing_classification != classification:
                row_changes.append(
                    [
                        row_number,
                        product_id,
                        existing.get("empresa_nome", ""),
                        supplier_name,
                        supplier_code,
                        existing.get("descricao", ""),
                        "Classificacao",
                        existing_classification,
                        classification,
                    ]
                )

            stats["rows"] = int(stats["rows"]) + 1
            try:
                self.save_supplier_product(int(existing.get("fornecedor_id") or 0), payload)
                stats["updated"] = int(stats["updated"]) + 1
                changed_rows.extend(row_changes)
            except Exception as exc:
                stats["errors"] = int(stats["errors"]) + 1
                error_messages = list(stats["error_messages"])
                if len(error_messages) < 10:
                    error_messages.append(f"Linha {row_number}, produto {product_id}: {exc}")
                stats["error_messages"] = error_messages
        if changed_rows:
            emit_progress(total_data_rows, total_data_rows, "Gravando log de alteracoes...")
            log_path = excel_path.with_name(f"{excel_path.stem}_log_alteracoes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            write_simple_excel_workbook(
                log_path,
                [
                    (
                        "Alteracoes",
                        ["Linha", "Produto ID", "Empresa", "Fornecedor", "Cod. Forn.", "Descricao", "Campo", "Valor anterior", "Valor novo"],
                        changed_rows,
                        {"include_total": False},
                    )
                ],
            )
            stats["log_path"] = str(log_path)
            stats["changed_values"] = len(changed_rows)
        emit_progress(total_data_rows, total_data_rows, "Importacao de produtos concluida.")
        return stats

    def ensure_product_type(self, environment: str, type_name: str, description: str = "") -> int:
        normalized_name = str(type_name or "").strip()
        if not normalized_name:
            raise ValueError("Nome da classificacao do produto obrigatorio.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT id
                FROM cad_tipos_produto
                WHERE ambiente = %s
                  AND UPPER(nome) = UPPER(%s)
                LIMIT 1
                """,
                (environment, normalized_name),
            )
            row = cursor.fetchone()
            if row:
                return int(row[0])
            cursor.execute(
                """
                INSERT INTO cad_tipos_produto (ambiente, nome, descricao)
                VALUES (%s, %s, %s)
                """,
                (environment, normalized_name, str(description or "").strip()),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def ensure_company(self, environment: str, name: str, cnpj: str = "", inscricao_estadual: str = "") -> int:
        normalized_name = str(name or "").strip()
        normalized_cnpj = self._only_digits(cnpj)
        normalized_ie = str(inscricao_estadual or "").strip()
        if not normalized_name:
            raise ValueError("Nome da empresa obrigatorio.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_cnpj:
                cursor.execute(
                    """
                    SELECT id
                    FROM cad_empresas
                    WHERE ambiente = %s
                      AND cnpj = %s
                    LIMIT 1
                    """,
                    (environment, normalized_cnpj),
                )
                row = cursor.fetchone()
                if row:
                    if normalized_ie:
                        cursor.execute(
                            """
                            UPDATE cad_empresas
                            SET inscricao_estadual = CASE
                                WHEN TRIM(COALESCE(inscricao_estadual, '')) = '' THEN %s
                                ELSE inscricao_estadual
                            END
                            WHERE id = %s
                            """,
                            (normalized_ie, int(row[0])),
                        )
                        connection.commit()
                    return int(row[0])
            cursor.execute(
                """
                SELECT id
                FROM cad_empresas
                WHERE ambiente = %s
                  AND UPPER(nome) = UPPER(%s)
                LIMIT 1
                """,
                (environment, normalized_name),
            )
            row = cursor.fetchone()
            if row:
                if normalized_ie:
                    cursor.execute(
                        """
                        UPDATE cad_empresas
                        SET inscricao_estadual = CASE
                            WHEN TRIM(COALESCE(inscricao_estadual, '')) = '' THEN %s
                            ELSE inscricao_estadual
                        END
                        WHERE id = %s
                        """,
                        (normalized_ie, int(row[0])),
                    )
                    connection.commit()
                return int(row[0])
            cursor.execute(
                """
                INSERT INTO cad_empresas (ambiente, nome, cnpj, inscricao_estadual, observacao)
                VALUES (%s, %s, %s, %s, 'Importado de XML')
                """,
                (environment, normalized_name, normalized_cnpj, normalized_ie),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def find_company_id(self, environment: str, name: str, cnpj: str = "") -> int | None:
        normalized_name = str(name or "").strip()
        normalized_cnpj = self._only_digits(cnpj)
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_cnpj:
                cursor.execute(
                    """
                    SELECT id
                    FROM cad_empresas
                    WHERE ambiente = %s
                      AND cnpj = %s
                    LIMIT 1
                    """,
                    (environment, normalized_cnpj),
                )
                row = cursor.fetchone()
                if row:
                    return int(row[0])
            cursor.execute(
                """
                SELECT id
                FROM cad_empresas
                WHERE ambiente = %s
                  AND UPPER(nome) = UPPER(%s)
                LIMIT 1
                """,
                (environment, normalized_name),
            )
            row = cursor.fetchone()
            return int(row[0]) if row else None
        finally:
            connection.close()

    def ensure_supplier(
        self,
        company_id: int,
        name: str,
        cnpj: str = "",
        inscricao_estadual: str = "",
        uf: str = "",
        regime_tributario: str = "LUCRO_REAL_PRESUMIDO",
        codigo: str = "",
    ) -> int:
        normalized_name = str(name or "").strip()
        normalized_codigo = str(codigo or "").strip()
        normalized_cnpj = self._only_digits(cnpj)
        normalized_ie = str(inscricao_estadual or "").strip()
        normalized_uf = self._trim_text(uf, 2).upper()
        # Para UPDATE usa vazio (CASE WHEN mantém valor existente); para INSERT usa padrão
        normalized_regime = self._trim_text(regime_tributario, 30)
        insert_regime = normalized_regime or "LUCRO_REAL_PRESUMIDO"
        if not company_id:
            raise ValueError("Empresa obrigatoria para fornecedor.")
        if not normalized_name:
            raise ValueError("Nome do fornecedor obrigatorio.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_cnpj:
                cursor.execute(
                    """
                    SELECT id
                    FROM cad_fornecedores
                    WHERE empresa_id = %s
                      AND cnpj = %s
                    LIMIT 1
                    """,
                    (company_id, normalized_cnpj),
                )
                row = cursor.fetchone()
                if row:
                    supplier_id = int(row[0])
                    # Verifica se outro fornecedor já usa o nome desejado
                    cursor.execute(
                        "SELECT id, cnpj FROM cad_fornecedores WHERE empresa_id = %s AND UPPER(nome) = UPPER(%s) AND id != %s LIMIT 1",
                        (company_id, normalized_name, supplier_id),
                    )
                    nome_conflict = cursor.fetchone()
                    if nome_conflict:
                        conflito_id = int(nome_conflict[0])
                        conflito_cnpj = str(nome_conflict[1] or "").strip()
                        if not conflito_cnpj:
                            # O conflitante não tem CNPJ — é um duplicado criado sem CNPJ.
                            # Preferir o que tem o nome correto; atualizar seu CNPJ e retornar ele.
                            cursor.execute(
                                """
                                UPDATE cad_fornecedores
                                SET cnpj = %s,
                                    inscricao_estadual = CASE WHEN %s <> '' THEN %s ELSE inscricao_estadual END,
                                    uf = CASE WHEN %s <> '' THEN %s ELSE uf END,
                                    regime_tributario = CASE WHEN %s <> '' THEN %s ELSE regime_tributario END,
                                    codigo = CASE WHEN %s <> '' THEN %s ELSE codigo END
                                WHERE id = %s
                                """,
                                (
                                    normalized_cnpj,
                                    normalized_ie, normalized_ie,
                                    normalized_uf, normalized_uf,
                                    normalized_regime, normalized_regime,
                                    normalized_codigo, normalized_codigo,
                                    conflito_id,
                                ),
                            )
                            connection.commit()
                            return conflito_id
                        else:
                            # Conflita com fornecedor de CNPJ diferente (filial distinta).
                            # Atualiza o encontrado por CNPJ sem alterar o nome.
                            cursor.execute(
                                """
                                UPDATE cad_fornecedores
                                SET cnpj = CASE WHEN %s <> '' THEN %s ELSE cnpj END,
                                    inscricao_estadual = CASE WHEN %s <> '' THEN %s ELSE inscricao_estadual END,
                                    uf = CASE WHEN %s <> '' THEN %s ELSE uf END,
                                    regime_tributario = CASE WHEN %s <> '' THEN %s ELSE regime_tributario END,
                                    codigo = CASE WHEN %s <> '' THEN %s ELSE codigo END
                                WHERE id = %s
                                """,
                                (
                                    normalized_cnpj, normalized_cnpj,
                                    normalized_ie, normalized_ie,
                                    normalized_uf, normalized_uf,
                                    normalized_regime, normalized_regime,
                                    normalized_codigo, normalized_codigo,
                                    supplier_id,
                                ),
                            )
                    else:
                        # Nome disponível — atualiza com o nome correto
                        cursor.execute(
                            """
                            UPDATE cad_fornecedores
                            SET nome = %s,
                                cnpj = CASE WHEN %s <> '' THEN %s ELSE cnpj END,
                                inscricao_estadual = CASE WHEN %s <> '' THEN %s ELSE inscricao_estadual END,
                                uf = CASE WHEN %s <> '' THEN %s ELSE uf END,
                                regime_tributario = CASE WHEN %s <> '' THEN %s ELSE regime_tributario END,
                                codigo = CASE WHEN %s <> '' THEN %s ELSE codigo END
                            WHERE id = %s
                            """,
                            (
                                normalized_name,
                                normalized_cnpj, normalized_cnpj,
                                normalized_ie, normalized_ie,
                                normalized_uf, normalized_uf,
                                normalized_regime, normalized_regime,
                                normalized_codigo, normalized_codigo,
                                supplier_id,
                            ),
                        )
                    connection.commit()
                    return supplier_id
            cursor.execute(
                """
                SELECT id
                FROM cad_fornecedores
                WHERE empresa_id = %s
                  AND UPPER(nome) = UPPER(%s)
                LIMIT 1
                """,
                (company_id, normalized_name),
            )
            row = cursor.fetchone()
            if row:
                supplier_id = int(row[0])
                cursor.execute(
                    """
                    UPDATE cad_fornecedores
                    SET cnpj = CASE WHEN %s <> '' THEN %s ELSE cnpj END,
                        inscricao_estadual = CASE WHEN %s <> '' THEN %s ELSE inscricao_estadual END,
                        uf = CASE WHEN %s <> '' THEN %s ELSE uf END,
                        regime_tributario = CASE WHEN %s <> '' THEN %s ELSE regime_tributario END,
                        codigo = CASE WHEN %s <> '' THEN %s ELSE codigo END
                    WHERE id = %s
                    """,
                    (
                        normalized_cnpj, normalized_cnpj,
                        normalized_ie, normalized_ie,
                        normalized_uf, normalized_uf,
                        normalized_regime, normalized_regime,
                        normalized_codigo, normalized_codigo,
                        supplier_id,
                    ),
                )
                connection.commit()
                return supplier_id
            try:
                cursor.execute(
                    """
                    INSERT INTO cad_fornecedores (empresa_id, nome, cnpj, inscricao_estadual, uf, codigo, regime_tributario, observacao)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, 'Importado de XML')
                    """,
                    (company_id, normalized_name, normalized_cnpj, normalized_ie, normalized_uf, normalized_codigo, insert_regime),
                )
                connection.commit()
                return int(cursor.lastrowid)
            except MySQLError as exc:
                if exc.errno == 1062:
                    cursor.execute(
                        "SELECT id FROM cad_fornecedores WHERE empresa_id = %s AND UPPER(nome) = UPPER(%s) LIMIT 1",
                        (company_id, normalized_name),
                    )
                    dup = cursor.fetchone()
                    if dup:
                        return int(dup[0])
                raise
        finally:
            connection.close()

    def find_supplier_id(self, company_id: int, name: str, cnpj: str = "") -> int | None:
        normalized_name = str(name or "").strip()
        normalized_cnpj = self._only_digits(cnpj)
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_cnpj:
                cursor.execute(
                    """
                    SELECT id
                    FROM cad_fornecedores
                    WHERE empresa_id = %s
                      AND cnpj = %s
                    LIMIT 1
                    """,
                    (company_id, normalized_cnpj),
                )
                row = cursor.fetchone()
                if row:
                    return int(row[0])
            cursor.execute(
                """
                SELECT id
                FROM cad_fornecedores
                WHERE empresa_id = %s
                  AND UPPER(nome) = UPPER(%s)
                LIMIT 1
                """,
                (company_id, normalized_name),
            )
            row = cursor.fetchone()
            return int(row[0]) if row else None
        finally:
            connection.close()

    def list_supplier_products(self, supplier_id: int) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    p.*,
                    COALESCE(t.nome, '') AS tipo_produto
                FROM cad_produtos_fornecedor p
                LEFT JOIN cad_tipos_produto t ON t.id = p.tipo_produto_id
                WHERE p.fornecedor_id = %s
                ORDER BY p.codigo_fornecedor, p.descricao
                """,
                (supplier_id,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def list_products_catalog(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    p.*,
                    f.empresa_id,
                    f.nome AS fornecedor_nome,
                    f.uf AS fornecedor_uf,
                    e.nome AS empresa_nome,
                    e.cnpj AS empresa_cnpj,
                    COALESCE(t.nome, '') AS tipo_produto
                FROM cad_produtos_fornecedor p
                INNER JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                INNER JOIN cad_empresas e ON e.id = f.empresa_id
                LEFT JOIN cad_tipos_produto t ON t.id = p.tipo_produto_id
                WHERE e.ambiente = %s
                ORDER BY e.nome, f.nome, p.codigo_fornecedor, p.descricao
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def save_supplier_product(self, supplier_id: int, data: dict[str, object]) -> int:
        product_id = int(data.get("id") or 0)
        if not supplier_id:
            raise ValueError("Selecione um fornecedor para cadastrar produto.")
        if not str(data.get("codigo_fornecedor", "")).strip():
            raise ValueError("Informe o codigo do produto do fornecedor.")
        cst_pis = self._trim_text(data.get("cst_pis", data.get("cst_pis_cofins", "")), 4)
        cst_cofins = self._trim_text(data.get("cst_cofins", data.get("cst_pis_cofins", "")), 4)
        aliquota_pis = self._decimal_text(data.get("aliquota_pis", data.get("aliquota_pis_cofins", "")))
        aliquota_cofins = self._decimal_text(data.get("aliquota_cofins", data.get("aliquota_pis_cofins", "")))
        cst_pis_cofins = self._trim_text(data.get("cst_pis_cofins", ""), 20)
        aliquota_pis_cofins = self._decimal_text(data.get("aliquota_pis_cofins", aliquota_pis))

        normalized_code = self._trim_text(data.get("codigo_fornecedor", ""), 80)
        normalized_ean = self._only_digits(data.get("ean", ""))[:30]
        values = (
            supplier_id,
            int(data.get("tipo_produto_id") or 0) or None,
            normalized_code,
            self._trim_text(data.get("codigo_empresa", ""), 80),
            self._trim_text(data.get("chave_nfe_origem", ""), 44),
            self._trim_text(data.get("status_produto", ""), 40),
            self._trim_text(data.get("descricao", ""), 255),
            normalized_ean,
            normalized_ean or None,
            self._only_digits(data.get("ncm", ""))[:20],
            self._only_digits(data.get("cest", ""))[:20],
            self._trim_text(data.get("origem_entrada", ""), 4),
            self._cfop_digits(data.get("cfop_saida_fornecedor", "")),
            self._cfop_digits(data.get("cfop_entrada", "")),
            self._cfop_digits(data.get("cfop_saida", "")),
            self._trim_text(data.get("natureza_receita_entrada", ""), 40),
            self._trim_text(data.get("c_classtrib", ""), 20),
            self._trim_text(data.get("c_benef", ""), 20),
            self._trim_text(data.get("origem_saida", ""), 4),
            self._trim_text(data.get("cst_icms_saida", ""), 4),
            self._cfop_digits(data.get("cfop_saida_empresa", "")),
            self._decimal_text(data.get("aliquota_icms_saida", "")),
            self._trim_text(data.get("cst_icms", ""), 4),
            self._decimal_text(data.get("reducao_bc_icms", "")),
            self._decimal_text(data.get("aliquota_icms", "")),
            self._trim_text(data.get("cst_ipi", ""), 4),
            self._decimal_text(data.get("aliquota_ipi", "")),
            cst_pis_cofins,
            aliquota_pis_cofins,
            cst_pis,
            self._trim_text(data.get("cst_pis_saida", ""), 4),
            cst_cofins,
            self._trim_text(data.get("cst_cofins_saida", ""), 4),
            self._trim_text(data.get("natureza_receita_saida", ""), 40),
            aliquota_pis,
            aliquota_cofins,
            self._decimal_text(data.get("bc_st", "")),
            self._decimal_text(data.get("mva", "")),
            self._decimal_text(data.get("valor_icms_st", "")),
            self._decimal_text(data.get("aliquota_icms_st", "")),
        )
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT id
                FROM cad_produtos_fornecedor
                WHERE fornecedor_id = %s
                  AND codigo_fornecedor = %s
                  AND id <> %s
                LIMIT 1
                """,
                (supplier_id, normalized_code, product_id),
            )
            if cursor.fetchone() is not None:
                raise ValueError("Ja existe produto com o mesmo Codigo do Fornecedor para este fornecedor.")
            if normalized_ean:
                cursor.execute(
                    """
                    SELECT id
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND ean = %s
                      AND id <> %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_ean, product_id),
                )
                if cursor.fetchone() is not None:
                    raise ValueError("Ja existe produto com o mesmo EAN para este fornecedor.")
            if product_id:
                cursor.execute(
                    """
                    UPDATE cad_produtos_fornecedor
                    SET fornecedor_id = %s, tipo_produto_id = %s, codigo_fornecedor = %s, codigo_empresa = %s,
                        chave_nfe_origem = %s, status_produto = %s, descricao = %s, ean = %s, ean_unico = %s, ncm = %s, cest = %s, origem_entrada = %s, cfop_saida_fornecedor = %s, cfop_entrada = %s, cfop_saida = %s, natureza_receita_entrada = %s, c_classtrib = %s, c_benef = %s, origem_saida = %s, cst_icms_saida = %s, cfop_saida_empresa = %s, aliquota_icms_saida = %s,
                        cst_icms = %s, reducao_bc_icms = %s, aliquota_icms = %s, cst_ipi = %s, aliquota_ipi = %s,
                        cst_pis_cofins = %s, aliquota_pis_cofins = %s, cst_pis = %s, cst_pis_saida = %s, cst_cofins = %s, cst_cofins_saida = %s, natureza_receita_saida = %s,
                        aliquota_pis = %s, aliquota_cofins = %s, bc_st = %s, mva = %s,
                        valor_icms_st = %s, aliquota_icms_st = %s
                    WHERE id = %s
                    """,
                    (*values, product_id),
                )
                connection.commit()
                return product_id
            cursor.execute(
                """
                INSERT INTO cad_produtos_fornecedor (
                    fornecedor_id, tipo_produto_id, codigo_fornecedor, codigo_empresa, chave_nfe_origem, status_produto, descricao, ean, ean_unico, ncm, cest, origem_entrada, cfop_saida_fornecedor, cfop_entrada, cfop_saida, natureza_receita_entrada,
                    c_classtrib, c_benef, origem_saida, cst_icms_saida, cfop_saida_empresa, aliquota_icms_saida, cst_icms, reducao_bc_icms, aliquota_icms, cst_ipi, aliquota_ipi,
                    cst_pis_cofins, aliquota_pis_cofins, cst_pis, cst_pis_saida, cst_cofins, cst_cofins_saida, natureza_receita_saida, aliquota_pis, aliquota_cofins,
                    bc_st, mva, valor_icms_st, aliquota_icms_st
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s
                )
                """,
                values,
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def fetch_supplier_product_by_fornecedor_code(
        self, supplier_id: int, codigo_fornecedor: str
    ) -> dict | None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT id, codigo_empresa
                FROM cad_produtos_fornecedor
                WHERE fornecedor_id = %s AND codigo_fornecedor = %s
                LIMIT 1
                """,
                (supplier_id, str(codigo_fornecedor).strip()),
            )
            return cursor.fetchone()
        finally:
            connection.close()

    def update_produto_codigo_fornecedor(self, product_id: int, new_codigo: str) -> None:
        """Atualiza apenas codigo_fornecedor de um produto, sem alterar outros campos."""
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "UPDATE cad_produtos_fornecedor SET codigo_fornecedor = %s WHERE id = %s",
                (str(new_codigo).strip()[:80], product_id),
            )
            connection.commit()
        finally:
            connection.close()

    def get_catalog_products_by_empresa_id(self, empresa_id: int) -> list[dict]:
        """Retorna produtos do catálogo de uma empresa (por empresa_id, não CNPJ)."""
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT p.id, p.fornecedor_id, p.codigo_fornecedor, p.codigo_empresa,
                       p.descricao, p.ean, p.ncm,
                       f.nome AS fornecedor_nome
                FROM cad_produtos_fornecedor p
                JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                WHERE f.empresa_id = %s
                """,
                (empresa_id,),
            )
            return cursor.fetchall() or []
        finally:
            connection.close()

    def _make_product_row(self, supplier_id: int, data: dict) -> tuple:
        """Monta a tupla de valores para INSERT/UPDATE em cad_produtos_fornecedor."""
        cst_pis    = self._trim_text(data.get("cst_pis",    data.get("cst_pis_cofins",    "")), 4)
        cst_cofins = self._trim_text(data.get("cst_cofins", data.get("cst_pis_cofins",    "")), 4)
        aliquota_pis    = self._decimal_text(data.get("aliquota_pis",    data.get("aliquota_pis_cofins", "")))
        aliquota_cofins = self._decimal_text(data.get("aliquota_cofins", data.get("aliquota_pis_cofins", "")))
        cst_pis_cofins      = self._trim_text(data.get("cst_pis_cofins", ""), 20)
        aliquota_pis_cofins = self._decimal_text(data.get("aliquota_pis_cofins", aliquota_pis))
        normalized_ean  = self._only_digits(data.get("ean", ""))[:30]
        normalized_code = self._trim_text(data.get("codigo_fornecedor", ""), 80)
        return (
            supplier_id,
            int(data.get("tipo_produto_id") or 0) or None,
            normalized_code,
            self._trim_text(data.get("codigo_empresa", ""), 80),
            self._trim_text(data.get("chave_nfe_origem", ""), 44),
            self._trim_text(data.get("status_produto", ""), 40),
            self._trim_text(data.get("descricao", ""), 255),
            normalized_ean,
            normalized_ean or None,
            self._only_digits(data.get("ncm", ""))[:20],
            self._only_digits(data.get("cest", ""))[:20],
            self._trim_text(data.get("origem_entrada", ""), 4),
            self._cfop_digits(data.get("cfop_saida_fornecedor", "")),
            self._cfop_digits(data.get("cfop_entrada", "")),
            self._cfop_digits(data.get("cfop_saida", "")),
            self._trim_text(data.get("natureza_receita_entrada", ""), 40),
            self._trim_text(data.get("c_classtrib", ""), 20),
            self._trim_text(data.get("c_benef", ""), 20),
            self._trim_text(data.get("origem_saida", ""), 4),
            self._trim_text(data.get("cst_icms_saida", ""), 4),
            self._cfop_digits(data.get("cfop_saida_empresa", "")),
            self._decimal_text(data.get("aliquota_icms_saida", "")),
            self._trim_text(data.get("cst_icms", ""), 4),
            self._decimal_text(data.get("reducao_bc_icms", "")),
            self._decimal_text(data.get("aliquota_icms", "")),
            self._trim_text(data.get("cst_ipi", ""), 4),
            self._decimal_text(data.get("aliquota_ipi", "")),
            cst_pis_cofins,
            aliquota_pis_cofins,
            cst_pis,
            self._trim_text(data.get("cst_pis_saida", ""), 4),
            cst_cofins,
            self._trim_text(data.get("cst_cofins_saida", ""), 4),
            self._trim_text(data.get("natureza_receita_saida", ""), 40),
            aliquota_pis,
            aliquota_cofins,
            self._decimal_text(data.get("bc_st", "")),
            self._decimal_text(data.get("mva", "")),
            self._decimal_text(data.get("valor_icms_st", "")),
            self._decimal_text(data.get("aliquota_icms_st", "")),
        )

    _INSERT_PRODUCT_SQL = """
        INSERT INTO cad_produtos_fornecedor (
            fornecedor_id, tipo_produto_id, codigo_fornecedor, codigo_empresa, chave_nfe_origem,
            status_produto, descricao, ean, ean_unico, ncm, cest, origem_entrada,
            cfop_saida_fornecedor, cfop_entrada, cfop_saida, natureza_receita_entrada,
            c_classtrib, c_benef, origem_saida, cst_icms_saida, cfop_saida_empresa,
            aliquota_icms_saida, cst_icms, reducao_bc_icms, aliquota_icms, cst_ipi, aliquota_ipi,
            cst_pis_cofins, aliquota_pis_cofins, cst_pis, cst_pis_saida, cst_cofins,
            cst_cofins_saida, natureza_receita_saida, aliquota_pis, aliquota_cofins,
            bc_st, mva, valor_icms_st, aliquota_icms_st
        ) VALUES (
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s
        )
    """

    _UPDATE_PRODUCT_SQL = """
        UPDATE cad_produtos_fornecedor
        SET fornecedor_id=%s, tipo_produto_id=%s, codigo_fornecedor=%s, codigo_empresa=%s,
            chave_nfe_origem=%s, status_produto=%s, descricao=%s, ean=%s, ean_unico=%s,
            ncm=%s, cest=%s, origem_entrada=%s, cfop_saida_fornecedor=%s, cfop_entrada=%s,
            cfop_saida=%s, natureza_receita_entrada=%s, c_classtrib=%s, c_benef=%s,
            origem_saida=%s, cst_icms_saida=%s, cfop_saida_empresa=%s, aliquota_icms_saida=%s,
            cst_icms=%s, reducao_bc_icms=%s, aliquota_icms=%s, cst_ipi=%s, aliquota_ipi=%s,
            cst_pis_cofins=%s, aliquota_pis_cofins=%s, cst_pis=%s, cst_pis_saida=%s,
            cst_cofins=%s, cst_cofins_saida=%s, natureza_receita_saida=%s,
            aliquota_pis=%s, aliquota_cofins=%s, bc_st=%s, mva=%s, valor_icms_st=%s,
            aliquota_icms_st=%s
        WHERE id=%s
    """

    def bulk_upsert_supplier_products(
        self,
        records: list[dict],
    ) -> list[int]:
        """Insere/atualiza múltiplos produtos em uma única conexão + transação.

        Cada record: {'supplier_id': int, 'data': dict, 'existing_id': int|None}
        Retorna lista de IDs na mesma ordem de records.
        """
        ids: list[int] = [0] * len(records)
        if not records:
            return ids

        to_insert = [(i, r) for i, r in enumerate(records) if not r["existing_id"]]
        to_update = [(i, r) for i, r in enumerate(records) if r["existing_id"]]

        connection = self.get_connection()
        try:
            cursor = connection.cursor()

            for i, rec in to_insert:
                row = self._make_product_row(rec["supplier_id"], rec["data"])
                cursor.execute(self._INSERT_PRODUCT_SQL, row)
                ids[i] = int(cursor.lastrowid)

            for i, rec in to_update:
                row = self._make_product_row(rec["supplier_id"], rec["data"])
                cursor.execute(self._UPDATE_PRODUCT_SQL, (*row, rec["existing_id"]))
                ids[i] = rec["existing_id"]

            connection.commit()
        finally:
            connection.close()

        return ids

    def delete_supplier_product(self, product_id: int) -> None:
        self._delete_by_id("cad_produtos_fornecedor", product_id)

    def upsert_supplier_product_by_ean(
        self,
        supplier_id: int,
        product_data: dict[str, object],
        allow_update_existing: bool = True,
    ) -> tuple[int, bool]:
        normalized_ean = self._only_digits(product_data.get("ean", ""))
        normalized_code = str(product_data.get("codigo_fornecedor", "")).strip()
        row = None
        existing_type_id: int | None = None
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_ean:
                cursor.execute(
                    """
                    SELECT id, tipo_produto_id
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND ean = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_ean),
                )
                row = cursor.fetchone()
            if row is None and normalized_code:
                cursor.execute(
                    """
                    SELECT id, tipo_produto_id
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND codigo_fornecedor = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_code),
                )
                row = cursor.fetchone()
            if row:
                existing_type_id = int(row[1]) if row[1] is not None else None
        finally:
            connection.close()

        if row and not allow_update_existing:
            return int(row[0]), True

        payload = dict(product_data)
        payload["id"] = int(row[0]) if row else 0
        payload["ean"] = normalized_ean
        payload["codigo_fornecedor"] = normalized_code
        if row and existing_type_id is not None:
            payload["tipo_produto_id"] = existing_type_id
        product_id = self.save_supplier_product(supplier_id, payload)
        return product_id, bool(row)

    def supplier_product_exists(self, supplier_id: int, ean: str, code: str) -> bool:
        normalized_ean = self._only_digits(ean)
        normalized_code = str(code or "").strip()
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if normalized_ean:
                cursor.execute(
                    """
                    SELECT 1
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND ean = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_ean),
                )
                if cursor.fetchone() is not None:
                    return True
            if normalized_code:
                cursor.execute(
                    """
                    SELECT 1
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND codigo_fornecedor = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_code),
                )
                return cursor.fetchone() is not None
            return False
        finally:
            connection.close()

    def find_supplier_product_by_key(self, supplier_id: int, ean: str, code: str) -> dict[str, object] | None:
        normalized_ean = self._only_digits(ean)
        normalized_code = str(code or "").strip()
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            if normalized_ean:
                cursor.execute(
                    """
                    SELECT *
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND ean = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_ean),
                )
                row = cursor.fetchone()
                if row:
                    return dict(row)
            if normalized_code:
                cursor.execute(
                    """
                    SELECT *
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND codigo_fornecedor = %s
                    LIMIT 1
                    """,
                    (supplier_id, normalized_code),
                )
                row = cursor.fetchone()
                if row:
                    return dict(row)
            return None
        finally:
            connection.close()

    def find_supplier_product_duplicates(self, supplier_id: int, ean: str, code: str) -> list[dict[str, object]]:
        normalized_ean = self._only_digits(ean)
        normalized_code = str(code or "").strip()
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            if normalized_ean:
                cursor.execute(
                    """
                    SELECT id, fornecedor_id, codigo_fornecedor, ean, descricao
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND ean = %s
                    ORDER BY id
                    """,
                    (supplier_id, normalized_ean),
                )
                rows = [dict(row) for row in cursor.fetchall()]
                if len(rows) > 1:
                    return rows
            if normalized_code:
                cursor.execute(
                    """
                    SELECT id, fornecedor_id, codigo_fornecedor, ean, descricao
                    FROM cad_produtos_fornecedor
                    WHERE fornecedor_id = %s
                      AND codigo_fornecedor = %s
                    ORDER BY id
                    """,
                    (supplier_id, normalized_code),
                )
                rows = [dict(row) for row in cursor.fetchall()]
                if len(rows) > 1:
                    return rows
            return []
        finally:
            connection.close()

    def backup_supplier_products_by_ids(self, product_ids: list[int]) -> Path:
        from datetime import datetime
        if not product_ids:
            raise ValueError("Nenhum produto selecionado para backup.")
        unique_ids = sorted({int(product_id) for product_id in product_ids if int(product_id or 0)})
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            rows: list[dict[str, object]] = []
            chunk_size = 800
            for index in range(0, len(unique_ids), chunk_size):
                chunk = unique_ids[index:index + chunk_size]
                placeholders = ", ".join(["%s"] * len(chunk))
                cursor.execute(
                    f"SELECT * FROM cad_produtos_fornecedor WHERE id IN ({placeholders}) ORDER BY id",
                    tuple(chunk),
                )
                rows.extend(dict(row) for row in cursor.fetchall())
        finally:
            connection.close()
        backup_dir = self.config_path.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        output = backup_dir / f"backup_cad_produtos_fornecedor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        output.write_text(json.dumps(rows, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return output

    def list_supplier_product_duplicates_catalog(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            rows: list[dict[str, object]] = []
            cursor.execute(
                """
                SELECT
                    p.id,
                    p.fornecedor_id,
                    e.nome AS empresa_nome,
                    f.nome AS fornecedor_nome,
                    p.codigo_fornecedor,
                    p.ean,
                    p.descricao,
                    'codigo_fornecedor' AS duplicate_type
                FROM cad_produtos_fornecedor p
                INNER JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                INNER JOIN cad_empresas e ON e.id = f.empresa_id
                INNER JOIN (
                    SELECT fornecedor_id, codigo_fornecedor
                    FROM cad_produtos_fornecedor
                    WHERE TRIM(COALESCE(codigo_fornecedor, '')) <> ''
                    GROUP BY fornecedor_id, codigo_fornecedor
                    HAVING COUNT(*) > 1
                ) d ON d.fornecedor_id = p.fornecedor_id AND d.codigo_fornecedor = p.codigo_fornecedor
                WHERE e.ambiente = %s
                ORDER BY e.nome, f.nome, p.codigo_fornecedor, p.id
                """,
                (environment,),
            )
            rows.extend([dict(row) for row in cursor.fetchall()])
            cursor.execute(
                """
                SELECT
                    p.id,
                    p.fornecedor_id,
                    e.nome AS empresa_nome,
                    f.nome AS fornecedor_nome,
                    p.codigo_fornecedor,
                    p.ean,
                    p.descricao,
                    'ean' AS duplicate_type
                FROM cad_produtos_fornecedor p
                INNER JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                INNER JOIN cad_empresas e ON e.id = f.empresa_id
                INNER JOIN (
                    SELECT fornecedor_id, ean
                    FROM cad_produtos_fornecedor
                    WHERE TRIM(COALESCE(ean, '')) <> ''
                    GROUP BY fornecedor_id, ean
                    HAVING COUNT(*) > 1
                ) d ON d.fornecedor_id = p.fornecedor_id AND d.ean = p.ean
                WHERE e.ambiente = %s
                ORDER BY e.nome, f.nome, p.ean, p.id
                """,
                (environment,),
            )
            rows.extend([dict(row) for row in cursor.fetchall()])
            return rows
        finally:
            connection.close()

    def delete_supplier_products_by_ids(self, product_ids: list[int]) -> int:
        if not product_ids:
            return 0
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            placeholders = ", ".join(["%s"] * len(product_ids))
            cursor.execute(f"DELETE FROM cad_produtos_fornecedor WHERE id IN ({placeholders})", tuple(int(pid) for pid in product_ids))
            connection.commit()
            return int(cursor.rowcount or 0)
        finally:
            connection.close()

    def _delete_by_id(self, table_name: str, row_id: int) -> None:
        allowed_tables = {"cad_empresas", "cad_fornecedores", "cad_tipos_produto", "cad_produtos_fornecedor", "ncm"}
        if table_name not in allowed_tables:
            raise ValueError("Tabela nao permitida para exclusao.")
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(f"DELETE FROM {table_name} WHERE id = %s", (int(row_id),))
            connection.commit()
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


    def get_sped_catalog_check(self, sped_arquivo_id: int) -> list[dict]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    s.codigo            AS codigo_sped,
                    s.descricao         AS descricao_sped,
                    s.ncm               AS ncm_sped,
                    s.cst_icms          AS cst_icms_sped,
                    s.aliquota_icms     AS aliquota_icms_sped,
                    s.cest              AS cest_sped,
                    IF(p.id IS NOT NULL, 'Cadastrado', 'Nao Cadastrado') AS status,
                    p.id                AS cad_id,
                    p.codigo_empresa    AS codigo_empresa_cad,
                    p.codigo_fornecedor AS codigo_fornecedor_cad,
                    p.descricao         AS descricao_cad,
                    f.nome              AS fornecedor,
                    f.cnpj              AS fornecedor_cnpj,
                    p.ncm               AS ncm_cad,
                    p.cest              AS cest_cad,
                    p.cst_icms          AS cst_icms_cad,
                    p.aliquota_icms     AS aliquota_icms_cad,
                    p.cst_pis           AS cst_pis_cad,
                    p.cst_cofins        AS cst_cofins_cad,
                    p.aliquota_pis      AS aliquota_pis_cad,
                    p.aliquota_cofins   AS aliquota_cofins_cad,
                    IF(p.chave_nfe_origem IS NOT NULL AND p.chave_nfe_origem != '', 'Sim', 'Nao') AS via_xml
                FROM sped_produtos_0200 s
                LEFT JOIN sped_arquivos sa ON sa.id = s.sped_arquivo_id
                LEFT JOIN sped_perfis sp ON sp.id = sa.perfil_id
                LEFT JOIN cad_empresas e ON e.cnpj = sp.empresa_cnpj_sped
                LEFT JOIN cad_fornecedores f ON f.empresa_id = e.id
                LEFT JOIN cad_produtos_fornecedor p
                    ON p.fornecedor_id = f.id
                   AND p.codigo_empresa = s.codigo
                WHERE s.sped_arquivo_id = %s
                ORDER BY
                    IF(p.id IS NOT NULL, 0, 1),
                    s.codigo,
                    f.nome
                """,
                (sped_arquivo_id,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def get_catalog_products_by_company_cnpj(self, environment: str, company_cnpj: str) -> list[dict]:
        """
        Retorna todos os produtos do cadastro (cad_produtos_fornecedor) vinculados
        a uma empresa identificada pelo CNPJ, incluindo nome do fornecedor.
        Usado na conferencia SPED x Cadastro.
        """
        digits_only = "".join(c for c in str(company_cnpj or "") if c.isdigit())
        if not digits_only:
            return []
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    p.*,
                    f.nome AS fornecedor_nome,
                    f.cnpj AS fornecedor_cnpj,
                    f.uf AS fornecedor_uf
                FROM cad_produtos_fornecedor p
                JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                JOIN cad_empresas e ON e.id = f.empresa_id
                WHERE e.cnpj = %s
                ORDER BY p.codigo_empresa, f.nome
                """,
                (digits_only,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def get_cfop_map_by_empresa_id(self, empresa_id: int) -> dict[str, str]:
        """
        Retorna {codigo_empresa: cfop_entrada} para todos os produtos da empresa
        que já tenham cfop_entrada preenchido (de qualquer fornecedor).
        Usado na importação para herdar CFOP de outro fornecedor com mesmo código.
        """
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT p.codigo_empresa, p.cfop_entrada
                FROM cad_produtos_fornecedor p
                JOIN cad_fornecedores f ON f.id = p.fornecedor_id
                WHERE f.empresa_id = %s
                  AND p.codigo_empresa != ''
                  AND p.cfop_entrada   != ''
                ORDER BY p.codigo_empresa
                """,
                (empresa_id,),
            )
            result: dict[str, str] = {}
            for row in cursor.fetchall():
                code = str(row["codigo_empresa"] or "").strip()
                cfop = str(row["cfop_entrada"] or "").strip()
                if code and cfop and code not in result:
                    result[code] = cfop
            return result
        finally:
            connection.close()

    @staticmethod
    def hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
        normalized_password = str(password or "")
        normalized_salt = salt or secrets.token_hex(16)
        digest = hashlib.sha256(f"{normalized_salt}:{normalized_password}".encode("utf-8")).hexdigest()
        return normalized_salt, digest

    def list_system_users(self, environment: str) -> list[dict[str, object]]:
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    u.id,
                    u.nome,
                    u.login,
                    u.ativo,
                    u.observacao,
                    u.created_at,
                    u.updated_at,
                    GROUP_CONCAT(p.permissao ORDER BY p.permissao SEPARATOR ', ') AS permissoes
                FROM sistema_usuarios u
                LEFT JOIN sistema_usuario_permissoes p
                    ON p.usuario_id = u.id
                   AND p.permitido = 1
                WHERE u.ambiente = %s
                GROUP BY u.id, u.nome, u.login, u.ativo, u.observacao, u.created_at, u.updated_at
                ORDER BY u.nome, u.login
                """,
                (environment,),
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            connection.close()

    def get_system_user_permissions(self, user_id: int) -> set[str]:
        if not user_id:
            return set()
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT permissao
                FROM sistema_usuario_permissoes
                WHERE usuario_id = %s
                  AND permitido = 1
                """,
                (user_id,),
            )
            return {str(row[0]) for row in cursor.fetchall()}
        finally:
            connection.close()

    def save_system_user(
        self,
        environment: str,
        user_id: int | None,
        data: dict[str, object],
        permissions: set[str],
    ) -> int:
        name = str(data.get("nome") or "").strip()
        login = str(data.get("login") or "").strip().lower()
        password = str(data.get("senha") or "")
        active = 1 if data.get("ativo", True) else 0
        note = str(data.get("observacao") or "").strip()
        if not name:
            raise ValueError("Informe o nome do usuario.")
        if not login:
            raise ValueError("Informe o login do usuario.")
        if not user_id and not password:
            raise ValueError("Informe a senha inicial do usuario.")

        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            if user_id:
                if password:
                    salt, password_hash = self.hash_password(password)
                    cursor.execute(
                        """
                        UPDATE sistema_usuarios
                           SET nome = %s,
                               login = %s,
                               senha_hash = %s,
                               senha_salt = %s,
                               ativo = %s,
                               observacao = %s
                         WHERE id = %s
                           AND ambiente = %s
                        """,
                        (name, login, password_hash, salt, active, note, user_id, environment),
                    )
                else:
                    cursor.execute(
                        """
                        UPDATE sistema_usuarios
                           SET nome = %s,
                               login = %s,
                               ativo = %s,
                               observacao = %s
                         WHERE id = %s
                           AND ambiente = %s
                        """,
                        (name, login, active, note, user_id, environment),
                    )
                saved_id = int(user_id)
            else:
                salt, password_hash = self.hash_password(password)
                cursor.execute(
                    """
                    INSERT INTO sistema_usuarios (
                        ambiente, nome, login, senha_hash, senha_salt, ativo, observacao
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """,
                    (environment, name, login, password_hash, salt, active, note),
                )
                saved_id = int(cursor.lastrowid)

            cursor.execute("DELETE FROM sistema_usuario_permissoes WHERE usuario_id = %s", (saved_id,))
            permission_rows = [(saved_id, permission, 1) for permission in sorted(permissions)]
            if permission_rows:
                cursor.executemany(
                    """
                    INSERT INTO sistema_usuario_permissoes (usuario_id, permissao, permitido)
                    VALUES (%s, %s, %s)
                    """,
                    permission_rows,
                )
            connection.commit()
            return saved_id
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def delete_system_user(self, environment: str, user_id: int) -> None:
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "DELETE FROM sistema_usuarios WHERE id = %s AND ambiente = %s",
                (user_id, environment),
            )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def authenticate_system_user(self, environment: str, login: str, password: str) -> dict[str, object] | None:
        normalized_login = str(login or "").strip().lower()
        if not normalized_login or not password:
            return None
        connection = self.get_connection()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT id, nome, login, senha_hash, senha_salt, ativo
                FROM sistema_usuarios
                WHERE ambiente = %s
                  AND login = %s
                LIMIT 1
                """,
                (environment, normalized_login),
            )
            user = cursor.fetchone()
            if not user or not int(user.get("ativo") or 0):
                return None
            salt = str(user.get("senha_salt") or "")
            expected_hash = str(user.get("senha_hash") or "")
            _salt, password_hash = self.hash_password(str(password), salt)
            if not secrets.compare_digest(expected_hash, password_hash):
                return None
            cursor.execute(
                """
                SELECT permissao
                FROM sistema_usuario_permissoes
                WHERE usuario_id = %s
                  AND permitido = 1
                """,
                (user["id"],),
            )
            user["permissoes"] = {str(row["permissao"]) for row in cursor.fetchall()}
            user.pop("senha_hash", None)
            user.pop("senha_salt", None)
            return dict(user)
        finally:
            connection.close()
