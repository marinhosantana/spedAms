from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable

_ProgressCallback = Callable[[int, int, str], None]

_MYSQLDUMP_SEARCH_PATHS = [
    r"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
    r"C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqldump.exe",
    r"C:\Program Files\MySQL\MySQL Server 5.7\bin\mysqldump.exe",
    r"C:\xampp\mysql\bin\mysqldump.exe",
    r"C:\wamp64\bin\mysql\mysql8.0.31\bin\mysqldump.exe",
]


def _find_mysqldump() -> str | None:
    found = shutil.which("mysqldump")
    if found:
        return found
    for candidate in _MYSQLDUMP_SEARCH_PATHS:
        if Path(candidate).exists():
            return candidate
    return None


def _escape_sql_value(value: object) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, bytes):
        hex_str = value.hex()
        return f"0x{hex_str}"
    text = str(value)
    text = text.replace("\\", "\\\\")
    text = text.replace("'", "\\'")
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\x00", "\\0")
    return f"'{text}'"


def _python_dump(
    connection_config: dict[str, object],
    dest_path: Path,
    progress: _ProgressCallback | None,
) -> None:
    try:
        import mysql.connector
    except ImportError as exc:
        raise RuntimeError("mysql-connector-python nao esta instalado.") from exc

    conn = mysql.connector.connect(**connection_config)
    try:
        cursor = conn.cursor()
        cursor.execute("SHOW TABLES")
        tables: list[str] = [row[0] for row in cursor.fetchall()]
        total_steps = len(tables) * 2 + 1

        with dest_path.open("w", encoding="utf-8") as f:
            database = str(connection_config.get("database", ""))
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"-- Backup gerado pelo Revisor SPED em {now}\n")
            f.write(f"-- Banco de dados: {database}\n")
            f.write(f"-- Python {sys.version.split()[0]}\n\n")
            f.write("SET FOREIGN_KEY_CHECKS=0;\n\n")

            step = 0
            for table in tables:
                step += 1
                if progress:
                    progress(step, total_steps, f"Exportando estrutura: {table}")

                cursor.execute(f"SHOW CREATE TABLE `{table}`")
                row = cursor.fetchone()
                create_stmt = str(row[1]) if row else ""
                f.write(f"-- Tabela: {table}\n")
                f.write(f"DROP TABLE IF EXISTS `{table}`;\n")
                f.write(create_stmt + ";\n\n")

                step += 1
                if progress:
                    progress(step, total_steps, f"Exportando dados: {table}")

                cursor.execute(f"SELECT * FROM `{table}`")
                rows = cursor.fetchall()
                if rows:
                    col_cursor = conn.cursor()
                    col_cursor.execute(f"SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=%s AND TABLE_NAME=%s ORDER BY ORDINAL_POSITION", (database, table))
                    columns = [r[0] for r in col_cursor.fetchall()]
                    col_cursor.close()
                    cols_sql = ", ".join(f"`{c}`" for c in columns)
                    chunk_size = 500
                    for i in range(0, len(rows), chunk_size):
                        chunk = rows[i : i + chunk_size]
                        values_list = []
                        for data_row in chunk:
                            vals = ", ".join(_escape_sql_value(v) for v in data_row)
                            values_list.append(f"({vals})")
                        f.write(f"INSERT INTO `{table}` ({cols_sql}) VALUES\n")
                        f.write(",\n".join(values_list))
                        f.write(";\n")
                f.write("\n")

            f.write("SET FOREIGN_KEY_CHECKS=1;\n")

        step += 1
        if progress:
            progress(total_steps, total_steps, "Backup concluido.")
    finally:
        conn.close()


def backup_database(
    connection_config: dict[str, object],
    dest_path: Path,
    progress: _ProgressCallback | None = None,
) -> None:
    """Dump the full database to dest_path (.sql).

    Tries mysqldump first; falls back to a Python-based export if not found.
    connection_config must have: host, port, user, password, database.
    Raises RuntimeError on failure.
    """
    mysqldump = _find_mysqldump()
    if mysqldump:
        _mysqldump_dump(connection_config, dest_path, mysqldump, progress)
    else:
        _python_dump(connection_config, dest_path, progress)


def _mysqldump_dump(
    connection_config: dict[str, object],
    dest_path: Path,
    mysqldump_exe: str,
    progress: _ProgressCallback | None,
) -> None:
    host = str(connection_config.get("host", "127.0.0.1"))
    port = str(connection_config.get("port", "3306"))
    user = str(connection_config.get("user", "root"))
    password = str(connection_config.get("password", ""))
    database = str(connection_config.get("database", ""))

    if progress:
        progress(0, 1, "Executando mysqldump...")

    cmd = [
        mysqldump_exe,
        f"--host={host}",
        f"--port={port}",
        f"--user={user}",
        f"--password={password}",
        "--single-transaction",
        "--routines",
        "--triggers",
        "--add-drop-table",
        "--default-character-set=utf8mb4",
        database,
    ]

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with dest_path.open("w", encoding="utf-8") as out_file:
            result = subprocess.run(
                cmd,
                stdout=out_file,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
        if result.returncode != 0:
            stderr = result.stderr.strip()
            # mysqldump emits a password warning on stderr but still succeeds; ignore it
            real_errors = [line for line in stderr.splitlines() if "password" not in line.lower() and line.strip()]
            if real_errors:
                raise RuntimeError(f"mysqldump falhou:\n{chr(10).join(real_errors)}")
    except FileNotFoundError as exc:
        raise RuntimeError(f"mysqldump nao encontrado em: {mysqldump_exe}") from exc

    if progress:
        progress(1, 1, "Backup concluido via mysqldump.")


def suggest_backup_filename(database_name: str) -> str:
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = database_name.replace(" ", "_")
    return f"backup_{safe_name}_{now}.sql"
