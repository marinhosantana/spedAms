from __future__ import annotations

import os
import sys
from pathlib import Path


VALID_APPLICATION_ENVIRONMENTS = {"dev", "prod"}
DEFAULT_SOURCE_APPLICATION_ENVIRONMENT = "dev"
DEFAULT_FROZEN_APPLICATION_ENVIRONMENT = "prod"
ENVIRONMENT_ALIASES = {
    "dev": "dev",
    "desenvolvimento": "dev",
    "development": "dev",
    "prod": "prod",
    "producao": "prod",
    "produção": "prod",
    "production": "prod",
}


def _default_application_environment() -> str:
    if getattr(sys, "frozen", False):
        return DEFAULT_FROZEN_APPLICATION_ENVIRONMENT
    return DEFAULT_SOURCE_APPLICATION_ENVIRONMENT


def get_application_environment() -> str:
    raw_environment = os.environ.get("SPED_ENV", "").strip().lower()
    if raw_environment:
        return ENVIRONMENT_ALIASES.get(raw_environment, _default_application_environment())
    return _default_application_environment()


def get_application_base_dir(source_file: str) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(source_file).resolve().parent


def get_project_root_dir(base_dir: Path) -> Path:
    if getattr(sys, "frozen", False):
        return base_dir
    return base_dir.parents[1]


def get_environment_config_path(base_dir: Path, filename_prefix: str, environment: str | None = None) -> Path:
    selected_environment = environment or get_application_environment()
    if selected_environment not in VALID_APPLICATION_ENVIRONMENTS:
        selected_environment = get_application_environment()
    return base_dir / f"{filename_prefix}.{selected_environment}.json"


def get_audit_log_path(base_dir: Path) -> Path:
    log_dir = base_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "auditoria_sped.log"


def get_runtime_rule_history_path(base_dir: Path) -> Path:
    if getattr(sys, "frozen", False):
        local_appdata = Path(os.environ.get("LOCALAPPDATA", base_dir))
        data_dir = local_appdata / "Revisor de SPED - DZ Consultoria"
        data_dir.mkdir(parents=True, exist_ok=True)
        return data_dir / "runtime_rule_history.json"
    return base_dir / "runtime_rule_history.json"


def get_sped_archive_storage_dir(project_root_dir: Path) -> Path:
    storage_dir = project_root_dir / "storage" / "original_speds"
    storage_dir.mkdir(parents=True, exist_ok=True)
    return storage_dir
