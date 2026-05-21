from __future__ import annotations

import os
import sys
from pathlib import Path


def get_application_base_dir(source_file: str) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(source_file).resolve().parent


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
