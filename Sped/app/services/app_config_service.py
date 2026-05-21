from __future__ import annotations

import json
from pathlib import Path

from app.config import APP_DEFAULT_CONFIG


def load_app_config(config_path: Path | None) -> dict[str, str]:
    if config_path is None:
        return dict(APP_DEFAULT_CONFIG)
    if not config_path.exists():
        save_app_config(config_path, APP_DEFAULT_CONFIG)
        return dict(APP_DEFAULT_CONFIG)
    try:
        loaded = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        loaded = {}
    config = dict(APP_DEFAULT_CONFIG)
    if isinstance(loaded, dict):
        for key, default_value in APP_DEFAULT_CONFIG.items():
            config[key] = str(loaded.get(key, default_value) or default_value)
    return config


def save_app_config(config_path: Path, config: dict[str, str]) -> None:
    payload = {
        key: str(config.get(key, APP_DEFAULT_CONFIG[key]) or APP_DEFAULT_CONFIG[key])
        for key in APP_DEFAULT_CONFIG
    }
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")
