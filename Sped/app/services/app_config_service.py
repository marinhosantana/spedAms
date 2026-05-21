from __future__ import annotations

import json
from pathlib import Path

from app.config import APP_DEFAULT_CONFIG


def load_app_config(config_path: Path | None, default_config: dict[str, str] | None = None) -> dict[str, str]:
    defaults = dict(default_config or APP_DEFAULT_CONFIG)
    if config_path is None:
        return defaults
    if not config_path.exists():
        save_app_config(config_path, defaults)
        return defaults
    try:
        loaded = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        loaded = {}
    config = dict(defaults)
    if isinstance(loaded, dict):
        for key, default_value in defaults.items():
            config[key] = str(loaded.get(key, default_value) or default_value)
    return config


def save_app_config(config_path: Path, config: dict[str, str], default_config: dict[str, str] | None = None) -> None:
    defaults = dict(default_config or APP_DEFAULT_CONFIG)
    payload = {
        key: str(config.get(key, defaults[key]) or defaults[key])
        for key in defaults
    }
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")
