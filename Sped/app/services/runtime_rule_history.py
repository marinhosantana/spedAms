from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from app.services.runtime_rules import runtime_rule_summary


def load_runtime_rule_history(history_path: Path) -> list[dict[str, object]]:
    if not history_path.exists():
        return []
    try:
        loaded = json.loads(history_path.read_text(encoding="utf-8"))
    except Exception:
        return []
    return loaded if isinstance(loaded, list) else []


def save_runtime_rule_history(history_path: Path, history: list[dict[str, object]]) -> None:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    history_path.write_text(
        json.dumps(history, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def remember_runtime_rules(
    history: list[dict[str, object]],
    rule_lines: list[str],
    source: str,
) -> list[dict[str, object]]:
    if not rule_lines:
        return history
    now = dt.datetime.now().replace(microsecond=0).isoformat()
    history_index = {
        str(item.get("rule_line", "")).strip(): item
        for item in history
        if str(item.get("rule_line", "")).strip()
    }
    for rule_line in rule_lines:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            continue
        item = history_index.get(normalized_line)
        if item is None:
            item = {
                "rule_line": normalized_line,
                "summary": runtime_rule_summary(normalized_line),
                "use_count": 0,
                "created_at": now,
                "last_used_at": now,
                "last_source": source,
            }
            history.append(item)
            history_index[normalized_line] = item
        item["summary"] = runtime_rule_summary(normalized_line)
        item["use_count"] = int(item.get("use_count", 0)) + 1
        item["last_used_at"] = now
        item["last_source"] = source
    return history


def remove_runtime_rule(history: list[dict[str, object]], rule_line: str) -> list[dict[str, object]]:
    normalized_line = str(rule_line or "").strip()
    return [
        item
        for item in history
        if str(item.get("rule_line", "")).strip() != normalized_line
    ]
