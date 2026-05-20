from __future__ import annotations

import os
from pathlib import Path


XML_SELECTION_COLLAPSE_THRESHOLD = 250


def parse_selected_paths(raw_value: str) -> list[Path]:
    selected_paths: list[Path] = []
    for chunk in str(raw_value or "").split(";"):
        text = chunk.strip().strip('"')
        if not text:
            continue
        path = Path(text)
        if path not in selected_paths:
            selected_paths.append(path)
    return selected_paths


def format_selected_paths(paths: list[Path]) -> str:
    return "; ".join(str(path) for path in paths)


def collapse_xml_selection_paths(selected_files: list[str]) -> tuple[list[Path], bool]:
    normalized_paths: list[Path] = []
    for selected in selected_files:
        selected_path = Path(selected)
        if selected_path not in normalized_paths:
            normalized_paths.append(selected_path)
    if len(normalized_paths) < XML_SELECTION_COLLAPSE_THRESHOLD:
        return normalized_paths, False

    parent_folders = {path.parent.resolve() for path in normalized_paths}
    if len(parent_folders) == 1:
        return [next(iter(parent_folders))], True
    return normalized_paths, False


def get_xml_worker_count(total_files: int) -> int:
    if total_files <= 1:
        return 1
    cpu_total = os.cpu_count() or 1
    return max(1, min(8, cpu_total, total_files))
