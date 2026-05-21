from __future__ import annotations

import hashlib
import shutil
from dataclasses import dataclass
from pathlib import Path

from app.parsers.sped_parser import get_field, normalize_sped_line


@dataclass(frozen=True)
class SpedArchiveMetadata:
    source_path: Path
    archived_path: Path
    file_name: str
    file_hash_sha256: str
    file_size: int
    company_name: str
    company_tax_id: str
    period_start: str
    period_end: str
    sped_type: str
    default_profile_name: str


def calculate_file_sha256(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as source_file:
        for chunk in iter(lambda: source_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_tax_id(value: object) -> str:
    return "".join(char for char in str(value or "") if char.isdigit())


def normalize_company_name(value: object) -> str:
    return " ".join(str(value or "").strip().split())


def looks_like_company_name(value: str) -> bool:
    text = normalize_company_name(value)
    if not text:
        return False
    digits = normalize_tax_id(text)
    if digits == text and len(digits) in {8, 11, 14}:
        return False
    return any(char.isalpha() for char in text)


def format_sped_date(value: str) -> str:
    text = normalize_tax_id(value)
    if len(text) != 8:
        return ""
    return f"{text[4:8]}-{text[2:4]}-{text[0:2]}"


def infer_sped_metadata(file_path: Path) -> dict[str, str]:
    company_name = ""
    company_tax_id = ""
    period_start = ""
    period_end = ""
    sped_type = "fiscal"
    seen_contrib_register = False

    with file_path.open("r", encoding="latin-1") as sped_file:
        for line_index, raw_line in enumerate(sped_file, start=1):
            if line_index > 5000 and company_tax_id and period_start and period_end:
                break
            if not raw_line.startswith("|"):
                continue
            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)
            if register == "0000":
                text_candidates = [
                    normalize_company_name(get_field(fields, 5)),
                    normalize_company_name(get_field(fields, 6)),
                    normalize_company_name(get_field(fields, 8)),
                ]
                for candidate in text_candidates:
                    if looks_like_company_name(candidate):
                        company_name = candidate
                        break
                tax_candidates = [
                    normalize_tax_id(get_field(fields, 7)),
                    normalize_tax_id(get_field(fields, 8)),
                    normalize_tax_id(get_field(fields, 9)),
                ]
                for candidate in tax_candidates:
                    if len(candidate) in {11, 14}:
                        company_tax_id = candidate
                        break
                dates = [format_sped_date(field) for field in fields if format_sped_date(field)]
                if dates:
                    period_start = dates[0]
                if len(dates) > 1:
                    period_end = dates[1]
                continue
            if register in {"M100", "M200", "M500", "M600", "F100"}:
                seen_contrib_register = True

    if seen_contrib_register:
        sped_type = "contribuicoes"

    return {
        "company_name": company_name,
        "company_tax_id": company_tax_id,
        "period_start": period_start,
        "period_end": period_end,
        "sped_type": sped_type,
    }


def build_default_profile_name(metadata: dict[str, str], file_path: Path) -> str:
    company = metadata.get("company_name", "") or metadata.get("company_tax_id", "") or "empresa_nao_identificada"
    start = metadata.get("period_start", "")
    end = metadata.get("period_end", "")
    sped_type = metadata.get("sped_type", "sped")
    if start and end:
        return f"{company} - {sped_type} - {start} a {end}"
    return f"{company} - {sped_type} - {file_path.stem}"


def archive_original_sped_file(source_path: Path, storage_root: Path, environment: str) -> SpedArchiveMetadata:
    resolved_source = source_path.resolve()
    if not resolved_source.exists() or not resolved_source.is_file():
        raise FileNotFoundError(f"Arquivo SPED nao encontrado: {resolved_source}")

    file_hash = calculate_file_sha256(resolved_source)
    metadata = infer_sped_metadata(resolved_source)
    period_folder = (metadata.get("period_start", "")[:7] or "sem_periodo").replace("-", "")
    archived_dir = storage_root / environment / period_folder / file_hash[:2]
    archived_dir.mkdir(parents=True, exist_ok=True)
    archived_path = archived_dir / f"{file_hash}{resolved_source.suffix.lower() or '.txt'}"
    if not archived_path.exists():
        shutil.copy2(resolved_source, archived_path)

    return SpedArchiveMetadata(
        source_path=resolved_source,
        archived_path=archived_path.resolve(),
        file_name=resolved_source.name,
        file_hash_sha256=file_hash,
        file_size=resolved_source.stat().st_size,
        company_name=metadata.get("company_name", ""),
        company_tax_id=metadata.get("company_tax_id", ""),
        period_start=metadata.get("period_start", ""),
        period_end=metadata.get("period_end", ""),
        sped_type=metadata.get("sped_type", ""),
        default_profile_name=build_default_profile_name(metadata, resolved_source),
    )
