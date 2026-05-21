from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Iterable

from app.parsers.sped_parser import parse_decimal
from app.services.tax_rules import format_decimal_sped


def format_audit_paths(paths: Iterable[Path] | Iterable[str]) -> str:
    formatted = [str(path) for path in paths if str(path).strip()]
    return "; ".join(formatted) if formatted else "Nao informado"


def audit_decimal_from_row(row: dict[str, object], *keys: str) -> Decimal:
    for key in keys:
        value = row.get(key)
        if isinstance(value, Decimal):
            return value
        if value not in (None, ""):
            try:
                return parse_decimal(str(value))
            except Exception:
                continue
    return Decimal("0.00")


def summarize_audit_detail_rows(rows: list[dict[str, object]]) -> dict[str, Decimal | int]:
    return {
        "linhas": len(rows),
        "valor_operacao": sum(
            (audit_decimal_from_row(row, "sale_value", "total_operation_value", "operation_value") for row in rows),
            Decimal("0.00"),
        ).quantize(Decimal("0.01")),
        "base_icms": sum((audit_decimal_from_row(row, "base_icms") for row in rows), Decimal("0.00")).quantize(
            Decimal("0.01")
        ),
        "valor_icms": sum((audit_decimal_from_row(row, "icms_value") for row in rows), Decimal("0.00")).quantize(
            Decimal("0.01")
        ),
    }


def format_audit_summary(summary: dict[str, Decimal | int]) -> str:
    return (
        f"linhas={summary.get('linhas', 0)}; "
        f"valor_operacao={format_decimal_sped(Decimal(summary.get('valor_operacao', Decimal('0.00'))))}; "
        f"base_icms={format_decimal_sped(Decimal(summary.get('base_icms', Decimal('0.00'))))}; "
        f"valor_icms={format_decimal_sped(Decimal(summary.get('valor_icms', Decimal('0.00'))))}"
    )
