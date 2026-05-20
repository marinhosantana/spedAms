from __future__ import annotations

from decimal import Decimal

from app.parsers.sped_parser import normalize_document_key
from app.services.analysis_utils import period_label_sort_key
from app.services.runtime_rules import apply_sped_icms_consistency_rules
from app.services.tax_rules import (
    compute_display_icms_rate,
    normalize_cst_icms_for_sped,
    normalize_operation_type,
    normalize_text,
)


def format_grouped_line_numbers(grouped_rows: list[dict[str, object]], max_items: int = 8) -> str:
    line_numbers = [str(row.get("line_number", "")).strip() for row in grouped_rows if str(row.get("line_number", "")).strip()]
    if not line_numbers:
        return ""
    preview = ", ".join(line_numbers[:max_items])
    if len(line_numbers) > max_items:
        preview += ", ..."
    return preview

def build_credit_diagnostic_product_sheet_payload(
    detail_rows: list[dict[str, object]],
) -> tuple[list[list[object]], list[dict[str, object]]]:
    product_rows: list[list[object]] = []
    footer_rows: list[dict[str, object]] = []

    for detail_row in detail_rows:
        period = str(detail_row.get("period", ""))
        reason = str(detail_row.get("reason", ""))
        source_register = str(detail_row.get("source_register", ""))
        cst_icms = str(detail_row.get("cst_icms", ""))
        cfop = str(detail_row.get("cfop", ""))
        nominal_rate = Decimal(detail_row.get("icms_rate", Decimal("0")))

        detail_operation_total = Decimal(detail_row.get("total_operation_value", Decimal("0")))
        detail_base_total = Decimal(detail_row.get("base_icms", Decimal("0")))
        detail_icms_total = Decimal(detail_row.get("icms_value", Decimal("0")))
        detail_effective_rate = compute_display_icms_rate(
            nominal_rate,
            detail_operation_total,
            detail_base_total,
            detail_icms_total,
        )

        product_operation_total = Decimal("0")
        product_base_total = Decimal("0")
        product_icms_total = Decimal("0")

        for product_row in detail_row.get("product_rows", []):
            product_sale = Decimal(product_row.get("sale_value", Decimal("0")))
            product_base = Decimal(product_row.get("base_icms", Decimal("0")))
            product_icms = Decimal(product_row.get("icms_value", Decimal("0")))
            product_ratio = (
                (product_base * Decimal("100") / product_sale).quantize(Decimal("0.01"))
                if product_sale > 0
                else Decimal("0.00")
            )
            product_effective_rate = compute_display_icms_rate(
                Decimal(product_row.get("icms_rate", Decimal("0"))),
                product_sale,
                product_base,
                product_icms,
            )
            product_rows.append(
                [
                    product_row.get("period", period),
                    reason,
                    source_register,
                    cst_icms,
                    cfop,
                    product_row.get("code", ""),
                    product_row.get("description", ""),
                    product_row.get("ncm", ""),
                    Decimal(product_row.get("icms_rate", nominal_rate)),
                    product_effective_rate,
                    product_row.get("row_count", 0),
                    product_sale,
                    product_base,
                    product_ratio,
                    product_icms,
                ]
            )
            product_operation_total += product_sale
            product_base_total += product_base
            product_icms_total += product_icms

        operation_diff = detail_operation_total - product_operation_total
        base_diff = detail_base_total - product_base_total
        icms_diff = detail_icms_total - product_icms_total
        if not any(value != 0 for value in (operation_diff, base_diff, icms_diff)):
            continue

        grouped_rows = list(detail_row.get("grouped_rows", []))
        line_numbers = format_grouped_line_numbers(grouped_rows)
        adjustment_label = "SEM PRODUTO IDENTIFICADO" if not detail_row.get("product_rows") else "AJUSTE RESUMO X ITENS"
        product_rows.append(
            [
                period,
                reason,
                source_register,
                cst_icms,
                cfop,
                "",
                adjustment_label,
                line_numbers,
                nominal_rate,
                detail_effective_rate,
                0,
                operation_diff,
                base_diff,
                "",
                icms_diff,
            ]
        )

    return product_rows, footer_rows

def classify_credit_base_reason(
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)[-2:]
    if total_operation_value <= 0 and base_icms <= 0 and icms_value <= 0:
        return "Sem valores fiscais"
    if base_icms <= 0 and icms_value > 0:
        return "Credito sem base fiscal"
    if normalized_cst in {"40", "41", "50"}:
        return "CST sem direito a credito"
    if normalized_cst in {"60", "61"}:
        return "Substituicao tributaria / ICMS ST"
    if normalized_cst in {"90"} and base_icms <= 0:
        return "Outras entradas sem base de credito"
    if total_operation_value > 0 and base_icms <= 0:
        return "Operacao sem base de credito"
    if total_operation_value > 0 and base_icms < total_operation_value:
        return "Base reduzida em relacao ao valor da operacao"
    if source_register in {"C590", "D190", "D590", "D730"}:
        return "Servico / energia / comunicacao"
    return "Base integral"

def classify_debit_base_reason(
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)[-2:]
    if total_operation_value <= 0 and base_icms <= 0 and icms_value <= 0:
        return "Sem valores fiscais"
    if base_icms <= 0 and icms_value > 0:
        return "Debito sem base fiscal"
    if normalized_cst in {"40", "41", "50"}:
        return "CST sem debito do imposto"
    if normalized_cst in {"60", "61"}:
        return "Substituicao tributaria / ICMS ST"
    if normalized_cst in {"90"} and base_icms <= 0:
        return "Outras saidas sem base de debito"
    if total_operation_value > 0 and base_icms <= 0:
        return "Operacao sem base de debito"
    if total_operation_value > 0 and icms_value <= 0:
        return "Operacao sem debito"
    if total_operation_value > 0 and base_icms < total_operation_value:
        return "Base reduzida em relacao ao valor da operacao"
    if source_register in {"C590", "D190", "D590", "D730"}:
        return "Servico / energia / comunicacao"
    return "Base integral"

def classify_icms_diagnostic_reason(
    operation_type: str,
    source_register: str,
    cst_icms: str,
    total_operation_value: Decimal,
    base_icms: Decimal,
    icms_value: Decimal,
) -> str:
    if normalize_operation_type(operation_type) == "saida":
        return classify_debit_base_reason(source_register, cst_icms, total_operation_value, base_icms, icms_value)
    return classify_credit_base_reason(source_register, cst_icms, total_operation_value, base_icms, icms_value)

def build_credit_diagnostic_datasets(
    summary_register_rows: list[dict[str, object]],
    operation_type: str,
    cst_filter_values: set[str],
    cfop_filter_values: set[str],
    detailed_sales: list[dict[str, object]] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, Decimal]]:
    normalized_operation = normalize_operation_type(operation_type)
    normalized_csts = {normalize_cst_icms_for_sped(str(value)) for value in cst_filter_values if str(value).strip()}
    normalized_cfops = {str(value).strip() for value in cfop_filter_values if str(value).strip()}

    detail_groups: dict[tuple[str, str, str, str, str, Decimal], dict[str, object]] = {}
    summary_groups: dict[tuple[str, str], dict[str, object]] = {}
    totals = {
        "total_operation_value": Decimal("0"),
        "base_icms": Decimal("0"),
        "icms_value": Decimal("0"),
        "base_gap": Decimal("0"),
    }

    for row in summary_register_rows:
        if normalize_operation_type(str(row.get("operation_type", ""))) != normalized_operation:
            continue
        cst_value = normalize_cst_icms_for_sped(str(row.get("cst_icms", "")))
        cfop_value = str(row.get("cfop", "")).strip()
        if normalized_csts and cst_value not in normalized_csts:
            continue
        if normalized_cfops and cfop_value not in normalized_cfops:
            continue

        source_register = str(row.get("source_register", "")).strip()
        period = str(row.get("period", "")).strip()
        icms_rate = Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
        total_operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
        base_icms = Decimal(row.get("base_icms", Decimal("0")))
        icms_value = Decimal(row.get("icms_value", Decimal("0")))
        reason = classify_icms_diagnostic_reason(normalized_operation, source_register, cst_value, total_operation_value, base_icms, icms_value)
        base_gap = max(Decimal("0"), total_operation_value - base_icms)
        base_ratio = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )

        totals["total_operation_value"] += total_operation_value
        totals["base_icms"] += base_icms
        totals["icms_value"] += icms_value
        totals["base_gap"] += base_gap

        detail_key = (period, reason, source_register, cst_value, cfop_value, icms_rate)
        detail_bucket = detail_groups.setdefault(
            detail_key,
            {
                "operation_type": normalized_operation,
                "period": period,
                "reason": reason,
                "source_register": source_register,
                "cst_icms": cst_value,
                "cfop": cfop_value,
                "icms_rate": icms_rate,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_gap": Decimal("0"),
                "row_count": 0,
                "grouped_rows": [],
            },
        )
        detail_bucket["total_operation_value"] += total_operation_value
        detail_bucket["base_icms"] += base_icms
        detail_bucket["icms_value"] += icms_value
        detail_bucket["base_gap"] += base_gap
        detail_bucket["row_count"] += 1
        grouped_rows = detail_bucket.get("grouped_rows")
        if isinstance(grouped_rows, list):
            grouped_rows.append(dict(row))

        summary_bucket = summary_groups.setdefault(
            (period, reason),
            {
                "period": period,
                "reason": reason,
                "total_operation_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
                "base_gap": Decimal("0"),
                "row_count": 0,
            },
        )
        summary_bucket["total_operation_value"] += total_operation_value
        summary_bucket["base_icms"] += base_icms
        summary_bucket["icms_value"] += icms_value
        summary_bucket["base_gap"] += base_gap
        summary_bucket["row_count"] += 1

    detail_rows = sorted(
        detail_groups.values(),
        key=lambda row: (
            Decimal(row["base_gap"]),
            Decimal(row["total_operation_value"]),
        ),
        reverse=True,
    )
    for row in detail_rows:
        total_operation_value = Decimal(row["total_operation_value"])
        base_icms = Decimal(row["base_icms"])
        row["base_ratio"] = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        row["product_rows"] = (
            build_credit_diagnostic_product_rows(
                detailed_sales,
                normalized_operation,
                str(row["period"]),
                str(row["source_register"]),
                str(row["reason"]),
                str(row["cst_icms"]),
                str(row["cfop"]),
                Decimal(row["icms_rate"]),
            )
            if detailed_sales
            else []
        )
        row["launch_details"] = [
            dict(detail)
            for product_row in row["product_rows"]
            for detail in product_row.get("launch_details", [])
        ]
        row["document_count"] = len(
            {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in row["launch_details"]
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
        )

    summary_rows = sorted(
        summary_groups.values(),
        key=lambda row: Decimal(row["base_gap"]),
        reverse=True,
    )
    for row in summary_rows:
        total_operation_value = Decimal(row["total_operation_value"])
        base_icms = Decimal(row["base_icms"])
        row["base_ratio"] = (
            (base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )

    return summary_rows, detail_rows, totals

def build_credit_diagnostic_product_rows(
    detailed_sales: list[dict[str, object]],
    operation_type: str,
    period: str,
    source_register: str,
    reason: str,
    cst_icms: str,
    cfop: str,
    icms_rate: Decimal,
) -> list[dict[str, object]]:
    normalized_operation = normalize_operation_type(operation_type)
    normalized_cst = normalize_cst_icms_for_sped(cst_icms)
    normalized_cfop = str(cfop or "").strip()
    rate_key = Decimal(icms_rate).quantize(Decimal("0.01"))
    candidate_items: list[tuple[dict[str, object], str, Decimal, Decimal, Decimal]] = []
    fallback_items: list[tuple[dict[str, object], str, Decimal, Decimal, Decimal]] = []

    for source_item in detailed_sales:
        if normalize_operation_type(str(source_item.get("operation_type", ""))) != normalized_operation:
            continue
        if str(source_item.get("period", "")).strip() != str(period or "").strip():
            continue
        item = apply_sped_icms_consistency_rules(source_item)
        item_cst = normalize_cst_icms_for_sped(str(item.get("cst_icms", "")))
        item_cfop = str(item.get("cfop", "")).strip()
        item_rate = (
            item.get("icms_rate", Decimal("0"))
            if isinstance(item.get("icms_rate", Decimal("0")), Decimal)
            else Decimal("0")
        ).quantize(Decimal("0.01"))
        if item_cst != normalized_cst or item_cfop != normalized_cfop or item_rate != rate_key:
            continue

        sale_value = item.get("sale_value", Decimal("0"))
        base_icms = item.get("base_icms", Decimal("0"))
        icms_value = item.get("icms_value", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_icms = base_icms if isinstance(base_icms, Decimal) else Decimal("0")
        icms_value = icms_value if isinstance(icms_value, Decimal) else Decimal("0")
        fallback_items.append((item, item_cst, sale_value, base_icms, icms_value))

        item_reason = classify_icms_diagnostic_reason(normalized_operation, source_register, item_cst, sale_value, base_icms, icms_value)
        if item_reason == reason:
            candidate_items.append((item, item_cst, sale_value, base_icms, icms_value))

    selected_items = candidate_items if candidate_items else fallback_items
    grouped_products: dict[tuple[str, str, str], dict[str, object]] = {}
    for item, _item_cst, sale_value, base_icms, icms_value in selected_items:
        item_rate = (
            item.get("icms_rate", Decimal("0"))
            if isinstance(item.get("icms_rate", Decimal("0")), Decimal)
            else Decimal("0")
        ).quantize(Decimal("0.01"))
        product_key = (
            str(item.get("code", "")).strip(),
            str(item.get("description", "")).strip(),
            str(item.get("ncm", "")).strip(),
        )
        bucket = grouped_products.setdefault(
            product_key,
            {
                "code": product_key[0],
                "description": product_key[1],
                "ncm": product_key[2],
                "period": period,
                "icms_rate": item_rate,
                "row_count": 0,
                "document_keys": set(),
                "launch_details": [],
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )
        bucket["row_count"] += 1
        document_identity = normalize_document_key(str(item.get("document_key", ""))) or str(item.get("document_number", "")).strip()
        if document_identity:
            bucket["document_keys"].add(document_identity)
        bucket["launch_details"].append(dict(item))
        bucket["sale_value"] += sale_value
        bucket["base_icms"] += base_icms
        bucket["icms_value"] += icms_value

    return sorted(
        grouped_products.values(),
        key=lambda row: (Decimal(row["sale_value"]), str(row["code"]), str(row["description"])),
        reverse=True,
    )

def build_credit_diagnostic_period_comparison_dataset(
    summary_register_rows: list[dict[str, object]],
    detailed_sales: list[dict[str, object]],
    cst_filter_values: set[str],
    cfop_filter_values: set[str],
    operation_type: str = "Entrada",
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]], list[dict[str, object]]]:
    summary_rows, detail_rows, _totals = build_credit_diagnostic_datasets(
        summary_register_rows,
        operation_type,
        cst_filter_values,
        cfop_filter_values,
        detailed_sales,
    )
    periods = sorted(
        {
            str(row.get("period", "")).strip()
            for row in summary_rows
            if str(row.get("period", "")).strip()
        },
        key=period_label_sort_key,
    )
    grouped_reasons: dict[str, dict[str, object]] = {}
    for row in summary_rows:
        reason = str(row.get("reason", "")).strip()
        if not reason:
            continue
        bucket = grouped_reasons.setdefault(
            reason,
            {
                "reason": reason,
                "periods": {},
                "total_rows": 0,
                "total_operation_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_base_gap": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        period = str(row.get("period", "")).strip()
        bucket["periods"][period] = {
            "row_count": int(row.get("row_count", 0) or 0),
            "operation_value": Decimal(row.get("total_operation_value", Decimal("0"))),
            "base_icms": Decimal(row.get("base_icms", Decimal("0"))),
            "base_gap": Decimal(row.get("base_gap", Decimal("0"))),
            "base_ratio": Decimal(row.get("base_ratio", Decimal("0"))),
            "icms_value": Decimal(row.get("icms_value", Decimal("0"))),
        }
        bucket["total_rows"] += int(row.get("row_count", 0) or 0)
        bucket["total_operation_value"] += Decimal(row.get("total_operation_value", Decimal("0")))
        bucket["total_base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        bucket["total_base_gap"] += Decimal(row.get("base_gap", Decimal("0")))
        bucket["total_icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

    export_headers = ["Motivo"]
    for period in periods:
        export_headers.extend(
            [
                f"{period} Linhas",
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Perda de Base",
                f"{period} % Base/Oper",
                f"{period} Valor ICMS",
            ]
        )
    export_headers.extend(
        [
            "Total Linhas",
            "Total Valor Operacao",
            "Total Base ICMS",
            "Total Perda de Base",
            "Total % Base/Oper",
            "Total Valor ICMS",
        ]
    )

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for reason in sorted(grouped_reasons.keys(), key=normalize_text):
        reason_bucket = grouped_reasons[reason]
        values: list[object] = [reason]
        for period in periods:
            period_data = reason_bucket["periods"].get(period)
            if period_data is None:
                values.extend([0, Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    int(period_data["row_count"]),
                    Decimal(period_data["operation_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["base_gap"]),
                    Decimal(period_data["base_ratio"]),
                    Decimal(period_data["icms_value"]),
                ]
            )
        total_operation_value = Decimal(reason_bucket["total_operation_value"])
        total_base_icms = Decimal(reason_bucket["total_base_icms"])
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        values.extend(
            [
                int(reason_bucket["total_rows"]),
                total_operation_value,
                total_base_icms,
                Decimal(reason_bucket["total_base_gap"]),
                total_base_ratio,
                Decimal(reason_bucket["total_icms_value"]),
            ]
        )
        display_rows.append({"reason": reason, "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows, detail_rows

def merge_credit_diagnostic_detail_rows(detail_rows: list[dict[str, object]], reason: str, period: str) -> dict[str, object] | None:
    matching_rows = [
        row
        for row in detail_rows
        if str(row.get("reason", "")).strip() == str(reason).strip()
        and str(row.get("period", "")).strip() == str(period).strip()
    ]
    if not matching_rows:
        return None
    if len(matching_rows) == 1:
        return dict(matching_rows[0])

    source_registers = sorted({str(row.get("source_register", "")).strip() for row in matching_rows if str(row.get("source_register", "")).strip()})
    csts = sorted({str(row.get("cst_icms", "")).strip() for row in matching_rows if str(row.get("cst_icms", "")).strip()})
    cfops = sorted({str(row.get("cfop", "")).strip() for row in matching_rows if str(row.get("cfop", "")).strip()})
    grouped_rows: list[dict[str, object]] = []
    launch_details: list[dict[str, object]] = []
    product_buckets: dict[tuple[str, str, str, str, Decimal], dict[str, object]] = {}

    for row in matching_rows:
        grouped_rows.extend(list(row.get("grouped_rows", [])))
        launch_details.extend(list(row.get("launch_details", [])))
        for product in row.get("product_rows", []):
            product_key = (
                str(product.get("period", "")).strip(),
                str(product.get("code", "")).strip(),
                str(product.get("description", "")).strip(),
                str(product.get("ncm", "")).strip(),
                Decimal(product.get("icms_rate", Decimal("0"))),
            )
            bucket = product_buckets.setdefault(
                product_key,
                {
                    "period": product_key[0],
                    "code": product_key[1],
                    "description": product_key[2],
                    "ncm": product_key[3],
                    "icms_rate": product_key[4],
                    "row_count": 0,
                    "document_keys": set(),
                    "launch_details": [],
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            bucket["row_count"] += int(product.get("row_count", 0) or 0)
            bucket["document_keys"].update(set(product.get("document_keys", set())))
            bucket["launch_details"].extend(list(product.get("launch_details", [])))
            bucket["sale_value"] += Decimal(product.get("sale_value", Decimal("0")))
            bucket["base_icms"] += Decimal(product.get("base_icms", Decimal("0")))
            bucket["icms_value"] += Decimal(product.get("icms_value", Decimal("0")))

    total_operation_value = sum((Decimal(row.get("total_operation_value", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_icms = sum((Decimal(row.get("base_icms", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_gap = sum((Decimal(row.get("base_gap", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_icms_value = sum((Decimal(row.get("icms_value", Decimal("0"))) for row in matching_rows), Decimal("0"))
    total_base_ratio = (
        (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
        if total_operation_value > 0
        else Decimal("0.00")
    )
    return {
        "operation_type": str(matching_rows[0].get("operation_type", "")).strip(),
        "period": period,
        "reason": reason,
        "source_register": " | ".join(source_registers) if source_registers else "",
        "cst_icms": " | ".join(csts) if csts else "",
        "cfop": " | ".join(cfops) if cfops else "",
        "icms_rate": next(
            (
                Decimal(row.get("icms_rate", Decimal("0")))
                for row in matching_rows
                if Decimal(row.get("icms_rate", Decimal("0"))) > 0
            ),
            Decimal("0"),
        ),
        "row_count": sum((int(row.get("row_count", 0) or 0) for row in matching_rows), 0),
        "total_operation_value": total_operation_value,
        "base_icms": total_base_icms,
        "base_gap": total_base_gap,
        "base_ratio": total_base_ratio,
        "icms_value": total_icms_value,
        "grouped_rows": grouped_rows,
        "launch_details": launch_details,
        "document_count": len(
            {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in launch_details
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
        ),
        "product_rows": sorted(
            product_buckets.values(),
            key=lambda row: (Decimal(row.get("sale_value", Decimal("0"))), str(row.get("code", ""))),
            reverse=True,
        ),
    }

def build_credit_diagnostic_period_detail_comparison_dataset(
    detail_rows: list[dict[str, object]],
    periods: list[str],
) -> tuple[list[str], list[dict[str, object]], list[list[object]]]:
    grouped_details: dict[tuple[str, str, str, str, Decimal], dict[str, object]] = {}

    for row in detail_rows:
        reason = str(row.get("reason", "")).strip()
        source_register = str(row.get("source_register", "")).strip()
        cst_icms = str(row.get("cst_icms", "")).strip()
        cfop = str(row.get("cfop", "")).strip()
        icms_rate = Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
        period = str(row.get("period", "")).strip()
        if not reason:
            continue
        detail_key = (reason, source_register, cst_icms, cfop, icms_rate)
        bucket = grouped_details.setdefault(
            detail_key,
            {
                "reason": reason,
                "source_register": source_register,
                "cst_icms": cst_icms,
                "cfop": cfop,
                "icms_rate": icms_rate,
                "periods": {},
                "total_rows": 0,
                "total_operation_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_base_gap": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
        base_icms = Decimal(row.get("base_icms", Decimal("0")))
        icms_value = Decimal(row.get("icms_value", Decimal("0")))
        effective_rate = compute_display_icms_rate(
            icms_rate,
            operation_value,
            base_icms,
            icms_value,
        )
        bucket["periods"][period] = {
            "row_count": int(row.get("row_count", 0) or 0),
            "operation_value": operation_value,
            "base_icms": base_icms,
            "base_gap": Decimal(row.get("base_gap", Decimal("0"))),
            "base_ratio": Decimal(row.get("base_ratio", Decimal("0"))),
            "icms_value": icms_value,
            "effective_rate": effective_rate,
        }
        bucket["total_rows"] += int(row.get("row_count", 0) or 0)
        bucket["total_operation_value"] += operation_value
        bucket["total_base_icms"] += base_icms
        bucket["total_base_gap"] += Decimal(row.get("base_gap", Decimal("0")))
        bucket["total_icms_value"] += icms_value

    export_headers = ["Motivo", "Registro", "CST", "CFOP", "Aliq"]
    for period in periods:
        export_headers.extend(
            [
                f"{period} Linhas",
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Perda de Base",
                f"{period} % Base/Oper",
                f"{period} Valor ICMS",
                f"{period} % Aliq Efetiva",
            ]
        )
    export_headers.extend(
        [
            "Total Linhas",
            "Total Valor Operacao",
            "Total Base ICMS",
            "Total Perda de Base",
            "Total % Base/Oper",
            "Total Valor ICMS",
            "Total % Aliq Efetiva",
        ]
    )

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for detail_key in sorted(
        grouped_details.keys(),
        key=lambda item: (
            normalize_text(item[0]),
            normalize_text(item[1]),
            normalize_text(item[2]),
            normalize_text(item[3]),
            item[4],
        ),
    ):
        bucket = grouped_details[detail_key]
        values: list[object] = [
            bucket["reason"],
            bucket["source_register"],
            bucket["cst_icms"],
            bucket["cfop"],
            Decimal(bucket["icms_rate"]),
        ]
        for period in periods:
            period_data = bucket["periods"].get(period)
            if period_data is None:
                values.extend([0, Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    int(period_data["row_count"]),
                    Decimal(period_data["operation_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["base_gap"]),
                    Decimal(period_data["base_ratio"]),
                    Decimal(period_data["icms_value"]),
                    Decimal(period_data["effective_rate"]),
                ]
            )
        total_operation_value = Decimal(bucket["total_operation_value"])
        total_base_icms = Decimal(bucket["total_base_icms"])
        total_icms_value = Decimal(bucket["total_icms_value"])
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        total_effective_rate = compute_display_icms_rate(
            Decimal(bucket["icms_rate"]),
            total_operation_value,
            total_base_icms,
            total_icms_value,
        )
        values.extend(
            [
                int(bucket["total_rows"]),
                total_operation_value,
                total_base_icms,
                Decimal(bucket["total_base_gap"]),
                total_base_ratio,
                total_icms_value,
                total_effective_rate,
            ]
        )
        display_rows.append(
            {
                "reason": bucket["reason"],
                "source_register": bucket["source_register"],
                "cst_icms": bucket["cst_icms"],
                "cfop": bucket["cfop"],
                "icms_rate": Decimal(bucket["icms_rate"]),
                "values": values,
            }
        )
        export_rows.append(values)

    return export_headers, display_rows, export_rows
