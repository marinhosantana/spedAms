from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from app.parsers.sped_parser import normalize_document_key, parse_decimal, read_sped_summary_register_rows
from app.services.analysis_utils import calculate_abc_curve_labels, infer_sped_period_label, period_label_sort_key
from app.services.tax_rules import normalize_operation_type, normalize_text


def build_product_monthly_summary_rows(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> list[dict[str, object]]:
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    totals_by_code: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))

    for row in filtered_rows:
        code = str(row.get("code", "")).strip()
        period = str(row.get("period", "")).strip()
        if not code or not period:
            continue

        key = (code, period)
        bucket = grouped_rows.setdefault(
            key,
            {
                "period": period,
                "operation_type": operation_type,
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "document_keys": set(),
                "launch_count": 0,
                "sale_value": Decimal("0"),
                "base_icms": Decimal("0"),
                "icms_value": Decimal("0"),
            },
        )

        if not bucket["description"] and str(row.get("description", "")).strip():
            bucket["description"] = str(row.get("description", "")).strip()
        if not bucket["ncm"] and str(row.get("ncm", "")).strip():
            bucket["ncm"] = str(row.get("ncm", "")).strip()

        for cfop in str(row.get("cfop", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                bucket["cfops"].add(cfop_text)
        for cst in str(row.get("cst_icms", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                bucket["csts"].add(cst_text)

        bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        bucket["base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        bucket["icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

        launch_details = list(row.get("launch_details", []))
        bucket["launch_count"] += len(launch_details)
        for detail in launch_details:
            document_identity = (
                normalize_document_key(str(detail.get("document_key", "")))
                or str(detail.get("document_number", "")).strip()
            )
            if document_identity:
                bucket["document_keys"].add(document_identity)

        totals_by_code[code] += Decimal(row.get("sale_value", Decimal("0")))

    curve_labels = calculate_abc_curve_labels(totals_by_code)
    monthly_rows: list[dict[str, object]] = []
    for (_, _), item in grouped_rows.items():
        base_icms = Decimal(item["base_icms"])
        icms_value = Decimal(item["icms_value"])
        monthly_rows.append(
            {
                "period": item["period"],
                "operation_type": item["operation_type"],
                "curve_abc": curve_labels.get(str(item["code"]).strip(), "C"),
                "code": item["code"],
                "description": item["description"],
                "ncm": item["ncm"],
                "cfops": " | ".join(sorted(item["cfops"])),
                "csts": " | ".join(sorted(item["csts"])),
                "document_count": len(item["document_keys"]),
                "launch_count": int(item["launch_count"]),
                "sale_value": Decimal(item["sale_value"]),
                "base_icms": base_icms,
                "icms_value": icms_value,
                "calculated_icms_rate": (
                    (icms_value * Decimal("100") / base_icms).quantize(Decimal("0.01"))
                    if base_icms > 0
                    else Decimal("0.00")
                ),
            }
        )

    monthly_rows.sort(
        key=lambda item: (
            normalize_text(str(item.get("period", ""))),
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("code", ""))),
        )
    )
    return monthly_rows

def build_product_monthly_linear_dataset(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]]]:
    monthly_rows = build_product_monthly_summary_rows(filtered_rows, operation_type)
    periods = sorted({str(row.get("period", "")).strip() for row in monthly_rows if str(row.get("period", "")).strip()}, key=period_label_sort_key)
    grouped_products: dict[str, dict[str, object]] = {}

    for row in monthly_rows:
        code = str(row.get("code", "")).strip()
        if not code:
            continue
        product_bucket = grouped_products.setdefault(
            code,
            {
                "curve_abc": str(row.get("curve_abc", "C")).strip() or "C",
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "periods": {},
                "total_sale_value": Decimal("0"),
                "total_base_icms": Decimal("0"),
                "total_icms_value": Decimal("0"),
            },
        )
        if not product_bucket["description"] and str(row.get("description", "")).strip():
            product_bucket["description"] = str(row.get("description", "")).strip()
        if not product_bucket["ncm"] and str(row.get("ncm", "")).strip():
            product_bucket["ncm"] = str(row.get("ncm", "")).strip()

        for cfop in str(row.get("cfops", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                product_bucket["cfops"].add(cfop_text)
        for cst in str(row.get("csts", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                product_bucket["csts"].add(cst_text)

        period = str(row.get("period", "")).strip()
        product_bucket["periods"][period] = {
            "sale_value": Decimal(row.get("sale_value", Decimal("0"))),
            "base_icms": Decimal(row.get("base_icms", Decimal("0"))),
            "icms_value": Decimal(row.get("icms_value", Decimal("0"))),
            "calculated_icms_rate": Decimal(row.get("calculated_icms_rate", Decimal("0"))),
        }
        product_bucket["total_sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        product_bucket["total_base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
        product_bucket["total_icms_value"] += Decimal(row.get("icms_value", Decimal("0")))

    fixed_headers = ["Curva ABC", "Codigo", "Descricao", "NCM", "CFOPs", "CSTs"]
    dynamic_headers: list[str] = []
    for period in periods:
        dynamic_headers.extend(
            [
                f"{period} Valor Operacao",
                f"{period} Base ICMS",
                f"{period} Valor ICMS",
                f"{period} Aliquota",
            ]
        )
    export_headers = [
        *fixed_headers,
        *dynamic_headers,
        "Total Valor Operacao",
        "Total Base ICMS",
        "Total Valor ICMS",
        "Aliquota Total",
    ]

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for product in sorted(
        grouped_products.values(),
        key=lambda item: (
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        ),
    ):
        total_base_icms = Decimal(product["total_base_icms"])
        total_icms_value = Decimal(product["total_icms_value"])
        total_rate = (total_icms_value * Decimal("100") / total_base_icms).quantize(Decimal("0.01")) if total_base_icms > 0 else Decimal("0.00")
        values: list[object] = [
            product["curve_abc"],
            product["code"],
            product["description"],
            product["ncm"],
            " | ".join(sorted(product["cfops"])),
            " | ".join(sorted(product["csts"])),
        ]
        for period in periods:
            period_data = product["periods"].get(period)
            if period_data is None:
                values.extend([Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")])
                continue
            values.extend(
                [
                    Decimal(period_data["sale_value"]),
                    Decimal(period_data["base_icms"]),
                    Decimal(period_data["icms_value"]),
                    Decimal(period_data["calculated_icms_rate"]),
                ]
            )
        values.extend(
            [
                Decimal(product["total_sale_value"]),
                total_base_icms,
                total_icms_value,
                total_rate,
            ]
        )
        display_rows.append({"curve_abc": product["curve_abc"], "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows

def build_contrib_operation_summary_rows(
    detailed_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[dict[str, object]], dict[str, Decimal]]:
    grouped_rows: dict[tuple[str, str, str, Decimal, Decimal], dict[str, object]] = {}
    normalized_operation = normalize_operation_type(operation_type)

    for source_row in detailed_rows:
        if normalize_operation_type(str(source_row.get("operation_type", ""))) != normalized_operation:
            continue
        cst_pis = str(source_row.get("cst_pis", "")).strip()
        cst_cofins = str(source_row.get("cst_cofins", "")).strip()
        cfop = str(source_row.get("cfop", "")).strip()
        aliquota_pis = source_row.get("aliquota_pis", Decimal("0"))
        aliquota_cofins = source_row.get("aliquota_cofins", Decimal("0"))
        aliquota_pis = aliquota_pis if isinstance(aliquota_pis, Decimal) else Decimal("0")
        aliquota_cofins = aliquota_cofins if isinstance(aliquota_cofins, Decimal) else Decimal("0")
        key = (
            cst_pis,
            cst_cofins,
            cfop,
            aliquota_pis.quantize(Decimal("0.01")),
            aliquota_cofins.quantize(Decimal("0.01")),
        )
        sale_value = source_row.get("sale_value", Decimal("0"))
        base_pis = source_row.get("base_pis", Decimal("0"))
        base_cofins = source_row.get("base_cofins", Decimal("0"))
        pis_value = source_row.get("pis_value", Decimal("0"))
        cofins_value = source_row.get("cofins_value", Decimal("0"))
        sale_value = sale_value if isinstance(sale_value, Decimal) else Decimal("0")
        base_pis = base_pis if isinstance(base_pis, Decimal) else Decimal("0")
        base_cofins = base_cofins if isinstance(base_cofins, Decimal) else Decimal("0")
        pis_value = pis_value if isinstance(pis_value, Decimal) else Decimal("0")
        cofins_value = cofins_value if isinstance(cofins_value, Decimal) else Decimal("0")

        if key not in grouped_rows:
            grouped_rows[key] = {
                "operation_type": normalized_operation,
                "cst_pis": cst_pis,
                "cst_cofins": cst_cofins,
                "cfop": cfop,
                "aliquota_pis": key[3],
                "aliquota_cofins": key[4],
                "sale_value": Decimal("0"),
                "base_pis": Decimal("0"),
                "pis_value": Decimal("0"),
                "base_cofins": Decimal("0"),
                "cofins_value": Decimal("0"),
                "document_keys": set(),
                "launch_count": 0,
            }

        grouped_rows[key]["sale_value"] += sale_value
        grouped_rows[key]["base_pis"] += base_pis
        grouped_rows[key]["pis_value"] += pis_value
        grouped_rows[key]["base_cofins"] += base_cofins
        grouped_rows[key]["cofins_value"] += cofins_value
        grouped_rows[key]["launch_count"] += 1
        document_identity = normalize_document_key(str(source_row.get("document_key", ""))) or str(source_row.get("document_number", "")).strip()
        if document_identity:
            grouped_rows[key]["document_keys"].add(document_identity)

    rows = sorted(
        grouped_rows.values(),
        key=lambda row: (
            str(row["cst_pis"]),
            str(row["cst_cofins"]),
            str(row["cfop"]),
            Decimal(row["aliquota_pis"]),
            Decimal(row["aliquota_cofins"]),
        ),
    )
    totals = {
        "sale_value": sum((Decimal(row["sale_value"]) for row in rows), Decimal("0")),
        "base_pis": sum((Decimal(row["base_pis"]) for row in rows), Decimal("0")),
        "pis_value": sum((Decimal(row["pis_value"]) for row in rows), Decimal("0")),
        "base_cofins": sum((Decimal(row["base_cofins"]) for row in rows), Decimal("0")),
        "cofins_value": sum((Decimal(row["cofins_value"]) for row in rows), Decimal("0")),
    }
    return rows, totals

def build_contrib_operation_launch_details_map(
    detailed_rows: list[dict[str, object]],
    operation_type: str,
) -> dict[tuple[str, str, str, Decimal, Decimal], list[dict[str, object]]]:
    grouped_details: dict[tuple[str, str, str, Decimal, Decimal], list[dict[str, object]]] = defaultdict(list)
    normalized_operation = normalize_operation_type(operation_type)

    for item in detailed_rows:
        if normalize_operation_type(str(item.get("operation_type", ""))) != normalized_operation:
            continue
        cst_pis = str(item.get("cst_pis", "")).strip()
        cst_cofins = str(item.get("cst_cofins", "")).strip()
        cfop = str(item.get("cfop", "")).strip()
        aliquota_pis = item.get("aliquota_pis", Decimal("0"))
        aliquota_cofins = item.get("aliquota_cofins", Decimal("0"))
        aliquota_pis = aliquota_pis if isinstance(aliquota_pis, Decimal) else Decimal("0")
        aliquota_cofins = aliquota_cofins if isinstance(aliquota_cofins, Decimal) else Decimal("0")
        key = (
            cst_pis,
            cst_cofins,
            cfop,
            aliquota_pis.quantize(Decimal("0.01")),
            aliquota_cofins.quantize(Decimal("0.01")),
        )
        grouped_details[key].append(dict(item))

    return grouped_details

def build_contrib_product_monthly_summary_rows(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> list[dict[str, object]]:
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    totals_by_code: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))

    for row in filtered_rows:
        code = str(row.get("code", "")).strip()
        period = str(row.get("period", "")).strip()
        if not code or not period:
            continue

        key = (code, period)
        bucket = grouped_rows.setdefault(
            key,
            {
                "period": period,
                "operation_type": operation_type,
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "document_keys": set(),
                "launch_count": 0,
                "sale_value": Decimal("0"),
                "base_pis": Decimal("0"),
                "pis_value": Decimal("0"),
                "base_cofins": Decimal("0"),
                "cofins_value": Decimal("0"),
            },
        )

        if not bucket["description"] and str(row.get("description", "")).strip():
            bucket["description"] = str(row.get("description", "")).strip()
        if not bucket["ncm"] and str(row.get("ncm", "")).strip():
            bucket["ncm"] = str(row.get("ncm", "")).strip()
        for cfop in str(row.get("cfop", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                bucket["cfops"].add(cfop_text)
        for cst in (str(row.get("cst_pis", "")).split("|") + str(row.get("cst_cofins", "")).split("|")):
            cst_text = cst.strip()
            if cst_text:
                bucket["csts"].add(cst_text)

        bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        bucket["base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
        bucket["pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
        bucket["base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
        bucket["cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

        launch_details = list(row.get("launch_details", []))
        bucket["launch_count"] += len(launch_details)
        for detail in launch_details:
            document_identity = normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            if document_identity:
                bucket["document_keys"].add(document_identity)

        totals_by_code[code] += Decimal(row.get("sale_value", Decimal("0")))

    curve_labels = calculate_abc_curve_labels(totals_by_code)
    monthly_rows: list[dict[str, object]] = []
    for item in grouped_rows.values():
        monthly_rows.append(
            {
                "period": item["period"],
                "operation_type": item["operation_type"],
                "curve_abc": curve_labels.get(str(item["code"]).strip(), "C"),
                "code": item["code"],
                "description": item["description"],
                "ncm": item["ncm"],
                "cfops": " | ".join(sorted(item["cfops"])),
                "csts": " | ".join(sorted(item["csts"])),
                "document_count": len(item["document_keys"]),
                "launch_count": int(item["launch_count"]),
                "sale_value": Decimal(item["sale_value"]),
                "base_pis": Decimal(item["base_pis"]),
                "pis_value": Decimal(item["pis_value"]),
                "base_cofins": Decimal(item["base_cofins"]),
                "cofins_value": Decimal(item["cofins_value"]),
            }
        )

    monthly_rows.sort(
        key=lambda item: (
            normalize_text(str(item.get("period", ""))),
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        )
    )
    return monthly_rows

def build_contrib_product_monthly_linear_dataset(
    filtered_rows: list[dict[str, object]],
    operation_type: str,
) -> tuple[list[str], list[str], list[dict[str, object]], list[list[object]]]:
    monthly_rows = build_contrib_product_monthly_summary_rows(filtered_rows, operation_type)
    periods = sorted({str(row.get("period", "")).strip() for row in monthly_rows if str(row.get("period", "")).strip()}, key=period_label_sort_key)
    grouped_products: dict[str, dict[str, object]] = {}

    for row in monthly_rows:
        code = str(row.get("code", "")).strip()
        if not code:
            continue
        product_bucket = grouped_products.setdefault(
            code,
            {
                "curve_abc": str(row.get("curve_abc", "C")).strip() or "C",
                "code": code,
                "description": str(row.get("description", "")).strip(),
                "ncm": str(row.get("ncm", "")).strip(),
                "cfops": set(),
                "csts": set(),
                "periods": {},
                "total_sale_value": Decimal("0"),
                "total_base_pis": Decimal("0"),
                "total_pis_value": Decimal("0"),
                "total_base_cofins": Decimal("0"),
                "total_cofins_value": Decimal("0"),
            },
        )
        for cfop in str(row.get("cfops", "")).split("|"):
            cfop_text = cfop.strip()
            if cfop_text:
                product_bucket["cfops"].add(cfop_text)
        for cst in str(row.get("csts", "")).split("|"):
            cst_text = cst.strip()
            if cst_text:
                product_bucket["csts"].add(cst_text)
        period = str(row.get("period", "")).strip()
        product_bucket["periods"][period] = {
            "sale_value": Decimal(row.get("sale_value", Decimal("0"))),
            "base_pis": Decimal(row.get("base_pis", Decimal("0"))),
            "pis_value": Decimal(row.get("pis_value", Decimal("0"))),
            "base_cofins": Decimal(row.get("base_cofins", Decimal("0"))),
            "cofins_value": Decimal(row.get("cofins_value", Decimal("0"))),
        }
        product_bucket["total_sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
        product_bucket["total_base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
        product_bucket["total_pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
        product_bucket["total_base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
        product_bucket["total_cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

    fixed_headers = ["Curva ABC", "Codigo", "Descricao", "NCM", "CFOPs", "CSTs"]
    dynamic_headers: list[str] = []
    for period in periods:
        dynamic_headers.extend(
            [
                f"{period} Valor Operacao",
                f"{period} Base PIS",
                f"{period} Valor PIS",
                f"{period} Base COFINS",
                f"{period} Valor COFINS",
            ]
        )
    export_headers = [
        *fixed_headers,
        *dynamic_headers,
        "Total Valor Operacao",
        "Total Base PIS",
        "Total Valor PIS",
        "Total Base COFINS",
        "Total Valor COFINS",
    ]

    display_rows: list[dict[str, object]] = []
    export_rows: list[list[object]] = []
    for product in sorted(
        grouped_products.values(),
        key=lambda item: (
            normalize_text(str(item.get("curve_abc", ""))),
            normalize_text(str(item.get("description", ""))),
            normalize_text(str(item.get("code", ""))),
        ),
    ):
        values: list[object] = [
            product["curve_abc"],
            product["code"],
            product["description"],
            product["ncm"],
            " | ".join(sorted(product["cfops"])),
            " | ".join(sorted(product["csts"])),
        ]
        for period in periods:
            period_data = product["periods"].get(
                period,
                {
                    "sale_value": Decimal("0"),
                    "base_pis": Decimal("0"),
                    "pis_value": Decimal("0"),
                    "base_cofins": Decimal("0"),
                    "cofins_value": Decimal("0"),
                },
            )
            values.extend(
                [
                    Decimal(period_data["sale_value"]),
                    Decimal(period_data["base_pis"]),
                    Decimal(period_data["pis_value"]),
                    Decimal(period_data["base_cofins"]),
                    Decimal(period_data["cofins_value"]),
                ]
            )
        values.extend(
            [
                Decimal(product["total_sale_value"]),
                Decimal(product["total_base_pis"]),
                Decimal(product["total_pis_value"]),
                Decimal(product["total_base_cofins"]),
                Decimal(product["total_cofins_value"]),
            ]
        )
        display_rows.append({"curve_abc": product["curve_abc"], "values": values})
        export_rows.append(values)

    return periods, export_headers, display_rows, export_rows

def build_monthly_aliquota_divergence_rows(
    headers: list[str],
    rows: list[list[object]],
    periods: list[str],
) -> list[list[object]]:
    fixed_column_count = 6
    monthly_block_size = 4
    divergence_rows: list[list[object]] = []

    for row in rows:
        valid_rates: list[Decimal] = []
        for period_index, _period in enumerate(periods):
            start = fixed_column_count + (period_index * monthly_block_size)
            if start + 3 >= len(row):
                continue
            sale_value = row[start]
            base_icms = row[start + 1]
            icms_value = row[start + 2]
            calculated_rate = row[start + 3]
            sale_value = sale_value if isinstance(sale_value, Decimal) else parse_decimal(str(sale_value or "0"))
            base_icms = base_icms if isinstance(base_icms, Decimal) else parse_decimal(str(base_icms or "0"))
            icms_value = icms_value if isinstance(icms_value, Decimal) else parse_decimal(str(icms_value or "0"))
            calculated_rate = calculated_rate if isinstance(calculated_rate, Decimal) else parse_decimal(str(calculated_rate or "0"))

            if base_icms <= 0 and sale_value <= 0 and icms_value <= 0:
                continue
            if base_icms <= 0:
                continue
            valid_rates.append(calculated_rate.quantize(Decimal("0.01")))

        distinct_rates = {rate for rate in valid_rates}
        if len(distinct_rates) > 1:
            divergence_rows.append(list(row))

    return divergence_rows

def build_summary_operation_totals_by_period(
    sped_paths: list[Path],
    operation_type: str,
) -> dict[str, dict[str, Decimal]]:
    totals_by_period: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: {
            "total_operation_value": Decimal("0"),
            "base_icms": Decimal("0"),
            "icms_value": Decimal("0"),
        }
    )
    normalized_operation = normalize_operation_type(operation_type)
    for sped_path in sped_paths:
        if not sped_path.exists():
            continue
        period_label = infer_sped_period_label(sped_path, [])
        summary_rows = read_sped_summary_register_rows(sped_path)
        for row in summary_rows:
            if normalize_operation_type(str(row.get("operation_type", ""))) != normalized_operation:
                continue
            totals_by_period[period_label]["total_operation_value"] += Decimal(row.get("total_operation_value", Decimal("0")))
            totals_by_period[period_label]["base_icms"] += Decimal(row.get("base_icms", Decimal("0")))
            totals_by_period[period_label]["icms_value"] += Decimal(row.get("icms_value", Decimal("0")))
    return totals_by_period
