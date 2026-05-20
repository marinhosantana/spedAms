from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Callable

from app.models import CompareInvestigationMatchedRow, CompareSheetInvoice, CompareSpedDocument, CompareXmlInvoice, CompareXmlItem
from app.parsers.compare_sheet import collect_compare_sheet_invoices
from app.parsers.compare_sped_reader import collect_compare_sped_documents, extract_company_tax_id_from_sped
from app.parsers.compare_xml import (
    describe_compare_ignored_xml,
    parse_compare_xml_file,
    collect_xml_cancellation_events,
)
from app.parsers.sped_parser import normalize_document_key
from app.services.compare_matching import (
    build_compare_sped_number_indexes,
    compare_decimal_value,
    describe_compare_xml_sped_value_difference,
    find_compare_sped_document_by_key_or_number,
)
from app.services.compare_operations import (
    classify_xml_item_operation,
    infer_compare_invoice_operation,
    normalize_compare_operation_scope,
)
from app.services.compare_sped_launcher import compare_register
from app.services.tax_rules import format_decimal_sped

def collect_compare_xml_invoices(
    xml_folder: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareXmlInvoice], int]:
    invoices: list[CompareXmlInvoice] = []
    ignored = 0
    xml_paths = sorted(xml_folder.rglob("*.xml"))
    total_xml = max(len(xml_paths), 1)
    if progress_callback:
        progress_callback(0, total_xml, f"Localizados {len(xml_paths)} XML(s).")
    for index, xml_path in enumerate(xml_paths, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
        else:
            invoices.append(invoice)
        if progress_callback and (index == total_xml or index % 10 == 0):
            progress_callback(index, total_xml, f"Lendo XML {index}/{total_xml}: {xml_path.name}")
    return invoices, ignored

def xml_contains_company_tax_id(xml_path: Path, company_tax_id: str) -> bool:
    company_key = normalize_document_key(company_tax_id)
    if not company_key:
        return False
    try:
        text = xml_path.read_text(encoding="utf-8", errors="ignore")
    except UnicodeDecodeError:
        text = xml_path.read_text(encoding="latin-1", errors="ignore")
    return company_key in normalize_document_key(text)

def build_xml_cfop_summary_rows(
    xml_source: Path,
    company_tax_id: str = "",
    operation_scope: str = "Todos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[dict[str, object]], dict[str, object]]:
    xml_files = [xml_source] if xml_source.is_file() else sorted(xml_source.rglob("*.xml"))
    total_files = max(len(xml_files), 1)
    grouped_rows: dict[tuple[str, str], dict[str, object]] = {}
    ignored = 0
    skipped_by_company = 0
    ignored_log_rows: list[dict[str, object]] = []
    document_log_rows: list[dict[str, object]] = []
    cancelled_display_rows: list[dict[str, object]] = []
    document_keys: set[str] = set()
    normalized_company_tax_id = normalize_document_key(company_tax_id)
    normalized_operation_scope = normalize_compare_operation_scope(operation_scope)
    cancellation_events = collect_xml_cancellation_events(xml_files)
    cancelled_documents = 0
    cancelled_value_total = Decimal("0")

    if progress_callback:
        progress_callback(0, total_files, f"Localizados {len(xml_files)} XML(s).")

    for index, xml_path in enumerate(xml_files, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
            ignored_log_rows.append(
                {
                    "reason": describe_compare_ignored_xml(xml_path),
                    "company_tax_id": normalized_company_tax_id,
                    "key": "",
                    "number": "",
                    "series": "",
                    "issue_date": "",
                    "issuer_cnpj": "",
                    "issuer_name": "",
                    "recipient_cnpj": "",
                    "recipient_name": "",
                    "total_value": Decimal("0"),
                    "file_path": str(xml_path),
                }
            )
        else:
            issuer_key = normalize_document_key(invoice.issuer_cnpj)
            recipient_key = normalize_document_key(invoice.recipient_cnpj)
            company_found_in_xml = (
                not normalized_company_tax_id
                or normalized_company_tax_id in {issuer_key, recipient_key}
                or xml_contains_company_tax_id(xml_path, normalized_company_tax_id)
            )
            if normalized_company_tax_id and not company_found_in_xml:
                skipped_by_company += 1
                ignored_log_rows.append(
                    {
                        "reason": "Fora da empresa",
                        "company_tax_id": normalized_company_tax_id,
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "issuer_cnpj": issuer_key,
                        "issuer_name": invoice.issuer_name,
                        "recipient_cnpj": recipient_key,
                        "recipient_name": invoice.recipient_name,
                        "total_value": Decimal(str(invoice.total_doc or 0)),
                        "file_path": str(xml_path),
                    }
                )
                if progress_callback and (index == total_files or index % 10 == 0):
                    progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
                continue
            if invoice.key:
                document_keys.add(invoice.key)
            cancellation_event = cancellation_events.get(invoice.key)
            if cancellation_event is not None:
                cancelled_documents += 1
                cancelled_value_total += Decimal(str(invoice.total_doc or 0))
                cancelled_cfops = sorted({str(item.cfop or "").strip() for item in invoice.items if str(item.cfop or "").strip()})
                cancelled_display_rows.append(
                    {
                        "operation_type": "Cancelado",
                        "cfop": " | ".join(cancelled_cfops) or "(sem CFOP)",
                        "document_count": 1,
                        "items": len(invoice.items),
                        "operation_value": Decimal(str(invoice.total_products or invoice.total_doc or 0)),
                        "discount_value": Decimal(str(invoice.total_discount or 0)),
                        "base_icms": Decimal(str(invoice.total_bc_icms or 0)),
                        "icms_value": Decimal(str(invoice.total_icms or 0)),
                        "base_icms_st": Decimal(str(invoice.total_bc_st or 0)),
                        "icms_st_value": Decimal(str(invoice.total_st or 0)),
                        "ipi_value": Decimal(str(invoice.total_ipi or 0)),
                        "base_pis": Decimal("0"),
                        "pis_value": Decimal(str(invoice.total_pis or 0)),
                        "base_cofins": Decimal("0"),
                        "cofins_value": Decimal(str(invoice.total_cofins or 0)),
                        "details": [],
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "file_path": str(xml_path),
                        "cancel_reason": cancellation_event.get("reason", ""),
                    }
                )
                ignored_log_rows.append(
                    {
                        "reason": f"NF-e/NFC-e cancelada: {cancellation_event.get('reason', '')}",
                        "company_tax_id": normalized_company_tax_id,
                        "key": invoice.key,
                        "number": invoice.number,
                        "series": invoice.series,
                        "issue_date": invoice.issue_date,
                        "issuer_cnpj": issuer_key,
                        "issuer_name": invoice.issuer_name,
                        "recipient_cnpj": recipient_key,
                        "recipient_name": invoice.recipient_name,
                        "total_value": Decimal(str(invoice.total_doc or 0)),
                        "file_path": str(xml_path),
                    }
                )
                if progress_callback and (index == total_files or index % 10 == 0):
                    progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
                continue
            invoice_operation_types = sorted(
                {
                    classify_xml_item_operation(invoice, item, normalized_company_tax_id)
                    for item in invoice.items
                }
            )
            if not invoice_operation_types:
                invoice_operation_types = [infer_compare_invoice_operation(invoice, normalized_company_tax_id)]
            document_log_rows.append(
                {
                    "operation_type": " | ".join(operation for operation in invoice_operation_types if operation),
                    "key": invoice.key,
                    "number": invoice.number,
                    "series": invoice.series,
                    "issue_date": invoice.issue_date,
                    "issuer_cnpj": issuer_key,
                    "issuer_name": invoice.issuer_name,
                    "recipient_cnpj": recipient_key,
                    "recipient_name": invoice.recipient_name,
                    "total_value": Decimal(str(invoice.total_doc or 0)),
                    "item_count": len(invoice.items),
                    "file_path": str(xml_path),
                }
            )
            for item in invoice.items:
                cfop = str(item.cfop or "").strip() or "(sem CFOP)"
                operation_type = classify_xml_item_operation(invoice, item, normalized_company_tax_id)
                if (
                    normalized_company_tax_id
                    and normalized_company_tax_id not in {issuer_key, recipient_key}
                    and company_found_in_xml
                ):
                    operation_type = "Entrada"
                if normalized_operation_scope != "Ambos" and operation_type != normalized_operation_scope:
                    continue
                key = (operation_type, cfop)
                bucket = grouped_rows.setdefault(
                    key,
                    {
                        "operation_type": operation_type,
                        "cfop": cfop,
                        "documents": set(),
                        "items": 0,
                        "operation_value": Decimal("0"),
                        "discount_value": Decimal("0"),
                        "base_icms": Decimal("0"),
                        "icms_value": Decimal("0"),
                        "base_icms_st": Decimal("0"),
                        "icms_st_value": Decimal("0"),
                        "ipi_value": Decimal("0"),
                        "base_pis": Decimal("0"),
                        "pis_value": Decimal("0"),
                        "base_cofins": Decimal("0"),
                        "cofins_value": Decimal("0"),
                        "details": [],
                    },
                )
                if invoice.key:
                    bucket["documents"].add(invoice.key)
                bucket["items"] += 1
                detail_row = {
                    "operation_type": operation_type,
                    "cfop": cfop,
                    "key": invoice.key,
                    "number": invoice.number,
                    "series": invoice.series,
                    "issue_date": invoice.issue_date,
                    "issuer_cnpj": issuer_key,
                    "issuer_name": invoice.issuer_name,
                    "recipient_cnpj": recipient_key,
                    "recipient_name": invoice.recipient_name,
                    "item_no": item.item_no,
                    "code": item.code,
                    "description": item.description,
                    "operation_value": Decimal(str(item.value or 0)),
                    "discount_value": Decimal(str(item.discount or 0)),
                    "base_icms": Decimal(str(item.vl_bc_icms or 0)),
                    "icms_value": Decimal(str(item.vl_icms or 0)),
                    "base_icms_st": Decimal(str(item.vl_bc_icms_st or 0)),
                    "icms_st_value": Decimal(str(item.vl_icms_st or 0)),
                    "ipi_value": Decimal(str(item.vl_ipi or 0)),
                    "base_pis": Decimal(str(item.vl_bc_pis or 0)),
                    "pis_value": Decimal(str(item.vl_pis or 0)),
                    "base_cofins": Decimal(str(item.vl_bc_cofins or 0)),
                    "cofins_value": Decimal(str(item.vl_cofins or 0)),
                    "file_path": str(xml_path),
                }
                bucket["details"].append(detail_row)
                bucket["operation_value"] += Decimal(str(item.value or 0))
                bucket["discount_value"] += Decimal(str(item.discount or 0))
                bucket["base_icms"] += Decimal(str(item.vl_bc_icms or 0))
                bucket["icms_value"] += Decimal(str(item.vl_icms or 0))
                bucket["base_icms_st"] += Decimal(str(item.vl_bc_icms_st or 0))
                bucket["icms_st_value"] += Decimal(str(item.vl_icms_st or 0))
                bucket["ipi_value"] += Decimal(str(item.vl_ipi or 0))
                bucket["base_pis"] += Decimal(str(item.vl_bc_pis or 0))
                bucket["pis_value"] += Decimal(str(item.vl_pis or 0))
                bucket["base_cofins"] += Decimal(str(item.vl_bc_cofins or 0))
                bucket["cofins_value"] += Decimal(str(item.vl_cofins or 0))
        if progress_callback and (index == total_files or index % 10 == 0):
            progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")

    rows: list[dict[str, object]] = []
    for bucket in grouped_rows.values():
        documents = bucket.pop("documents")
        row = dict(bucket)
        row["document_count"] = len(documents)
        rows.append(row)
    rows.sort(key=lambda row: (str(row.get("operation_type", "")), str(row.get("cfop", ""))))
    stats = {
        "xml_total": len(xml_files),
        "xml_valid": len(document_keys),
        "xml_ignored": ignored,
        "xml_skipped_by_company": skipped_by_company,
        "xml_cancelled": cancelled_documents,
        "xml_cancelled_value_total": cancelled_value_total.quantize(Decimal("0.01")),
        "company_tax_id": normalized_company_tax_id,
        "ignored_log_rows": ignored_log_rows,
        "document_log_rows": document_log_rows,
        "cancellation_events": list(cancellation_events.values()),
        "cancelled_display_rows": cancelled_display_rows,
    }
    return rows, stats

def build_xml_entry_credit_rows(
    xml_source: Path,
    company_tax_id: str = "",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    xml_files = [xml_source] if xml_source.is_file() else sorted(xml_source.rglob("*.xml"))
    total_files = max(len(xml_files), 1)
    normalized_company_tax_id = normalize_document_key(company_tax_id)
    invoice_rows: list[dict[str, object]] = []
    grouped_cfop: dict[str, dict[str, object]] = {}
    ignored = 0
    skipped_by_company = 0
    cancellation_events = collect_xml_cancellation_events(xml_files)
    cancelled_documents = 0
    cancelled_value_total = Decimal("0")

    if progress_callback:
        progress_callback(0, total_files, f"Localizados {len(xml_files)} XML(s).")

    for index, xml_path in enumerate(xml_files, start=1):
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            ignored += 1
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        issuer_key = normalize_document_key(invoice.issuer_cnpj)
        recipient_key = normalize_document_key(invoice.recipient_cnpj)
        company_found_in_xml = (
            not normalized_company_tax_id
            or normalized_company_tax_id in {issuer_key, recipient_key}
            or xml_contains_company_tax_id(xml_path, normalized_company_tax_id)
        )
        if normalized_company_tax_id and not company_found_in_xml:
            skipped_by_company += 1
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue
        if cancellation_events.get(invoice.key) is not None:
            cancelled_documents += 1
            cancelled_value_total += Decimal(str(invoice.total_doc or 0))
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        entry_items: list[CompareXmlItem] = []
        for item in invoice.items:
            operation_type = classify_xml_item_operation(invoice, item, normalized_company_tax_id)
            if (
                normalized_company_tax_id
                and normalized_company_tax_id not in {issuer_key, recipient_key}
                and company_found_in_xml
            ):
                operation_type = "Entrada"
            if operation_type == "Entrada":
                entry_items.append(item)

        if not entry_items:
            if progress_callback and (index == total_files or index % 10 == 0):
                progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")
            continue

        details: list[dict[str, object]] = []
        cfops: set[str] = set()
        total_operation = Decimal("0")
        total_discount = Decimal("0")
        total_base_icms = Decimal("0")
        total_icms = Decimal("0")
        total_base_icms_st = Decimal("0")
        total_icms_st = Decimal("0")
        total_ipi = Decimal("0")

        for item in entry_items:
            cfop = str(item.cfop or "").strip() or "(sem CFOP)"
            cfops.add(cfop)
            item_operation = Decimal(str(item.value or 0))
            item_discount = Decimal(str(item.discount or 0))
            item_base_icms = Decimal(str(item.vl_bc_icms or 0))
            item_icms = Decimal(str(item.vl_icms or 0))
            item_base_icms_st = Decimal(str(item.vl_bc_icms_st or 0))
            item_icms_st = Decimal(str(item.vl_icms_st or 0))
            item_ipi = Decimal(str(item.vl_ipi or 0))
            detail = {
                "key": invoice.key,
                "number": invoice.number,
                "series": invoice.series,
                "issue_date": invoice.issue_date,
                "issuer_cnpj": issuer_key,
                "issuer_name": invoice.issuer_name,
                "recipient_cnpj": recipient_key,
                "recipient_name": invoice.recipient_name,
                "item_no": item.item_no,
                "code": item.code,
                "description": item.description,
                "ncm": item.ncm,
                "cfop": cfop,
                "cst_icms": item.cst_icms,
                "quantity": Decimal(str(item.quantity or 0)),
                "operation_value": item_operation,
                "discount_value": item_discount,
                "base_icms": item_base_icms,
                "icms_rate": Decimal(str(item.aliq_icms or 0)),
                "icms_value": item_icms,
                "base_icms_st": item_base_icms_st,
                "icms_st_value": item_icms_st,
                "ipi_value": item_ipi,
                "file_path": str(xml_path),
            }
            details.append(detail)
            total_operation += item_operation
            total_discount += item_discount
            total_base_icms += item_base_icms
            total_icms += item_icms
            total_base_icms_st += item_base_icms_st
            total_icms_st += item_icms_st
            total_ipi += item_ipi

            bucket = grouped_cfop.setdefault(
                cfop,
                {
                    "cfop": cfop,
                    "documents": set(),
                    "items": 0,
                    "operation_value": Decimal("0"),
                    "discount_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "base_icms_st": Decimal("0"),
                    "icms_st_value": Decimal("0"),
                    "ipi_value": Decimal("0"),
                    "details": [],
                },
            )
            if invoice.key:
                bucket["documents"].add(invoice.key)
            bucket["items"] += 1
            bucket["operation_value"] += item_operation
            bucket["discount_value"] += item_discount
            bucket["base_icms"] += item_base_icms
            bucket["icms_value"] += item_icms
            bucket["base_icms_st"] += item_base_icms_st
            bucket["icms_st_value"] += item_icms_st
            bucket["ipi_value"] += item_ipi
            bucket["details"].append(detail)

        invoice_rows.append(
            {
                "key": invoice.key,
                "number": invoice.number,
                "series": invoice.series,
                "issue_date": invoice.issue_date,
                "issuer_cnpj": issuer_key,
                "issuer_name": invoice.issuer_name,
                "recipient_cnpj": recipient_key,
                "recipient_name": invoice.recipient_name,
                "cfops": " | ".join(sorted(cfops)),
                "items": len(entry_items),
                "operation_value": total_operation,
                "discount_value": total_discount,
                "base_icms": total_base_icms,
                "icms_value": total_icms,
                "base_icms_st": total_base_icms_st,
                "icms_st_value": total_icms_st,
                "ipi_value": total_ipi,
                "file_path": str(xml_path),
                "details": details,
            }
        )

        if progress_callback and (index == total_files or index % 10 == 0):
            progress_callback(index, total_files, f"Lendo XML {index}/{total_files}: {xml_path.name}")

    cfop_rows: list[dict[str, object]] = []
    for bucket in grouped_cfop.values():
        documents = bucket.pop("documents")
        row = dict(bucket)
        row["document_count"] = len(documents)
        cfop_rows.append(row)
    invoice_rows.sort(key=lambda row: (str(row.get("issue_date", "")), str(row.get("number", "")), str(row.get("series", ""))))
    cfop_rows.sort(key=lambda row: str(row.get("cfop", "")))
    stats = {
        "xml_total": len(xml_files),
        "xml_valid_entries": len(invoice_rows),
        "xml_ignored": ignored,
        "xml_skipped_by_company": skipped_by_company,
        "xml_cancelled": cancelled_documents,
        "xml_cancelled_value_total": cancelled_value_total.quantize(Decimal("0.01")),
        "company_tax_id": normalized_company_tax_id,
        "cancellation_events": list(cancellation_events.values()),
    }
    return invoice_rows, cfop_rows, stats

def collect_compare_fiscal_basis_totals(sped_file: Path, operation_scope: str = "Ambos") -> dict[str, object]:
    selected_scope = normalize_compare_operation_scope(operation_scope)
    totals: dict[str, Decimal] = {
        "c100_total": Decimal("0.00"),
        "a100_total": Decimal("0.00"),
        "c190_total": Decimal("0.00"),
        "c590_total": Decimal("0.00"),
        "d590_total": Decimal("0.00"),
    }
    counts: dict[str, int] = {
        "c100_count": 0,
        "a100_count": 0,
        "c190_count": 0,
        "c590_count": 0,
        "d590_count": 0,
    }
    current_c_operation = ""
    current_d_operation = ""
    current_c100: dict[str, object] | None = None
    current_c500: dict[str, object] | None = None
    current_d500: dict[str, object] | None = None
    c100_c190_rows: list[list[object]] = []
    c500_c590_rows: list[list[object]] = []
    d500_d590_rows: list[list[object]] = []

    def finish_c100_document() -> None:
        nonlocal current_c100
        if not current_c100:
            return
        c100_value = Decimal(current_c100.get("document_value", Decimal("0.00")))
        c190_value = Decimal(current_c100.get("summary_value", Decimal("0.00")))
        difference = (c190_value - c100_value).quantize(Decimal("0.01"))
        if difference:
            c100_c190_rows.append(
                [
                    current_c100.get("operation_type", ""),
                    current_c100.get("model", ""),
                    current_c100.get("key", ""),
                    current_c100.get("number", ""),
                    current_c100.get("series", ""),
                    current_c100.get("issue_date", ""),
                    c100_value,
                    c190_value,
                    difference,
                    current_c100.get("line_number", ""),
                    "; ".join(str(item) for item in current_c100.get("summary_details", [])),
                    Decimal(current_c100.get("vl_merc", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_desc", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_frt", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_seg", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_out", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_st", Decimal("0.00"))),
                    Decimal(current_c100.get("vl_ipi", Decimal("0.00"))),
                ]
            )
        current_c100 = None

    def finish_c500_document() -> None:
        nonlocal current_c500
        if not current_c500:
            return
        summary_value = Decimal(current_c500.get("summary_value", Decimal("0.00"))).quantize(Decimal("0.01"))
        if summary_value:
            c500_c590_rows.append(
                [
                    current_c500.get("operation_type", ""),
                    current_c500.get("model", ""),
                    current_c500.get("number", ""),
                    current_c500.get("series", ""),
                    current_c500.get("issue_date", ""),
                    summary_value,
                    current_c500.get("line_number", ""),
                    "; ".join(str(item) for item in current_c500.get("summary_details", [])),
                ]
            )
        current_c500 = None

    def finish_d500_document() -> None:
        nonlocal current_d500
        if not current_d500:
            return
        summary_value = Decimal(current_d500.get("summary_value", Decimal("0.00"))).quantize(Decimal("0.01"))
        if summary_value:
            d500_d590_rows.append(
                [
                    current_d500.get("operation_type", ""),
                    current_d500.get("model", ""),
                    current_d500.get("number", ""),
                    current_d500.get("series", ""),
                    current_d500.get("issue_date", ""),
                    summary_value,
                    current_d500.get("line_number", ""),
                    "; ".join(str(item) for item in current_d500.get("summary_details", [])),
                ]
            )
        current_d500 = None

    with sped_file.open("r", encoding="latin-1", errors="ignore") as current_file:
        for line_number, raw_line in enumerate(current_file, start=1):
            line = raw_line.strip()
            register = compare_register(line)
            fields = line.split("|")
            if register in {"C100", "C500"}:
                finish_c100_document()
                finish_c500_document()
                current_c_operation = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
            elif register.startswith("C") and register not in {"C190", "C590"} and register in {"C001", "C990"}:
                finish_c100_document()
                finish_c500_document()
                current_c_operation = ""
            if register == "D500":
                finish_d500_document()
                current_d_operation = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
            elif register.startswith("D") and register not in {"D590"} and register in {"D001", "D990"}:
                finish_d500_document()
                current_d_operation = ""

            if register == "C100":
                operation_type = current_c_operation
                if selected_scope == "Ambos" or operation_type == selected_scope:
                    document_value = compare_decimal_value(fields[12] if len(fields) > 12 else "")
                    totals["c100_total"] += document_value
                    counts["c100_count"] += 1
                    current_c100 = {
                        "operation_type": operation_type,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[8].strip() if len(fields) > 8 else "",
                        "key": fields[9].strip() if len(fields) > 9 else "",
                        "issue_date": fields[10].strip() if len(fields) > 10 else "",
                        "document_value": document_value,
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                        "vl_merc": compare_decimal_value(fields[16] if len(fields) > 16 else ""),
                        "vl_desc": compare_decimal_value(fields[14] if len(fields) > 14 else ""),
                        "vl_frt": compare_decimal_value(fields[18] if len(fields) > 18 else ""),
                        "vl_seg": compare_decimal_value(fields[19] if len(fields) > 19 else ""),
                        "vl_out": compare_decimal_value(fields[20] if len(fields) > 20 else ""),
                        "vl_st": compare_decimal_value(fields[24] if len(fields) > 24 else ""),
                        "vl_ipi": compare_decimal_value(fields[25] if len(fields) > 25 else ""),
                    }
                else:
                    current_c100 = None
            elif register == "A100":
                operation_type = "Entrada" if len(fields) > 2 and fields[2].strip() == "0" else "Saida" if len(fields) > 2 and fields[2].strip() == "1" else ""
                if selected_scope == "Ambos" or operation_type == selected_scope:
                    totals["a100_total"] += compare_decimal_value(fields[12] if len(fields) > 12 else "")
                    counts["a100_count"] += 1
            elif register == "C500":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    current_c500 = {
                        "operation_type": current_c_operation,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[10].strip() if len(fields) > 10 else "",
                        "issue_date": fields[11].strip() if len(fields) > 11 else "",
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                    }
                else:
                    current_c500 = None
            elif register == "D500":
                if selected_scope == "Ambos" or current_d_operation == selected_scope:
                    current_d500 = {
                        "operation_type": current_d_operation,
                        "model": fields[5].strip() if len(fields) > 5 else "",
                        "series": fields[7].strip() if len(fields) > 7 else "",
                        "number": fields[8].strip() if len(fields) > 8 else "",
                        "issue_date": fields[10].strip() if len(fields) > 10 else "",
                        "summary_value": Decimal("0.00"),
                        "line_number": line_number,
                        "summary_details": [],
                    }
                else:
                    current_d500 = None
            elif register == "C190":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["c190_total"] += summary_value
                    counts["c190_count"] += 1
                    if current_c100 is not None:
                        current_c100["summary_value"] = Decimal(current_c100.get("summary_value", Decimal("0.00"))) + summary_value
                        current_c100.setdefault("summary_details", []).append(
                            f"L{line_number} C190 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
            elif register == "C590":
                if selected_scope == "Ambos" or current_c_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["c590_total"] += summary_value
                    counts["c590_count"] += 1
                    if current_c500 is not None:
                        current_c500["summary_value"] = Decimal(current_c500.get("summary_value", Decimal("0.00"))) + summary_value
                        current_c500.setdefault("summary_details", []).append(
                            f"L{line_number} C590 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
            elif register == "D590":
                if selected_scope == "Ambos" or current_d_operation == selected_scope:
                    summary_value = compare_decimal_value(fields[5] if len(fields) > 5 else "")
                    totals["d590_total"] += summary_value
                    counts["d590_count"] += 1
                    if current_d500 is not None:
                        current_d500["summary_value"] = Decimal(current_d500.get("summary_value", Decimal("0.00"))) + summary_value
                        current_d500.setdefault("summary_details", []).append(
                            f"L{line_number} D590 CST={fields[2].strip() if len(fields) > 2 else ''} "
                            f"CFOP={fields[3].strip() if len(fields) > 3 else ''} "
                            f"ALIQ={fields[4].strip() if len(fields) > 4 else ''} "
                            f"VL_OPR={format_decimal_sped(summary_value)}"
                        )
    finish_c100_document()
    finish_c500_document()
    finish_d500_document()
    document_total = (totals["c100_total"] + totals["a100_total"]).quantize(Decimal("0.01"))
    fiscal_total = (totals["c190_total"] + totals["c590_total"] + totals["d590_total"]).quantize(Decimal("0.01"))
    result: dict[str, object] = {key: value.quantize(Decimal("0.01")) for key, value in totals.items()}
    result.update(counts)
    result["document_total"] = document_total
    result["fiscal_total"] = fiscal_total
    result["fiscal_document_difference"] = (fiscal_total - document_total).quantize(Decimal("0.01"))
    result["c100_c190_rows"] = c100_c190_rows
    result["c500_c590_rows"] = c500_c590_rows
    result["d500_d590_rows"] = d500_d590_rows
    return result

def compare_sped_with_xml_folder(
    sped_file: Path,
    xml_folder: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareXmlInvoice], list[CompareXmlInvoice], list[CompareSpedDocument], dict[str, object]]:
    def report_range(start: int, end: int, current: int, total: int, message: str) -> None:
        if progress_callback:
            span = max(end - start, 1)
            percent = start + int(span * current / max(total, 1))
            progress_callback(percent, 100, message)

    sped_documents = collect_compare_sped_documents(
        sped_file,
        operation_scope,
        lambda current, total, message: report_range(0, 35, current, total, message),
    )
    sped_keys = set(sped_documents.keys())
    sped_documents_by_number_series, sped_documents_by_number = build_compare_sped_number_indexes(sped_documents)
    company_tax_id = extract_company_tax_id_from_sped(sped_file)
    selected_scope = normalize_compare_operation_scope(operation_scope)
    xml_invoices, ignored_xml = collect_compare_xml_invoices(
        xml_folder,
        lambda current, total, message: report_range(35, 85, current, total, message),
    )
    cancellation_events = collect_xml_cancellation_events(sorted(xml_folder.rglob("*.xml")))
    xml_operation_by_key = {
        invoice.key: infer_compare_invoice_operation(invoice, company_tax_id)
        for invoice in xml_invoices
    }
    filtered_xml_invoices = [
        invoice
        for invoice in xml_invoices
        if selected_scope == "Ambos" or xml_operation_by_key.get(invoice.key) == selected_scope
    ]
    cancelled_xml_invoices = [
        invoice
        for invoice in filtered_xml_invoices
        if invoice.key in cancellation_events
    ]
    active_xml_invoices = [
        invoice
        for invoice in filtered_xml_invoices
        if invoice.key not in cancellation_events
    ]
    if progress_callback:
        progress_callback(90, 100, "Comparando chaves entre SPED e XMLs.")
    xml_keys = {invoice.key for invoice in active_xml_invoices}
    xml_keys_with_cancelled = xml_keys | {invoice.key for invoice in cancelled_xml_invoices}
    xml_matches_by_key: dict[str, CompareSpedDocument] = {}
    xml_matched_by_number_keys: set[str] = set()
    matched_rows: list[CompareInvestigationMatchedRow] = []
    present: list[CompareXmlInvoice] = []
    missing: list[CompareXmlInvoice] = []
    for invoice in active_xml_invoices:
        document, match_type = find_compare_sped_document_by_key_or_number(
            invoice,
            sped_documents,
            sped_documents_by_number_series,
            sped_documents_by_number,
        )
        if document is None:
            missing.append(invoice)
            continue
        present.append(invoice)
        xml_matches_by_key[invoice.key] = document
        xml_value = compare_decimal_value(invoice.total_value)
        sped_value = compare_decimal_value(document.total_value)
        matched_rows.append(
            CompareInvestigationMatchedRow(
                key=invoice.key,
                number=invoice.number,
                series=invoice.series,
                model=getattr(invoice, "model", ""),
                issuer_cnpj=invoice.issuer_cnpj,
                issue_date=invoice.issue_date,
                xml_value=format_decimal_sped(xml_value),
                sped_value=format_decimal_sped(sped_value),
                difference=format_decimal_sped((sped_value - xml_value).quantize(Decimal("0.01"))),
                match_type=match_type,
                xml_file=invoice.file_path,
                sped_key=document.key,
                reason=describe_compare_xml_sped_value_difference(invoice, document),
            )
        )
        if match_type == "number":
            xml_matched_by_number_keys.add(document.key)

    matched_sped_keys = {document.key for document in xml_matches_by_key.values()}
    sped_missing_xml = [
        document
        for key, document in sped_documents.items()
        if key not in xml_keys_with_cancelled and key not in matched_sped_keys
    ]
    present.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    missing.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    cancelled_xml_invoices.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    sped_missing_xml.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    operation_by_key = {key: document.operation_type for key, document in sped_documents.items()}
    operation_by_key.update({key: document.operation_type for key, document in xml_matches_by_key.items()})
    fiscal_basis = collect_compare_fiscal_basis_totals(sped_file, operation_scope)
    present_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in present), Decimal("0.00")).quantize(Decimal("0.01"))
    present_sped_value_total = sum((compare_decimal_value(row.sped_value) for row in matched_rows), Decimal("0.00")).quantize(Decimal("0.01"))
    missing_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in missing), Decimal("0.00")).quantize(Decimal("0.01"))
    sped_missing_xml_value_total = sum((compare_decimal_value(document.total_value) for document in sped_missing_xml), Decimal("0.00")).quantize(Decimal("0.01"))
    cancelled_xml_value_total = sum((compare_decimal_value(invoice.total_value) for invoice in cancelled_xml_invoices), Decimal("0.00")).quantize(Decimal("0.01"))
    return present, missing, sped_missing_xml, {
        "sped_keys": len(sped_keys),
        "xml_total": len(active_xml_invoices),
        "xml_total_lidos": len(xml_invoices),
        "xml_ignored": ignored_xml,
        "xml_cancelled": len(cancelled_xml_invoices),
        "xml_cancelled_value_total": cancelled_xml_value_total,
        "cancelled_xml_invoices": cancelled_xml_invoices,
        "cancellation_events": list(cancellation_events.values()),
        "present_total": len(matched_sped_keys),
        "source_present_total": len(present),
        "missing_total": len(missing),
        "sped_missing_xml_total": len(sped_missing_xml),
        "match_by_number_total": len(xml_matched_by_number_keys),
        "source_match_by_number_total": sum(1 for invoice in present if xml_matches_by_key.get(invoice.key) and xml_matches_by_key[invoice.key].key in xml_matched_by_number_keys),
        "operation_by_key": operation_by_key,
        "source_operation_by_key": xml_operation_by_key,
        "matched_rows": matched_rows,
        "fiscal_basis": fiscal_basis,
        "present_xml_value_total": present_xml_value_total,
        "present_sped_value_total": present_sped_value_total,
        "missing_xml_value_total": missing_xml_value_total,
        "sped_missing_xml_value_total": sped_missing_xml_value_total,
    }

def compare_sped_with_sheet(
    sped_file: Path,
    sheet_file: Path,
    operation_scope: str = "Ambos",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[CompareSheetInvoice], list[CompareSheetInvoice], dict[str, object]]:
    def report_range(start: int, end: int, current: int, total: int, message: str) -> None:
        if progress_callback:
            span = max(end - start, 1)
            percent = start + int(span * current / max(total, 1))
            progress_callback(percent, 100, message)

    selected_scope = normalize_compare_operation_scope(operation_scope)
    all_sped_documents = collect_compare_sped_documents(sped_file, "Ambos")
    sped_documents = collect_compare_sped_documents(
        sped_file,
        operation_scope,
        lambda current, total, message: report_range(0, 45, current, total, message),
    )
    sped_keys = set(sped_documents.keys())
    sheet_invoices, ignored_rows = collect_compare_sheet_invoices(
        sheet_file,
        lambda current, total, message: report_range(45, 85, current, total, message),
    )
    sped_documents_by_number_series, sped_documents_by_number = build_compare_sped_number_indexes(sped_documents)
    sheet_matches_by_key: dict[str, CompareSpedDocument] = {}
    sheet_matched_by_number_keys: set[str] = set()
    sheet_source_matched_by_number = 0

    def find_sped_document_for_sheet_invoice(invoice: CompareSheetInvoice) -> CompareSpedDocument | None:
        nonlocal sheet_source_matched_by_number
        document, match_type = find_compare_sped_document_by_key_or_number(
            invoice,
            sped_documents,
            sped_documents_by_number_series,
            sped_documents_by_number,
        )
        if document is not None:
            sheet_matches_by_key[invoice.key] = document
            if match_type == "number":
                sheet_matched_by_number_keys.add(document.key)
                sheet_source_matched_by_number += 1
        return document

    filtered_sheet_invoices = [
        invoice
        for invoice in sheet_invoices
        if (
            selected_scope == "Ambos"
            or invoice.operation_type == selected_scope
            or (
                not invoice.operation_type
                and (
                    invoice.key not in all_sped_documents
                    or all_sped_documents[invoice.key].operation_type == selected_scope
                )
            )
        )
    ]
    if progress_callback:
        progress_callback(90, 100, "Comparando chaves entre SPED e planilha.")
    present = [row for row in filtered_sheet_invoices if find_sped_document_for_sheet_invoice(row) is not None]
    missing = [row for row in filtered_sheet_invoices if row.key not in {invoice.key for invoice in present}]
    present.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    missing.sort(key=lambda item: (item.issue_date, item.number, item.series, item.key))
    matched_sped_keys = {document.key for document in sheet_matches_by_key.values()}
    operation_by_key = {key: document.operation_type for key, document in all_sped_documents.items()}
    operation_by_key.update({key: document.operation_type for key, document in sheet_matches_by_key.items()})
    return present, missing, {
        "sped_keys": len(sped_keys),
        "sheet_total": len(filtered_sheet_invoices),
        "sheet_total_lidas": len(sheet_invoices),
        "sheet_ignored": ignored_rows,
        "present_total": len(matched_sped_keys),
        "source_present_total": len(present),
        "missing_total": len(missing),
        "match_by_number_total": len(sheet_matched_by_number_keys),
        "source_match_by_number_total": sheet_source_matched_by_number,
        "operation_by_key": operation_by_key,
        "source_operation_by_key": {invoice.key: invoice.operation_type for invoice in sheet_invoices if invoice.operation_type},
    }
