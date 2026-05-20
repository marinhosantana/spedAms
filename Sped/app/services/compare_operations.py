from __future__ import annotations

from decimal import Decimal

from app.models import CompareXmlInvoice, CompareXmlItem
from app.parsers.compare_xml import compare_date_to_sped
from app.parsers.sped_parser import normalize_document_key
from app.services.tax_rules import normalize_text


def classify_xml_item_operation(invoice: CompareXmlInvoice, item: CompareXmlItem, company_tax_id: str = "") -> str:
    company_key = normalize_document_key(company_tax_id)
    issuer_key = normalize_document_key(invoice.issuer_cnpj)
    recipient_key = normalize_document_key(invoice.recipient_cnpj)
    model = str(invoice.model or "").strip()
    if model == "65":
        return "Saida"
    if model.upper().startswith("NFS"):
        if company_key:
            if issuer_key == company_key:
                return "Saida"
            if recipient_key == company_key:
                return "Entrada"
        return "Entrada"
    cfop = str(item.cfop or "").strip()
    is_outgoing_cfop = cfop.startswith(("5", "6"))

    if company_key:
        if issuer_key == company_key:
            return "Saida" if is_outgoing_cfop else "Entrada"
        if recipient_key == company_key:
            return "Entrada"
    return "Saida" if is_outgoing_cfop else "Entrada"


def normalize_compare_operation_scope(value: object) -> str:
    normalized = normalize_text(str(value))
    if normalized in {"E", "ENT", "ENTR"} or normalized.startswith("ENTRADA"):
        return "Entrada"
    if normalized in {"S", "SAI"} or normalized.startswith("SAIDA"):
        return "Saida"
    return "Ambos"


def filter_xml_summary_rows_by_scope(rows: list[dict[str, object]], operation_scope: object) -> list[dict[str, object]]:
    selected_scope = normalize_compare_operation_scope(operation_scope)
    if selected_scope == "Ambos":
        return list(rows)
    return [
        row
        for row in rows
        if normalize_compare_operation_scope(row.get("operation_type", "")) == selected_scope
    ]


def infer_compare_invoice_operation(invoice: CompareXmlInvoice, company_tax_id: str = "") -> str:
    operations = {
        classify_xml_item_operation(invoice, item, company_tax_id)
        for item in invoice.items
    }
    operations.discard("")
    if len(operations) == 1:
        return next(iter(operations))
    if len(operations) > 1:
        return "Entrada/Saida"
    return classify_xml_item_operation(
        invoice,
        CompareXmlItem(
            item_no="",
            code="",
            description="",
            unit="",
            ncm="",
            cfop="",
            quantity=0.0,
            value=0.0,
            discount=0.0,
            cst_icms="",
            vl_bc_icms=0.0,
            aliq_icms=0.0,
            vl_icms=0.0,
            vl_bc_icms_st=0.0,
            aliq_icms_st=0.0,
            vl_icms_st=0.0,
            vl_ipi=0.0,
            cst_pis="",
            vl_bc_pis=0.0,
            aliq_pis=0.0,
            vl_pis=0.0,
            cst_cofins="",
            vl_bc_cofins=0.0,
            aliq_cofins=0.0,
            vl_cofins=0.0,
        ),
        company_tax_id,
    )


def build_compare_invoice_note_snapshot(invoice: CompareXmlInvoice) -> dict[str, object]:
    items: list[dict[str, object]] = []
    for item in invoice.items:
        sale_value = Decimal(str(item.value or 0)).quantize(Decimal("0.01"))
        ipi_value = Decimal(str(item.vl_ipi or 0)).quantize(Decimal("0.01"))
        icms_st_value = Decimal(str(item.vl_icms_st or 0)).quantize(Decimal("0.01"))
        total_operation_value = (sale_value + ipi_value + icms_st_value).quantize(Decimal("0.01"))
        items.append(
            {
                "item_number": item.item_no,
                "code": item.code,
                "description": item.description,
                "ncm": item.ncm,
                "cest": "",
                "cfop": item.cfop,
                "cst_icms": item.cst_icms,
                "icms_rate": Decimal(str(item.aliq_icms or 0)).quantize(Decimal("0.01")),
                "quantity": Decimal(str(item.quantity or 0)),
                "sale_value": sale_value,
                "total_operation_value": total_operation_value,
                "discount_value": Decimal(str(item.discount or 0)).quantize(Decimal("0.01")),
                "ipi_value": ipi_value,
                "base_icms": Decimal(str(item.vl_bc_icms or 0)).quantize(Decimal("0.01")),
                "icms_value": Decimal(str(item.vl_icms or 0)).quantize(Decimal("0.01")),
                "icms_st_value": icms_st_value,
                "source_register": "XML",
            }
        )
    return {
        "period": "",
        "document_date": compare_date_to_sped(invoice.issue_date) or invoice.issue_date,
        "document_number": invoice.number,
        "document_series": invoice.series,
        "document_model": invoice.model,
        "document_key": invoice.key,
        "participant_code": "",
        "participant_name": invoice.issuer_name,
        "participant_tax_id": normalize_document_key(invoice.issuer_cnpj),
        "items": items,
    }
