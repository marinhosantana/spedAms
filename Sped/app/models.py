from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal


@dataclass
class ProductRecord:
    code: str
    description: str
    ncm: str
    cst_icms: str
    icms_rate: Decimal | None
    cest: str = ""


@dataclass
class CompareXmlItem:
    item_no: str
    code: str
    ean: str
    description: str
    unit: str
    ncm: str
    cest: str
    c_classtrib: str
    cfop: str
    quantity: float
    value: float
    discount: float
    cst_icms: str
    vl_bc_icms: float
    aliq_icms: float
    vl_icms: float
    vl_bc_icms_st: float
    aliq_icms_st: float
    vl_icms_st: float
    vl_ipi: float
    cst_pis: str
    vl_bc_pis: float
    aliq_pis: float
    vl_pis: float
    cst_cofins: str
    vl_bc_cofins: float
    aliq_cofins: float
    vl_cofins: float
    icms_code_source: str = ""


@dataclass
class CompareXmlInvoice:
    key: str
    number: str
    series: str
    issue_date: str
    issuer_cnpj: str
    issuer_name: str
    issuer_ie: str
    issuer_city_code: str
    recipient_cnpj: str
    recipient_name: str
    recipient_ie: str
    total_value: str
    date_doc: str
    date_entry: str
    model: str
    freight_mode: str
    total_doc: float
    total_discount: float
    total_products: float
    total_freight: float
    total_insurance: float
    total_other: float
    total_bc_icms: float
    total_icms: float
    total_bc_st: float
    total_st: float
    total_ipi: float
    total_pis: float
    total_cofins: float
    file_path: str
    items: list[CompareXmlItem]
    issuer_address: str = ""
    issuer_number: str = ""
    issuer_complement: str = ""
    issuer_district: str = ""


@dataclass
class CompareSheetInvoice:
    key: str
    number: str
    series: str
    issue_date: str
    issuer_cnpj: str
    total_value: str
    file_path: str
    operation_type: str = ""


@dataclass
class CompareSpedDocument:
    key: str
    number: str
    series: str
    issue_date: str
    issuer_cnpj: str
    total_value: str
    file_path: str
    operation_type: str
    model: str = ""
    generated_nfse_service: bool = False


@dataclass
class CompareInvestigationMatchedRow:
    key: str
    number: str
    series: str
    model: str
    issuer_cnpj: str
    issue_date: str
    xml_value: str
    sped_value: str
    difference: str
    match_type: str
    xml_file: str
    sped_key: str
    reason: str = ""



