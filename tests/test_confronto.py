from __future__ import annotations

import os
import tempfile
from pathlib import Path

from app.exporters.excel_base import (
    EXCEL_STYLE_DEFAULT,
    EXCEL_STYLE_HEADER_BLUE,
    EXCEL_STYLE_ROW_DIV,
    EXCEL_STYLE_ROW_NAO,
    EXCEL_STYLE_ROW_OK,
    EXCEL_STYLE_STATUS_DIV,
    EXCEL_STYLE_STATUS_NAO,
    EXCEL_STYLE_STATUS_OK,
)
from app.exporters.workbook_exporter import write_simple_excel_workbook
from app.services.confronto_rules_builder import build_rules_from_confronto
from app.services.confronto_sped_cadastro import build_confronto_data


def test_build_rules_from_confronto_generates_rules_and_log() -> None:
    grouped_rows = [
        {
            "code": "PROD001",
            "status": "Divergencia: CST, CFOP",
            "descricao_sped": "BISCOITO",
            "cst_sped": "060",
            "cfop_sped": "1401",
            "cst_cad": "000",
            "cfop_cad": "1102",
            "fornecedor": "FORNEC TESTE",
            "cnpj_fornecedor_cad": "12345678000199",
            "aliq_sped": "18",
            "aliq_cad": "18",
            "periodo": "05/2026",
        },
        {
            "code": "PROD002",
            "status": "Divergencia: CFOP",
            "descricao_sped": "REFRIGERANTE",
            "cst_sped": "060",
            "cfop_sped": "1401",
            "cst_cad": "060",
            "cfop_cad": "1403",
            "fornecedor": "COCA COLA",
            "cnpj_fornecedor_cad": "",
            "aliq_sped": "0",
            "aliq_cad": "",
            "periodo": "05/2026",
        },
        {
            "code": "PROD003",
            "status": "Nao Cadastrado",
            "descricao_sped": "PAO FRANCES",
            "cst_sped": "041",
            "cfop_sped": "1401",
            "cst_cad": "",
            "cfop_cad": "",
            "fornecedor": "PADARIA",
            "cnpj_fornecedor_cad": "",
            "periodo": "05/2026",
        },
        {
            "code": "PROD004",
            "status": "OK",
            "descricao_sped": "AGUA MINERAL",
            "cst_sped": "041",
            "cfop_sped": "1403",
            "cst_cad": "041",
            "cfop_cad": "1403",
            "fornecedor": "AGUA LTDA",
            "cnpj_fornecedor_cad": "98765432000100",
            "periodo": "05/2026",
        },
    ]

    rules, log = build_rules_from_confronto(grouped_rows, "Entrada")

    assert len(rules) == 2
    assert rules[0]["new_cst"] == "000"
    assert rules[0]["new_cfop"] == "1102"
    assert rules[0]["document_tax_id"] == "12345678000199"
    assert "document_tax_id" not in rules[1]
    assert len(log) == 3
    assert log[2]["tipo"] == "NAO ENCONTRADO"


def test_build_confronto_data_includes_supplier_tax_id() -> None:
    catalog = [
        {
            "codigo_empresa": "PROD001",
            "cst_icms": "000",
            "cfop_entrada": "1102",
            "aliquota_icms": "18",
            "descricao": "BISCOITO CAD",
            "ncm": "1905",
            "fornecedor_nome": "FORNEC SA",
            "fornecedor_cnpj": "12345678000199",
        },
        {
            "codigo_empresa": "PROD002",
            "cst_icms": "060",
            "cfop_entrada": "1401",
            "aliquota_icms": "0",
            "descricao": "REFRI CAD",
            "ncm": "2202",
            "fornecedor_nome": "COCA LTDA",
            "fornecedor_cnpj": "",
        },
    ]
    sped_rows = [
        {
            "code": "PROD001",
            "description": "BISCOITO",
            "period": "05/2026",
            "cst_icms": "060",
            "cfop": "1401",
            "display_icms_rate": "18",
            "ncm": "1905",
            "suppliers": "FORNEC TESTE",
            "sale_value": "100",
            "base_icms": "100",
            "icms_value": "18",
            "launch_details": [],
        },
        {
            "code": "PROD003",
            "description": "PAO",
            "period": "05/2026",
            "cst_icms": "041",
            "cfop": "1401",
            "display_icms_rate": "0",
            "ncm": "",
            "suppliers": "PADARIA",
            "sale_value": "50",
            "base_icms": "0",
            "icms_value": "0",
            "launch_details": [],
        },
    ]

    grouped_rows, _detail_rows = build_confronto_data(sped_rows, catalog, "Entrada")

    assert grouped_rows[0]["status"].startswith("Divergencia")
    assert grouped_rows[0]["cnpj_fornecedor_cad"] == "12345678000199"
    assert grouped_rows[1]["status"] == "Nao Cadastrado"


def test_write_confronto_workbook_with_status_styles() -> None:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        temp_path = Path(temp_file.name)

    try:
        write_simple_excel_workbook(
            temp_path,
            [
                (
                    "Resumo",
                    ["Campo", "Resultado"],
                    [["Empresa", "Teste SA"], ["Total", "10"], ["Divergentes", "8"]],
                    {
                        "row_style_ids": [
                            [EXCEL_STYLE_HEADER_BLUE, EXCEL_STYLE_HEADER_BLUE],
                            [EXCEL_STYLE_DEFAULT, EXCEL_STYLE_DEFAULT],
                            [EXCEL_STYLE_DEFAULT, EXCEL_STYLE_DEFAULT],
                        ],
                        "include_total": False,
                    },
                ),
                (
                    "Divergencias (8)",
                    ["Tipo", "Codigo", "Descricao"],
                    [["CST e CFOP", "PROD001", "BISCOITO"], ["CFOP", "PROD002", "REFRI"]],
                    {
                        "row_style_ids": [
                            [EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV, EXCEL_STYLE_ROW_DIV],
                            [EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV, EXCEL_STYLE_ROW_DIV],
                        ],
                        "include_total": False,
                    },
                ),
                (
                    "Nao Cadastrados (1)",
                    ["Codigo", "Descricao", "Observacao"],
                    [["PROD003", "PAO", "Nao encontrado"]],
                    {"row_style_ids": [[EXCEL_STYLE_STATUS_NAO, EXCEL_STYLE_ROW_NAO, EXCEL_STYLE_ROW_NAO]], "include_total": False},
                ),
                (
                    "OK (1)",
                    ["Codigo", "Descricao"],
                    [["PROD004", "AGUA"]],
                    {"row_style_ids": [[EXCEL_STYLE_STATUS_OK, EXCEL_STYLE_ROW_OK]], "include_total": False},
                ),
            ],
        )

        assert os.path.getsize(temp_path) > 3000
    finally:
        temp_path.unlink(missing_ok=True)
