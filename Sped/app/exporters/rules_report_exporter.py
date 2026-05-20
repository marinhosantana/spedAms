from __future__ import annotations

import datetime as dt
import html
from decimal import Decimal
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from app.services.tax_rules import normalize_operation_type



def format_rule_decimal(value: object) -> str:
    if not isinstance(value, Decimal):
        return str(value or "").strip()
    normalized = value.quantize(Decimal("0.01"))
    text = format(normalized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", ",")


def describe_operation_type(value: object) -> str:
    normalized = normalize_operation_type(str(value or "")).strip()
    return normalized or "Qualquer"


def describe_codes(value: object) -> str:
    if not isinstance(value, set):
        return ""
    codes = sorted(str(code).strip() for code in value if str(code).strip())
    return ", ".join(codes)


def describe_rule_conditions(rule: dict[str, object]) -> str:
    parts: list[str] = []
    operation_type = str(rule.get("operation_type", "")).strip()
    if operation_type:
        parts.append(f"Tipo {describe_operation_type(operation_type)}")
    cst_icms = str(rule.get("cst_icms", "")).strip()
    if cst_icms:
        parts.append(f"CST {cst_icms}")
    cfop = str(rule.get("cfop", "")).strip()
    if cfop:
        parts.append(f"CFOP {cfop}")
    match_rate = rule.get("match_rate")
    if isinstance(match_rate, Decimal):
        parts.append(f"Aliquota {format_rule_decimal(match_rate)}")
    match_rates = rule.get("match_rates")
    if isinstance(match_rates, set) and match_rates:
        formatted = ", ".join(sorted(format_rule_decimal(rate) for rate in match_rates))
        parts.append(f"Aliquotas {formatted}")
    match_sale_value = rule.get("match_sale_value")
    if isinstance(match_sale_value, Decimal):
        parts.append(f"Valor da operacao {format_rule_decimal(match_sale_value)}")
    match_base_icms = rule.get("match_base_icms")
    if isinstance(match_base_icms, Decimal):
        parts.append(f"Base ICMS {format_rule_decimal(match_base_icms)}")
    match_codes = describe_codes(rule.get("match_codes"))
    if match_codes:
        parts.append(f"Codigos {match_codes}")
    return " | ".join(parts) if parts else "Sem filtro especifico"


def describe_rule_actions(rule: dict[str, object]) -> str:
    parts: list[str] = []
    new_cst = str(rule.get("new_cst", "")).strip()
    if new_cst:
        parts.append(f"alterar CST para {new_cst}")
    new_cfop = str(rule.get("new_cfop", "")).strip()
    if new_cfop:
        parts.append(f"alterar CFOP para {new_cfop}")
    force_rate = rule.get("force_rate")
    if isinstance(force_rate, Decimal):
        parts.append(f"definir aliquota ICMS em {format_rule_decimal(force_rate)}")
    set_base_icms = rule.get("set_base_icms")
    if isinstance(set_base_icms, Decimal):
        parts.append(f"definir base ICMS em {format_rule_decimal(set_base_icms)}")
    set_icms_value = rule.get("set_icms_value")
    if isinstance(set_icms_value, Decimal):
        parts.append(f"definir valor do ICMS em {format_rule_decimal(set_icms_value)}")
    if rule.get("use_sale_value_as_base"):
        parts.append("usar o valor da operacao como base do ICMS")
    set_sale_value = rule.get("set_sale_value")
    if isinstance(set_sale_value, Decimal):
        parts.append(f"definir valor da operacao em {format_rule_decimal(set_sale_value)}")
    sale_value_factor = rule.get("sale_value_factor")
    if isinstance(sale_value_factor, Decimal):
        parts.append(f"aplicar fator {format_rule_decimal(sale_value_factor)} ao valor da operacao")
    icms_per_quantity = rule.get("icms_per_quantity")
    if isinstance(icms_per_quantity, Decimal):
        parts.append(f"calcular ICMS por quantidade em {format_rule_decimal(icms_per_quantity)} por unidade")
    if rule.get("recalculate_icms_value"):
        parts.append("recalcular o valor do ICMS")
    if rule.get("zero_icms"):
        parts.append("zerar base, aliquota e valor do ICMS")
    if rule.get("preserve_icms_value"):
        parts.append("preservar o valor final do ICMS apos o calculo especifico")
    return "; ".join(parts) if parts else "Sem acao de alteracao"


def describe_rule_notes(rule: dict[str, object]) -> str:
    notes: list[str] = ["Nao altera o valor do IPI."]
    split_into = rule.get("split_into")
    if isinstance(split_into, list) and split_into:
        notes.append(f"Desdobra o item em {len(split_into)} lancamento(s).")
    return " ".join(notes)


def build_rule_report_entries(rule: dict[str, object], prefix: str = "") -> list[str]:
    summary_line = f"{prefix}{describe_rule_conditions(rule)} -> {describe_rule_actions(rule)}. {describe_rule_notes(rule)}"
    lines = [summary_line]
    split_into = rule.get("split_into")
    if isinstance(split_into, list) and split_into:
        for split_index, split_rule in enumerate(split_into, start=1):
            if isinstance(split_rule, dict):
                lines.extend(build_rule_report_entries(split_rule, prefix=f"{prefix}Desdobramento {split_index}: "))
    return lines


def make_word_paragraph(text: str, style: str | None = None) -> str:
    escaped = html.escape(str(text or ""))
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return (
        "<w:p>"
        f"{style_xml}"
        f"<w:r><w:t xml:space=\"preserve\">{escaped}</w:t></w:r>"
        "</w:p>"
    )


def write_rules_report_docx(
    output_path: Path,
    report_title: str,
    sections: list[dict[str, object]],
) -> None:
    body_parts: list[str] = [
        make_word_paragraph(report_title, "Heading1"),
        make_word_paragraph(f"Gerado em: {dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"),
        make_word_paragraph("Observacao geral: as regras documentadas neste relatorio nao alteram o valor do IPI."),
    ]

    for section in sections:
        title = str(section.get("title", "")).strip()
        lines = [str(line).strip() for line in section.get("lines", []) if str(line).strip()]
        if not title or not lines:
            continue
        body_parts.append(make_word_paragraph(""))
        body_parts.append(make_word_paragraph(title, "Heading2"))
        for line in lines:
            body_parts.append(make_word_paragraph(line))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14">'
        "<w:body>"
        + "".join(body_parts)
        + (
            "<w:sectPr>"
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="708" w:footer="708" w:gutter="0"/>'
            "</w:sectPr>"
            "</w:body></w:document>"
        )
    )

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
</w:styles>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Word</Application>
</Properties>
"""
    created_at = dt.datetime.now().replace(microsecond=0).isoformat()
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{html.escape(report_title)}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>
</cp:coreProperties>
"""

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("docProps/app.xml", app_xml)
        docx.writestr("docProps/core.xml", core_xml)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/_rels/document.xml.rels", document_rels)
