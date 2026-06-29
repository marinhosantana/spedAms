from __future__ import annotations

import csv
import datetime as dt
import io
import sys
import zipfile as _zipmod
from decimal import Decimal
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

# ── Logo DZ Consultoria ───────────────────────────────────────────────────────
def _get_logo_path() -> Path | None:
    if getattr(sys, "frozen", False):
        p = Path(sys.executable).parent / "assets" / "logo.png"
    else:
        p = Path(__file__).parent.parent.parent / "assets" / "logo.png"
    return p if p.exists() else None

_LOGO_PATH_XLSX: Path | None = _get_logo_path()

# Dimensões da logo na planilha: 5cm × 3.63cm → EMU (1 cm = 360000 EMU)
_LOGO_EMU_CX = 1800000
_LOGO_EMU_CY = 1308000


def _inject_logo_xlsx(xlsx_path: Path) -> None:
    """Insere a logo DZ na primeira planilha do XLSX gerado, sem alterar o conteúdo."""
    if not _LOGO_PATH_XLSX:
        return
    try:
        logo_bytes = _LOGO_PATH_XLSX.read_bytes()

        with _zipmod.ZipFile(xlsx_path, "r") as zin:
            entries: dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}

        sheet_key = "xl/worksheets/sheet1.xml"
        if sheet_key not in entries:
            return

        # Adiciona xmlns:r no <worksheet> se ainda não existir
        sheet_xml = entries[sheet_key].decode("utf-8")
        if "xmlns:r=" not in sheet_xml:
            sheet_xml = sheet_xml.replace(
                "<worksheet ",
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                1,
            )
        # Insere referência ao drawing antes de </worksheet>
        if "<drawing" not in sheet_xml:
            sheet_xml = sheet_xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>', 1)
        entries[sheet_key] = sheet_xml.encode("utf-8")

        # xl/worksheets/_rels/sheet1.xml.rels
        sheet_rels_key = "xl/worksheets/_rels/sheet1.xml.rels"
        if sheet_rels_key in entries:
            rels_xml = entries[sheet_rels_key].decode("utf-8")
            new_rel = (
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"'
                ' Target="../drawings/drawing1.xml"/>'
            )
            rels_xml = rels_xml.replace("</Relationships>", new_rel + "</Relationships>", 1)
            entries[sheet_rels_key] = rels_xml.encode("utf-8")
        else:
            entries[sheet_rels_key] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"'
                ' Target="../drawings/drawing1.xml"/>'
                "</Relationships>"
            ).encode("utf-8")

        # [Content_Types].xml — adiciona PNG e drawing
        ct_key = "[Content_Types].xml"
        if ct_key in entries:
            ct_xml = entries[ct_key].decode("utf-8")
            if 'Extension="png"' not in ct_xml:
                ct_xml = ct_xml.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>', 1)
            if "xl/drawings/drawing1.xml" not in ct_xml:
                ct_xml = ct_xml.replace(
                    "</Types>",
                    '<Override PartName="/xl/drawings/drawing1.xml"'
                    ' ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>',
                    1,
                )
            entries[ct_key] = ct_xml.encode("utf-8")

        # xl/drawings/drawing1.xml
        drawing_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<xdr:oneCellAnchor>"
            "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>"
            "<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:ext cx=\"{_LOGO_EMU_CX}\" cy=\"{_LOGO_EMU_CY}\"/>"
            "<xdr:pic>"
            '<xdr:nvPicPr><xdr:cNvPr id="2" name="Logo DZ"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            '<xdr:blipFill><a:blip r:embed="rId1"/>'
            "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
            "<xdr:spPr>"
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_LOGO_EMU_CX}" cy="{_LOGO_EMU_CY}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            "</xdr:spPr>"
            "</xdr:pic>"
            "<xdr:clientData/>"
            "</xdr:oneCellAnchor>"
            "</xdr:wsDr>"
        )
        drawing_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
            ' Target="../media/logo.png"/>'
            "</Relationships>"
        )
        entries["xl/media/logo.png"] = logo_bytes
        entries["xl/drawings/drawing1.xml"] = drawing_xml.encode("utf-8")
        entries["xl/drawings/_rels/drawing1.xml.rels"] = drawing_rels.encode("utf-8")

        buf = io.BytesIO()
        with _zipmod.ZipFile(buf, "w", compression=ZIP_DEFLATED) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        xlsx_path.write_bytes(buf.getvalue())

    except Exception:
        pass  # logo é opcional; não deve quebrar a exportação


from app.exporters.excel_base import (
    EXCEL_STYLE_HEADER,
    EXCEL_STYLE_HEADER_BLUE,
    EXCEL_STYLE_HEADER_YELLOW,
    EXCEL_STYLE_TOTAL_CURRENCY,
    EXCEL_STYLE_TOTAL_LABEL,
    EXCEL_STYLE_TOTAL_NUMBER,
    EXCEL_STYLE_TOTAL_PERCENT,
    body_style_for_column_kind,
    build_excel_styles_xml,
    build_sheet_rows_with_metadata,
    build_sheet_xml,
    detect_excel_column_kind,
    sanitize_excel_sheet_name,
    xml_escape,
)


def format_decimal_sped(value: Decimal) -> str:
    normalized = value.quantize(Decimal("0.01"))
    return format(normalized, "f").replace(".", ",")


def write_simple_excel_workbook(
    output_path: Path,
    sheets: list[tuple[str, list[str], list[list[object]]] | tuple[str, list[str], list[list[object]], dict[str, object]]],
) -> None:
    workbook_sheets: list[tuple[str, list[str], list[list[object]], list[list[int]]]] = []
    used_names: set[str] = set()
    for index, sheet_spec in enumerate(sheets, start=1):
        if len(sheet_spec) == 4:
            sheet_name, headers, rows, options = sheet_spec
        else:
            sheet_name, headers, rows = sheet_spec
            options = {}
        base_name = sanitize_excel_sheet_name(sheet_name, f"Planilha {index}")
        candidate = base_name
        suffix = 2
        while candidate in used_names:
            trimmed = base_name[: max(0, 31 - len(str(suffix)) - 1)]
            candidate = f"{trimmed} {suffix}"
            suffix += 1
        used_names.add(candidate)
        prepared_rows, _column_kinds, row_style_ids = build_sheet_rows_with_metadata(
            headers,
            rows,
            footer_rows=list(options.get("footer_rows", [])),
            include_total=bool(options.get("include_total", True)),
        )
        if "row_style_ids" in options:
            row_style_ids = list(options["row_style_ids"])
        workbook_sheets.append((candidate, headers, prepared_rows, row_style_ids))

    sheet_entries = []
    rel_entries = []
    content_entries = []
    for index, (sheet_name, _, _, _) in enumerate(workbook_sheets, start=1):
        sheet_entries.append(f'    <sheet name="{xml_escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>')
        rel_entries.append(
            f'  <Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        )
        content_entries.append(
            f'  <Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{chr(10).join(sheet_entries)}
  </sheets>
</workbook>
"""
    workbook_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{chr(10).join(rel_entries)}
  <Relationship Id="rId{len(workbook_sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{chr(10).join(content_entries)}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    styles_xml = build_excel_styles_xml()
    created_at = dt.datetime.now().replace(microsecond=0).isoformat()
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>
</cp:coreProperties>
"""
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>
"""

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", root_rels)
        workbook.writestr("docProps/core.xml", core_xml)
        workbook.writestr("docProps/app.xml", app_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook.writestr("xl/styles.xml", styles_xml)
        for index, (_, headers, rows, row_style_ids) in enumerate(workbook_sheets, start=1):
            workbook.writestr(
                f"xl/worksheets/sheet{index}.xml",
                build_sheet_xml(headers, rows, row_style_ids=row_style_ids),
            )

    _inject_logo_xlsx(output_path)


def build_month_color_style_maps(
    headers: list[str],
    rows: list[list[object]],
    periods: list[str],
    fixed_column_count: int = 6,
    monthly_block_size: int = 4,
) -> tuple[list[int], list[int]]:
    column_values = [[row[column_index] for row in rows if column_index < len(row)] for column_index in range(len(headers))]
    column_kinds = [detect_excel_column_kind(header, values) for header, values in zip(headers, column_values)]
    header_style_ids = [EXCEL_STYLE_HEADER] * len(headers)
    body_style_ids = [body_style_for_column_kind(column_kinds[index]) for index in range(len(headers))]
    for period_index, _period in enumerate(periods):
        header_style = EXCEL_STYLE_HEADER_BLUE if period_index % 2 == 0 else EXCEL_STYLE_HEADER_YELLOW
        fill_name = "blue" if period_index % 2 == 0 else "yellow"
        start = fixed_column_count + (period_index * monthly_block_size)
        end = min(start + monthly_block_size, len(headers))
        for column_index in range(start, end):
            header_style_ids[column_index] = header_style
            body_style_ids[column_index] = body_style_for_column_kind(column_kinds[column_index], fill_name)
    return header_style_ids, body_style_ids


def write_monthly_colored_excel_workbook(
    output_path: Path,
    sheet_name: str,
    headers: list[str],
    rows: list[list[object]],
    periods: list[str],
) -> None:
    write_monthly_colored_excel_workbook_with_sheets(
        output_path,
        [(sheet_name, headers, rows, periods)],
    )


def write_monthly_colored_excel_workbook_with_sheets(
    output_path: Path,
    sheets: list[tuple[str, list[str], list[list[object]], list[str]]],
) -> None:
    normalized_sheets: list[tuple[str, list[str], list[list[object]], list[str], list[int], list[list[int]]]] = []
    used_names: set[str] = set()
    for index, (sheet_name, headers, rows, periods) in enumerate(sheets, start=1):
        base_name = sanitize_excel_sheet_name(sheet_name, f"Planilha {index}")
        candidate = base_name
        suffix = 2
        while candidate in used_names:
            trimmed = base_name[: max(0, 31 - len(str(suffix)) - 1)]
            candidate = f"{trimmed} {suffix}"
            suffix += 1
        used_names.add(candidate)
        prepared_rows, _column_kinds, row_style_ids = build_sheet_rows_with_metadata(headers, rows)
        header_style_ids, body_style_ids = build_month_color_style_maps(headers, prepared_rows, periods)
        row_style_ids = [
            [
                body_style_ids[column_index]
                if row_style_ids[row_index][column_index] != EXCEL_STYLE_TOTAL_LABEL
                and row_style_ids[row_index][column_index] not in {EXCEL_STYLE_TOTAL_CURRENCY, EXCEL_STYLE_TOTAL_PERCENT, EXCEL_STYLE_TOTAL_NUMBER}
                else row_style_ids[row_index][column_index]
                for column_index in range(len(row_style_ids[row_index]))
            ]
            for row_index in range(len(row_style_ids))
        ]
        normalized_sheets.append((candidate, headers, prepared_rows, periods, header_style_ids, row_style_ids))

    sheet_entries = []
    rel_entries = []
    content_entries = []
    for index, (sheet_name, _, _, _, _, _) in enumerate(normalized_sheets, start=1):
        sheet_entries.append(f'    <sheet name="{xml_escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>')
        rel_entries.append(
            f'  <Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        )
        content_entries.append(
            f'  <Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{chr(10).join(sheet_entries)}
  </sheets>
</workbook>
"""
    workbook_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{chr(10).join(rel_entries)}
  <Relationship Id="rId{len(normalized_sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{chr(10).join(content_entries)}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    styles_xml = build_excel_styles_xml()
    created_at = dt.datetime.now().replace(microsecond=0).isoformat()
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>
</cp:coreProperties>
"""
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>
"""

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", root_rels)
        workbook.writestr("docProps/core.xml", core_xml)
        workbook.writestr("docProps/app.xml", app_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook.writestr("xl/styles.xml", styles_xml)
        for index, (_, headers, rows, _, header_style_ids, row_style_ids) in enumerate(normalized_sheets, start=1):
            workbook.writestr(
                f"xl/worksheets/sheet{index}.xml",
                build_sheet_xml(headers, rows, header_style_ids, row_style_ids),
            )

    _inject_logo_xlsx(output_path)


def write_simple_csv_file(
    output_path: Path,
    headers: list[str],
    rows: list[list[object]],
) -> None:
    with output_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file, delimiter=";")
        writer.writerow(headers)
        for row in rows:
            serialized_row: list[str] = []
            for value in row:
                if isinstance(value, Decimal):
                    serialized_row.append(format_decimal_sped(value))
                else:
                    serialized_row.append(str(value or ""))
            writer.writerow(serialized_row)


def serialize_value_for_clipboard(value: object) -> str:
    if isinstance(value, Decimal):
        return format_decimal_sped(value)
    return str(value or "")
