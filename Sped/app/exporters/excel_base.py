from __future__ import annotations

import html
import re
import unicodedata
from decimal import Decimal


def excel_column_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xml_escape(value: object) -> str:
    return html.escape(str(value), quote=False)


def build_cell(value: object, cell_ref: str, style_id: int = 0) -> str:
    if value is None or value == "":
        return f'<c r="{cell_ref}" s="{style_id}"/>'

    if isinstance(value, Decimal):
        return f'<c r="{cell_ref}" s="{style_id}"><v>{value}</v></c>'

    if isinstance(value, (int, float)):
        return f'<c r="{cell_ref}" s="{style_id}"><v>{value}</v></c>'

    return (
        f'<c r="{cell_ref}" s="{style_id}" t="inlineStr">'
        f"<is><t>{xml_escape(value)}</t></is></c>"
    )


def build_sheet_xml(
    headers: list[str],
    rows: list[list[object]],
    header_style_ids: list[int] | None = None,
    row_style_ids: list[list[int]] | None = None,
) -> str:
    max_column = len(headers)
    max_row = len(rows) + 1
    sheet_rows = []

    header_cells = []
    for column_index, value in enumerate(headers, start=1):
        style_id = header_style_ids[column_index - 1] if header_style_ids and column_index - 1 < len(header_style_ids) else 1
        header_cells.append(build_cell(value, f"{excel_column_name(column_index)}1", style_id=style_id))
    sheet_rows.append(f'<row r="1">{"".join(header_cells)}</row>')

    for row_index, row in enumerate(rows, start=2):
        cells = []
        for column_index, value in enumerate(row, start=1):
            style_id = 0
            if row_style_ids and row_index - 2 < len(row_style_ids) and column_index - 1 < len(row_style_ids[row_index - 2]):
                style_id = row_style_ids[row_index - 2][column_index - 1]
            cells.append(build_cell(value, f"{excel_column_name(column_index)}{row_index}", style_id=style_id))
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    dimension = f"A1:{excel_column_name(max_column)}{max_row}"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{dimension}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    {''.join(sheet_rows)}
  </sheetData>
</worksheet>
"""


def sanitize_excel_sheet_name(name: str, fallback: str) -> str:
    cleaned = re.sub(r'[:\\/?*\[\]]', " ", str(name or "")).strip()
    if not cleaned:
        cleaned = fallback
    return cleaned[:31]


EXCEL_STYLE_DEFAULT = 0
EXCEL_STYLE_HEADER = 1
EXCEL_STYLE_CURRENCY = 2
EXCEL_STYLE_PERCENT = 3
EXCEL_STYLE_NUMBER = 4
EXCEL_STYLE_TOTAL_LABEL = 5
EXCEL_STYLE_TOTAL_CURRENCY = 6
EXCEL_STYLE_TOTAL_PERCENT = 7
EXCEL_STYLE_TOTAL_NUMBER = 8
EXCEL_STYLE_HEADER_BLUE = 9
EXCEL_STYLE_HEADER_YELLOW = 10
EXCEL_STYLE_BODY_BLUE = 11
EXCEL_STYLE_BODY_YELLOW = 12
EXCEL_STYLE_CURRENCY_BLUE = 13
EXCEL_STYLE_CURRENCY_YELLOW = 14
EXCEL_STYLE_PERCENT_BLUE = 15
EXCEL_STYLE_PERCENT_YELLOW = 16
EXCEL_STYLE_NUMBER_BLUE = 17
EXCEL_STYLE_NUMBER_YELLOW = 18


def normalize_excel_header_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return text.strip().lower()


def is_excel_numeric_value(value: object) -> bool:
    return isinstance(value, (Decimal, int, float)) and not isinstance(value, bool)


def decimal_from_excel_value(value: object) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    return Decimal("0")


def detect_excel_column_kind(header: str, values: list[object]) -> str:
    numeric_values = [value for value in values if is_excel_numeric_value(value)]
    if not numeric_values:
        return "text"

    normalized = normalize_excel_header_text(header)
    if (
        "%" in header
        or "percent" in normalized
        or "percentual" in normalized
        or "aliq" in normalized
        or "aliquota" in normalized
        or "ratio" in normalized
    ):
        return "percent"

    if any(
        token in normalized
        for token in (
            "valor",
            "base",
            "total",
            "venda",
            "credito",
            "debito",
            "desconto",
            "frete",
            "seguro",
            "outras",
            "outra",
            "operacao",
            "icms",
            "ipi",
            "pis",
            "cofins",
            "st",
            "bc",
            "perda",
        )
    ):
        return "currency"

    return "number"


def should_total_excel_column(column_kind: str) -> bool:
    return column_kind in {"currency", "number"}


def body_style_for_column_kind(column_kind: str, banded_fill: str | None = None) -> int:
    if banded_fill == "blue":
        if column_kind == "currency":
            return EXCEL_STYLE_CURRENCY_BLUE
        if column_kind == "percent":
            return EXCEL_STYLE_PERCENT_BLUE
        if column_kind == "number":
            return EXCEL_STYLE_NUMBER_BLUE
        return EXCEL_STYLE_BODY_BLUE
    if banded_fill == "yellow":
        if column_kind == "currency":
            return EXCEL_STYLE_CURRENCY_YELLOW
        if column_kind == "percent":
            return EXCEL_STYLE_PERCENT_YELLOW
        if column_kind == "number":
            return EXCEL_STYLE_NUMBER_YELLOW
        return EXCEL_STYLE_BODY_YELLOW
    if column_kind == "currency":
        return EXCEL_STYLE_CURRENCY
    if column_kind == "percent":
        return EXCEL_STYLE_PERCENT
    if column_kind == "number":
        return EXCEL_STYLE_NUMBER
    return EXCEL_STYLE_DEFAULT


def total_style_for_column_kind(column_kind: str) -> int:
    if column_kind == "currency":
        return EXCEL_STYLE_TOTAL_CURRENCY
    if column_kind == "percent":
        return EXCEL_STYLE_TOTAL_PERCENT
    if column_kind == "number":
        return EXCEL_STYLE_TOTAL_NUMBER
    return EXCEL_STYLE_TOTAL_LABEL


def build_sheet_rows_with_metadata(
    headers: list[str],
    rows: list[list[object]],
    banded_fills: list[str | None] | None = None,
    footer_rows: list[dict[str, object] | list[object]] | None = None,
    include_total: bool = True,
) -> tuple[list[list[object]], list[str], list[list[int]]]:
    column_values: list[list[object]] = []
    for column_index in range(len(headers)):
        values = [row[column_index] for row in rows if column_index < len(row)]
        column_values.append(values)

    column_kinds = [detect_excel_column_kind(header, values) for header, values in zip(headers, column_values)]
    prepared_rows = [list(row) for row in rows]
    row_style_ids: list[list[int]] = []

    for row in prepared_rows:
        styles: list[int] = []
        for column_index, value in enumerate(row):
            fill = banded_fills[column_index] if banded_fills and column_index < len(banded_fills) else None
            column_kind = column_kinds[column_index] if column_index < len(column_kinds) else "text"
            if is_excel_numeric_value(value):
                if column_kind == "percent":
                    row[column_index] = decimal_from_excel_value(value) / Decimal("100")
                styles.append(body_style_for_column_kind(column_kind, fill))
            else:
                styles.append(body_style_for_column_kind("text", fill))
        row_style_ids.append(styles)

    first_text_column_index = next((index for index, kind in enumerate(column_kinds) if kind == "text"), 0)
    total_row = [""] * len(headers)
    total_styles = [EXCEL_STYLE_TOTAL_LABEL] * len(headers)
    has_totals = False

    for column_index, values in enumerate(column_values):
        column_kind = column_kinds[column_index]
        if not should_total_excel_column(column_kind):
            continue
        numeric_values = [decimal_from_excel_value(value) for value in values if is_excel_numeric_value(value)]
        if not numeric_values:
            continue
        has_totals = True
        total_row[column_index] = sum(numeric_values, Decimal("0"))
        total_styles[column_index] = total_style_for_column_kind(column_kind)

    if include_total and has_totals:
        total_row[first_text_column_index] = "Total"
        prepared_rows.append(total_row)
        row_style_ids.append(total_styles)

    if footer_rows:
        for footer_row in footer_rows:
            row_type = "body"
            values = footer_row
            if isinstance(footer_row, dict):
                row_type = str(footer_row.get("row_type", "body")).strip().lower() or "body"
                values = list(footer_row.get("values", []))
            padded_values = list(values)[: len(headers)] + [""] * max(0, len(headers) - len(values))
            styles: list[int] = []
            for column_index, value in enumerate(padded_values):
                column_kind = column_kinds[column_index] if column_index < len(column_kinds) else "text"
                if row_type == "blank":
                    styles.append(EXCEL_STYLE_DEFAULT)
                elif row_type == "section":
                    styles.append(EXCEL_STYLE_HEADER)
                elif is_excel_numeric_value(value):
                    styles.append(body_style_for_column_kind(column_kind))
                else:
                    styles.append(EXCEL_STYLE_DEFAULT)
            prepared_rows.append(padded_values)
            row_style_ids.append(styles)

    return prepared_rows, column_kinds, row_style_ids


def build_excel_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="&quot;R$&quot; #,##0.00"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
    <numFmt numFmtId="166" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFBE7C6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEEF6FC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF4DF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="19">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""
