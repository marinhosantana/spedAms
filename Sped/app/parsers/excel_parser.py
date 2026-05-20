from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def normalize_text(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value or "")
    no_accents = "".join(char for char in decomposed if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", no_accents).strip().upper()


def normalize_header(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", normalize_text(value))


def parse_excel_shared_strings(root: ET.Element | None) -> list[str]:
    if root is None:
        return []
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    values: list[str] = []
    for string_item in root.findall("main:si", namespace):
        parts = [node.text or "" for node in string_item.findall(".//main:t", namespace)]
        values.append("".join(parts))
    return values


def extract_excel_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//main:t", namespace)).strip()

    value = cell.findtext("main:v", default="", namespaces=namespace).strip()
    if cell_type == "s" and value:
        try:
            return shared_strings[int(value)].strip()
        except (ValueError, IndexError):
            return ""
    return value


def read_xlsx_sheet_rows(file_path: Path, sheet_name: str) -> list[list[str]]:
    # Le o XML interno do .xlsx sem depender de bibliotecas externas como openpyxl.
    namespace = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "docrel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with ZipFile(file_path, "r") as workbook:
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        workbook_rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        shared_strings_root = None
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_strings_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
        shared_strings = parse_excel_shared_strings(shared_strings_root)

        relationship_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in workbook_rels_root.findall("rel:Relationship", namespace)
        }

        sheet_target = None
        for sheet in workbook_root.findall("main:sheets/main:sheet", namespace):
            if sheet.attrib.get("name", "").strip() == sheet_name:
                relation_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                sheet_target = relationship_map.get(relation_id)
                break

        if not sheet_target:
            raise ValueError(f'A aba "{sheet_name}" nao foi encontrada na planilha.')

        sheet_path = f"xl/{sheet_target.lstrip('/')}"
        sheet_root = ET.fromstring(workbook.read(sheet_path))

    rows: list[list[str]] = []
    for row in sheet_root.findall("main:sheetData/main:row", namespace):
        cells: dict[int, str] = {}
        max_index = 0
        for cell in row.findall("main:c", namespace):
            reference = cell.attrib.get("r", "")
            column_letters = "".join(char for char in reference if char.isalpha())
            if not column_letters:
                continue
            column_index = 0
            for letter in column_letters:
                column_index = (column_index * 26) + (ord(letter.upper()) - 64)
            max_index = max(max_index, column_index)
            cells[column_index] = extract_excel_cell_value(cell, shared_strings)
        if max_index == 0:
            continue
        rows.append([cells.get(index, "").strip() for index in range(1, max_index + 1)])

    return rows


def read_xlsx_sheet(file_path: Path, sheet_name: str) -> list[dict[str, str]]:
    rows = read_xlsx_sheet_rows(file_path, sheet_name)
    if not rows:
        return []

    headers = [normalize_header(value) for value in rows[0]]
    records: list[dict[str, str]] = []
    for values in rows[1:]:
        if not any(str(value).strip() for value in values):
            continue
        record = {header: values[index] if index < len(values) else "" for index, header in enumerate(headers) if header}
        if record:
            records.append(record)
    return records


def get_first_xlsx_sheet_name(file_path: Path) -> str:
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(file_path, "r") as workbook:
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        first_sheet = workbook_root.find("main:sheets/main:sheet", namespace)
        if first_sheet is None:
            raise ValueError("A planilha nao possui abas.")
        return first_sheet.attrib.get("name", "").strip()


def read_filter_descriptions_file(file_path: Path) -> list[str]:
    if file_path.suffix.lower() == ".txt":
        lines = file_path.read_text(encoding="utf-8").splitlines()
        return [line.strip() for line in lines if line.strip()]

    if file_path.suffix.lower() == ".xlsx":
        sheet_name = get_first_xlsx_sheet_name(file_path)
        rows = read_xlsx_sheet(file_path, sheet_name)
        descriptions: list[str] = []
        for row in rows:
            for value in row.values():
                text = str(value).strip()
                if text:
                    descriptions.append(text)
                    break
        return descriptions

    raise ValueError("Arquivo de filtro deve ser .txt ou .xlsx.")


def parse_filter_values(raw_text: str) -> set[str]:
    values: set[str] = set()
    for chunk in re.split(r"[\r\n,;]+", raw_text or ""):
        text = str(chunk).strip()
        if text:
            values.add(text)
    return values


def has_active_item_filters(
    descriptions: list[str],
    cst_values: set[str],
    cfop_values: set[str],
) -> bool:
    return bool(descriptions or cst_values or cfop_values)
