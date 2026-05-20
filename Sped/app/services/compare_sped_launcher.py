from __future__ import annotations

import re
from pathlib import Path

from app.models import CompareXmlInvoice, CompareXmlItem
from app.parsers.compare_xml import (
    compare_clean,
    compare_format_float,
    compare_is_nfse_invoice,
    compare_sanitize,
)


def normalize_compare_document_number(value: object) -> str:
    text = compare_clean(value).upper()
    digits = re.sub(r"\D+", "", text)
    if digits:
        return digits.lstrip("0") or "0"
    return re.sub(r"[^A-Z0-9]+", "", text).lstrip("0")


def compare_register(line: str) -> str:
    parts = line.rstrip("\n").split("|")
    return parts[1] if len(parts) > 1 else ""


def compare_line(fields: list[str]) -> str:
    text = "|".join(fields)
    return text if text.endswith("|") else f"{text}|"


def find_compare_index(lines: list[str], register: str) -> int:
    for index, line in enumerate(lines):
        if compare_register(line) == register:
            return index
    return -1


def compare_find_c100_by_key(lines: list[str], key: str) -> bool:
    for line in lines:
        if compare_register(line) != "C100":
            continue
        fields = line.split("|")
        if len(fields) > 9 and fields[9].strip() == key:
            return True
    return False


def compare_find_existing_invoice(lines: list[str], invoice: CompareXmlInvoice, participant_code: str = "") -> bool:
    if not compare_is_nfse_invoice(invoice):
        return compare_find_c100_by_key(lines, invoice.key)

    normalized_number = normalize_compare_document_number(invoice.number)
    normalized_series = normalize_compare_document_number("01" if compare_is_nfse_invoice(invoice) else invoice.series)
    for line in lines:
        if compare_register(line) != "C100":
            continue
        fields = line.split("|")
        cod_mod = fields[5].strip() if len(fields) > 5 else ""
        series = fields[7].strip() if len(fields) > 7 else ""
        number = fields[8].strip() if len(fields) > 8 else ""
        code = fields[4].strip() if len(fields) > 4 else ""
        if cod_mod != "01":
            continue
        if normalize_compare_document_number(number) != normalized_number:
            continue
        if normalized_series and normalize_compare_document_number(series) != normalized_series:
            continue
        if participant_code and code and code != participant_code:
            continue
        return True
    return False


def compare_next_code(existing: set[str]) -> str:
    numeric_codes = [int(code) for code in existing if code.isdigit()]
    return str(max(numeric_codes) + 1) if numeric_codes else "900000"


def compare_ensure_0150(lines: list[str], invoice: CompareXmlInvoice) -> str:
    cnpj_to_code: dict[str, str] = {}
    cnpj_to_index: dict[str, int] = {}
    existing_codes: set[str] = set()
    for index, line in enumerate(lines):
        if compare_register(line) != "0150":
            continue
        fields = line.split("|")
        code = fields[2].strip() if len(fields) > 2 else ""
        cnpj = fields[5].strip() if len(fields) > 5 else ""
        if code:
            existing_codes.add(code)
        if cnpj:
            cnpj_to_code[cnpj] = code
            cnpj_to_index[cnpj] = index
    if invoice.issuer_cnpj in cnpj_to_code and cnpj_to_code[invoice.issuer_cnpj]:
        index = cnpj_to_index.get(invoice.issuer_cnpj)
        if index is not None:
            fields = lines[index].split("|")
            while len(fields) <= 13:
                fields.append("")
            fields[3] = compare_sanitize(invoice.issuer_name) or fields[3]
            if compare_is_nfse_invoice(invoice):
                fields[7] = ""
            else:
                fields[7] = compare_sanitize(invoice.issuer_ie) or fields[7]
            fields[8] = compare_sanitize(invoice.issuer_city_code) or fields[8]
            fields[10] = compare_sanitize(invoice.issuer_address) or fields[10]
            fields[11] = compare_sanitize(invoice.issuer_number) or fields[11]
            fields[12] = compare_sanitize(invoice.issuer_complement) or fields[12]
            fields[13] = compare_sanitize(invoice.issuer_district) or fields[13]
            lines[index] = compare_line(fields)
        return cnpj_to_code[invoice.issuer_cnpj]
    code = compare_next_code(existing_codes)
    insert_index = -1
    for index, line in enumerate(lines):
        register = compare_register(line)
        if register == "0150":
            insert_index = index + 1
        elif register in {"0190", "0200", "0400", "0450", "0460", "0990"}:
            if insert_index == -1:
                insert_index = index
            break
    if insert_index == -1:
        insert_index = find_compare_index(lines, "0990")
        if insert_index == -1:
            insert_index = 0
    lines.insert(
        insert_index,
        compare_line([
            "",
            "0150",
            code,
            compare_sanitize(invoice.issuer_name),
            "1058",
            invoice.issuer_cnpj,
            "",
            "" if compare_is_nfse_invoice(invoice) else compare_sanitize(invoice.issuer_ie),
            compare_sanitize(invoice.issuer_city_code),
            "",
            compare_sanitize(invoice.issuer_address),
            compare_sanitize(invoice.issuer_number),
            compare_sanitize(invoice.issuer_complement),
            compare_sanitize(invoice.issuer_district),
            "",
        ]),
    )
    return code


def compare_ensure_0190(lines: list[str], items: list[CompareXmlItem]) -> None:
    existing_units: set[str] = set()
    for line in lines:
        if compare_register(line) == "0190":
            fields = line.split("|")
            if len(fields) > 2 and fields[2].strip():
                existing_units.add(fields[2].strip())
    insert_index = -1
    for index, line in enumerate(lines):
        register = compare_register(line)
        if register == "0190":
            insert_index = index + 1
        elif register in {"0200", "0400", "0450", "0460", "0990"}:
            if insert_index == -1:
                insert_index = index
            break
    if insert_index == -1:
        insert_index = find_compare_index(lines, "0990")
        if insert_index == -1:
            insert_index = len(lines)
    for item in items:
        unit = (compare_sanitize(item.unit) or "UN").upper()
        if unit in existing_units:
            continue
        lines.insert(insert_index, compare_line(["", "0190", unit, unit, ""]))
        insert_index += 1
        existing_units.add(unit)


def compare_ensure_0200(lines: list[str], items: list[CompareXmlItem]) -> dict[str, str]:
    existing_codes: set[str] = set()
    for line in lines:
        if compare_register(line) == "0200":
            fields = line.split("|")
            if len(fields) > 2 and fields[2].strip():
                existing_codes.add(fields[2].strip())
    insert_index = -1
    for index, line in enumerate(lines):
        register = compare_register(line)
        if register == "0200":
            insert_index = index + 1
        elif register in {"0400", "0450", "0460", "0990"}:
            if insert_index == -1:
                insert_index = index
            break
    if insert_index == -1:
        insert_index = find_compare_index(lines, "0990")
        if insert_index == -1:
            insert_index = len(lines)
    mapped_codes: dict[str, str] = {}
    for item_index, item in enumerate(items, start=1):
        code = compare_sanitize(item.code) or f"ITEM{item_index:03d}"
        base_code = code
        suffix = 1
        while code in mapped_codes.values():
            suffix += 1
            code = f"{base_code}_{suffix}"
        if code not in existing_codes:
            item_type = "09" if code == "32156" else "00"
            lines.insert(
                insert_index,
                compare_line([
                    "",
                    "0200",
                    code,
                    compare_sanitize(item.description),
                    "",
                    "",
                    compare_sanitize(item.unit) or "UN",
                    item_type,
                    compare_sanitize(item.ncm),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ]),
            )
            insert_index += 1
            existing_codes.add(code)
        mapped_codes[item.item_no or str(item_index)] = code
    return mapped_codes


def build_compare_c100(invoice: CompareXmlInvoice, participant_code: str) -> str:
    cod_mod = "01" if compare_is_nfse_invoice(invoice) else invoice.model or "55"
    cod_sit = "00"
    series = "01" if compare_is_nfse_invoice(invoice) else invoice.series
    key = "" if compare_is_nfse_invoice(invoice) else invoice.key
    return compare_line([
        "",
        "C100",
        "0",
        "1",
        participant_code,
        cod_mod,
        cod_sit,
        series,
        invoice.number,
        key,
        invoice.date_doc,
        invoice.date_entry,
        compare_format_float(invoice.total_doc),
        "0",
        compare_format_float(invoice.total_discount),
        "",
        compare_format_float(invoice.total_products),
        invoice.freight_mode if invoice.freight_mode in {"0", "1", "2", "3", "4", "9"} else "9",
        compare_format_float(invoice.total_freight),
        compare_format_float(invoice.total_insurance),
        compare_format_float(invoice.total_other),
        compare_format_float(invoice.total_bc_icms),
        compare_format_float(invoice.total_icms),
        compare_format_float(invoice.total_bc_st),
        compare_format_float(invoice.total_st),
        compare_format_float(invoice.total_ipi),
        compare_format_float(invoice.total_pis),
        compare_format_float(invoice.total_cofins),
        "0,00",
        "0,00",
        "",
    ])


def build_compare_nfse_c170(invoice: CompareXmlInvoice, mapped_codes: dict[str, str]) -> list[str]:
    total_doc = invoice.total_doc
    pis_value = round(total_doc * 0.0165, 2)
    cofins_value = round(total_doc * 0.076, 2)
    code = mapped_codes.get("1", "32156")
    return [
        compare_line([
            "",
            "C170",
            "001",
            code,
            "",
            compare_format_float(1.0, 5),
            "UN",
            compare_format_float(total_doc),
            "0,00",
            "1",
            "090",
            "1933",
            "23",
            "0,00",
            "0,00",
            "0,00",
            "0,00",
            "0,00",
            "0,00",
            "0",
            "",
            "",
            "",
            "",
            "0,00",
            "50",
            compare_format_float(total_doc),
            "1,6500",
            "",
            "",
            compare_format_float(pis_value),
            "50",
            compare_format_float(total_doc),
            "7,6000",
            "",
            "",
            compare_format_float(cofins_value),
            "",
            "",
            "",
        ])
    ]


def build_compare_c170(invoice: CompareXmlInvoice, mapped_codes: dict[str, str]) -> list[str]:
    if compare_is_nfse_invoice(invoice):
        return build_compare_nfse_c170(invoice, mapped_codes)
    rows: list[str] = []
    for item_index, item in enumerate(invoice.items, start=1):
        item_number = item.item_no.zfill(3) if item.item_no else f"{item_index:03d}"
        code = mapped_codes.get(item.item_no or str(item_index), compare_sanitize(item.code) or f"ITEM{item_index:03d}")
        cst_icms = (item.cst_icms or "000").zfill(3)
        aliq_icms = item.aliq_icms
        aliq_st = item.aliq_icms_st
        vl_bc_icms = item.vl_bc_icms
        vl_icms = item.vl_icms
        vl_bc_st = item.vl_bc_icms_st
        vl_st = item.vl_icms_st
        if cst_icms == "060":
            aliq_icms = 0.0
            aliq_st = 0.0
            vl_bc_icms = 0.0
            vl_icms = 0.0
        rows.append(
            compare_line([
                "",
                "C170",
                item_number,
                code,
                "",
                compare_format_float(item.quantity, 5),
                compare_sanitize(item.unit) or "UN",
                compare_format_float(item.value),
                compare_format_float(item.discount),
                "0",
                cst_icms,
                compare_sanitize(item.cfop) or "1102",
                "11",
                compare_format_float(vl_bc_icms),
                compare_format_float(aliq_icms, 2),
                compare_format_float(vl_icms),
                compare_format_float(vl_bc_st),
                compare_format_float(aliq_st, 2),
                compare_format_float(vl_st),
                "0",
                "",
                "",
                "",
                "",
                compare_format_float(item.vl_ipi),
                item.cst_pis or "99",
                compare_format_float(item.vl_bc_pis),
                compare_format_float(item.aliq_pis, 4),
                "",
                "",
                compare_format_float(item.vl_pis),
                item.cst_cofins or "99",
                compare_format_float(item.vl_bc_cofins),
                compare_format_float(item.aliq_cofins, 4),
                "",
                "",
                compare_format_float(item.vl_cofins),
                "",
                "",
                "",
            ])
        )
    return rows


def build_compare_c190(invoice: CompareXmlInvoice) -> list[str]:
    grouped: dict[tuple[str, str, str], dict[str, float]] = {}
    for item in invoice.items:
        cst = (item.cst_icms or "000").zfill(3)
        cfop = compare_sanitize(item.cfop) or "1102"
        rate = compare_format_float(item.aliq_icms, 2)
        key = (cst, cfop, rate)
        bucket = grouped.setdefault(
            key,
            {"vl_opr": 0.0, "vl_bc_icms": 0.0, "vl_icms": 0.0, "vl_bc_st": 0.0, "vl_st": 0.0, "vl_ipi": 0.0},
        )
        bucket["vl_opr"] += item.value
        bucket["vl_bc_icms"] += item.vl_bc_icms
        bucket["vl_icms"] += item.vl_icms
        bucket["vl_bc_st"] += item.vl_bc_icms_st
        bucket["vl_st"] += item.vl_icms_st
        bucket["vl_ipi"] += item.vl_ipi
    rows: list[str] = []
    for (cst, cfop, rate), values in sorted(grouped.items()):
        rows.append(
            compare_line([
                "",
                "C190",
                cst,
                cfop,
                rate,
                compare_format_float(values["vl_opr"]),
                compare_format_float(values["vl_bc_icms"]),
                compare_format_float(values["vl_icms"]),
                compare_format_float(values["vl_bc_st"]),
                compare_format_float(values["vl_st"]),
                "0,00",
                compare_format_float(values["vl_ipi"]),
                "",
                "",
            ])
        )
    return rows


def recalc_compare_block_counts(lines: list[str]) -> None:
    starts: dict[str, int] = {}
    ends: dict[str, int] = {}
    for index, line in enumerate(lines):
        register = compare_register(line)
        if len(register) == 4 and register.endswith("001"):
            starts[register[0]] = index
        if len(register) == 4 and register.endswith("990"):
            ends[register[0]] = index
    for block, start_index in starts.items():
        end_index = ends.get(block)
        if end_index is None or end_index < start_index:
            continue
        fields = lines[end_index].split("|")
        if len(fields) > 2:
            fields[2] = str(end_index - start_index + 1)
            lines[end_index] = compare_line(fields)
    index_0000 = find_compare_index(lines, "0000")
    index_0990 = find_compare_index(lines, "0990")
    if index_0000 >= 0 and index_0990 >= 0 and index_0990 >= index_0000:
        fields_0990 = lines[index_0990].split("|")
        if len(fields_0990) > 2:
            fields_0990[2] = str(index_0990 - index_0000 + 1)
            lines[index_0990] = compare_line(fields_0990)
    index_9001 = find_compare_index(lines, "9001")
    index_9990 = find_compare_index(lines, "9990")
    index_9999 = find_compare_index(lines, "9999")
    if index_9001 >= 0 and index_9990 >= 0 and index_9999 >= 0 and index_9999 >= index_9001:
        fields_9990 = lines[index_9990].split("|")
        if len(fields_9990) > 2:
            fields_9990[2] = str(index_9999 - index_9001 + 1)
            lines[index_9990] = compare_line(fields_9990)
    if index_9999 >= 0:
        fields_9999 = lines[index_9999].split("|")
        if len(fields_9999) > 2:
            fields_9999[2] = str(len(lines))
            lines[index_9999] = compare_line(fields_9999)


def recalc_compare_9900(lines: list[str]) -> None:
    counts: dict[str, int] = {}
    for line in lines:
        register = compare_register(line)
        if register:
            counts[register] = counts.get(register, 0) + 1
    for index, line in enumerate(lines):
        if compare_register(line) != "9900":
            continue
        fields = line.split("|")
        if len(fields) > 3:
            fields[3] = str(counts.get(fields[2].strip(), 0))
            lines[index] = compare_line(fields)


def recalc_compare_all(lines: list[str]) -> None:
    recalc_compare_block_counts(lines)
    recalc_compare_9900(lines)
    recalc_compare_block_counts(lines)


def build_compare_updated_path(original: Path) -> Path:
    return original.with_name(f"{original.stem}_ATUALIZADO{original.suffix}")


def find_compare_c100_insert_index(lines: list[str]) -> int:
    later_block_c_registers = {
        "C300", "C310", "C320", "C321", "C330", "C350", "C370", "C390",
        "C400", "C405", "C410", "C420", "C425", "C430", "C460", "C465",
        "C470", "C490", "C495", "C500", "C510", "C590", "C591", "C595",
        "C597", "C600", "C601", "C610", "C690", "C700", "C790", "C791",
        "C800", "C850", "C860", "C890", "C990",
    }
    index_c990 = find_compare_index(lines, "C990")
    for index, line in enumerate(lines):
        register = compare_register(line)
        if register in later_block_c_registers:
            return index
    if index_c990 >= 0:
        return index_c990
    raise RuntimeError("Registro C990 nao encontrado.")


def add_compare_invoice_to_lines(lines: list[str], invoice: CompareXmlInvoice) -> tuple[bool, str]:
    participant_code = compare_ensure_0150(lines, invoice)
    if compare_find_existing_invoice(lines, invoice, participant_code):
        return False, f"Nota {invoice.number} ja existe no SPED."
    compare_ensure_0190(lines, invoice.items)
    mapped_codes = compare_ensure_0200(lines, invoice.items)
    insert_index = find_compare_c100_insert_index(lines)
    has_c190 = any(compare_register(line) == "C190" for line in lines)
    lines.insert(insert_index, build_compare_c100(invoice, participant_code))
    for offset, row in enumerate(build_compare_c170(invoice, mapped_codes), start=1):
        lines.insert(insert_index + offset, row)
    if has_c190:
        base_index = insert_index + 1 + len(invoice.items)
        c190_rows = build_compare_c190(invoice)
        for offset, row in enumerate(c190_rows):
            lines.insert(base_index + offset, row)
        if compare_is_nfse_invoice(invoice):
            lines.insert(base_index + len(c190_rows), compare_line(["", "C195", "SP1933", "Obs. do lancamento fiscal - CFOP 1933", ""]))
    return True, f"Nota {invoice.number} lancada."


def launch_compare_invoices_in_sped(sped_file: Path, invoices: list[CompareXmlInvoice]) -> tuple[int, Path, str]:
    lines = sped_file.read_text(encoding="latin-1", errors="ignore").splitlines()
    launched = 0
    messages: list[str] = []
    for invoice in invoices:
        if not invoice.items:
            messages.append(f"Nota {invoice.number}: XML sem itens.")
            continue
        ok, message = add_compare_invoice_to_lines(lines, invoice)
        messages.append(message)
        if ok:
            launched += 1
    recalc_compare_all(lines)
    output_path = build_compare_updated_path(sped_file)
    output_path.write_text("\n".join(lines) + "\n", encoding="latin-1", errors="ignore")
    return launched, output_path, "\n".join(messages)


def launch_compare_invoice_in_sped(sped_file: Path, invoice: CompareXmlInvoice) -> tuple[bool, Path, str]:
    launched, output_path, message = launch_compare_invoices_in_sped(sped_file, [invoice])
    return launched > 0, output_path if launched > 0 else sped_file, message or "Nota ja existe no SPED."
