from __future__ import annotations

from decimal import Decimal, InvalidOperation

from app.parsers.sped_parser import normalize_document_key


def _norm_cst(v: object) -> str:
    s = str(v or "").strip()
    return s.lstrip("0") or "0" if s else ""


def _norm_cfop(v: object) -> str:
    return "".join(c for c in str(v or "") if c.isdigit())


def _norm_aliq(v: object) -> Decimal | None:
    s = str(v or "").strip().replace(",", ".")
    if not s:
        return None
    try:
        return round(Decimal(s), 2)
    except InvalidOperation:
        return None


def _norm_ncm(v: object) -> str:
    return "".join(c for c in str(v or "") if c.isdigit())


def _cst_diverges(sped_val: str, cat_val: str) -> bool:
    s = _norm_cst(sped_val)
    if not s:
        return False  # SPED sem CST, nada a comparar
    return s != _norm_cst(cat_val)  # cadastro vazio ou diferente = divergencia


def _cfop_diverges(sped_val: str, cat_val: str) -> bool:
    """Ambos têm CFOP mas os valores diferem."""
    s = _norm_cfop(sped_val)
    c = _norm_cfop(cat_val)
    return bool(s) and bool(c) and s != c


def _cfop_zerado(sped_val: str, cat_val: str) -> bool:
    """SPED tem CFOP mas o cadastro está vazio/zerado."""
    s = _norm_cfop(sped_val)
    c = _norm_cfop(cat_val)
    return bool(s) and not bool(c)


def _aliq_diverges(sped_aliq: str, cat_aliq: str) -> bool:
    s, c = _norm_aliq(sped_aliq), _norm_aliq(cat_aliq)
    return s is not None and c is not None and s != c


def _ncm_diverges(sped_ncm: str, cat_ncm: str) -> bool:
    s = _norm_ncm(sped_ncm)
    c = _norm_ncm(cat_ncm)
    return bool(s) and bool(c) and s != c


def _status(issues: list[str], cat_found: bool) -> str:
    if not cat_found:
        return "Nao Cadastrado"
    return "OK" if not issues else "Divergencia: " + ", ".join(dict.fromkeys(issues))


def _append_unique_document_key(keys: list[str], value: object) -> None:
    key = normalize_document_key(str(value or "")) or str(value or "").strip()
    if key and key not in keys:
        keys.append(key)


# Campos disponíveis para seleção pelo usuário
COMPARE_FIELD_LABELS: dict[str, str] = {
    "cst":      "CST ICMS",
    "cfop":     "CFOP",
    "aliquota": "Alíquota ICMS %",
    "ncm":      "NCM",
}

DEFAULT_COMPARE_FIELDS: frozenset[str] = frozenset({"cst", "cfop"})


def build_confronto_data(
    rows: list[dict],
    catalog_products: list[dict],
    operation_type: str,
    compare_fields: frozenset[str] | set[str] = DEFAULT_COMPARE_FIELDS,
) -> tuple[list[dict], list[dict]]:
    """
    Compara lançamentos do SPED (com launch_details) contra o cadastro local de produtos.

    compare_fields controla quais campos entram na detecção de divergência:
      "cst"      → CST ICMS
      "cfop"     → CFOP
      "aliquota" → Alíquota ICMS %
      "ncm"      → NCM

    Retorna (grouped_rows, detail_rows).
    """
    if operation_type == "Entrada":
        cst_f  = "cst_icms"
        cfop_f = "cfop_entrada"
        aliq_f = "aliquota_icms"
        pis_f  = "cst_pis"
        cof_f  = "cst_cofins"
    else:
        cst_f  = "cst_icms_saida"
        cfop_f = "cfop_saida_empresa"
        aliq_f = "aliquota_icms_saida"
        pis_f  = "cst_pis_saida"
        cof_f  = "cst_cofins_saida"

    check_cst  = "cst"      in compare_fields
    check_cfop = "cfop"     in compare_fields
    check_aliq = "aliquota" in compare_fields
    check_ncm  = "ncm"      in compare_fields

    # Indexa catálogo por codigo_empresa e por ean (ambos com merge de campos).
    # O código do SPED pode ser o codigo_empresa (código interno) ou o EAN
    # do produto — a busca tenta codigo_empresa primeiro, depois EAN.
    def _merge_into(index: dict[str, dict], key: str, prod: dict) -> None:
        if not key:
            return
        if key not in index:
            index[key] = dict(prod)
        else:
            existing = index[key]
            for k, v in prod.items():
                if not str(existing.get(k) or "").strip() and str(v or "").strip():
                    existing[k] = v

    by_code: dict[str, dict] = {}   # codigo_empresa → produto mesclado
    by_ean:  dict[str, dict] = {}   # ean (só dígitos) → produto mesclado
    for prod in catalog_products:
        _merge_into(by_code, str(prod.get("codigo_empresa") or "").strip(), prod)
        ean_digits = "".join(c for c in str(prod.get("ean") or "") if c.isdigit())
        _merge_into(by_ean, ean_digits, prod)

    grouped_rows: list[dict] = []
    detail_rows: list[dict] = []

    for row in rows:
        code = str(row.get("code") or "").strip()
        sped_cst  = str(row.get("cst_icms") or "").strip()
        sped_cfop = str(row.get("cfop") or "").strip()
        sped_aliq = str(row.get("display_icms_rate") or row.get("icms_rate") or "").strip()
        sped_ncm  = str(row.get("ncm") or "").strip()

        # Tenta por codigo_empresa; se não achar, tenta pelo EAN (só dígitos)
        cat = by_code.get(code)
        if cat is None:
            code_digits = "".join(c for c in code if c.isdigit())
            if code_digits:
                cat = by_ean.get(code_digits)
        cat_cst   = str(cat.get(cst_f) or "")  if cat else ""
        cat_cfop  = str(cat.get(cfop_f) or "") if cat else ""
        cat_aliq  = str(cat.get(aliq_f) or "") if cat else ""
        cat_pis   = str(cat.get(pis_f) or "")  if cat else ""
        cat_cof   = str(cat.get(cof_f) or "")  if cat else ""
        cat_ncm   = str(cat.get("ncm") or "")  if cat else ""

        if cat:
            issues: list[str] = []
            if check_cst and any(_cst_diverges(p.strip(), cat_cst) for p in sped_cst.split("|")):
                issues.append("CST")
            if check_cfop:
                cfop_parts = sped_cfop.split("|")
                if any(_cfop_zerado(p.strip(), cat_cfop) for p in cfop_parts):
                    issues.append("CFOP zerado")
                elif any(_cfop_diverges(p.strip(), cat_cfop) for p in cfop_parts):
                    issues.append("CFOP")
            if check_aliq and _aliq_diverges(sped_aliq, cat_aliq):
                issues.append("Alíquota")
            if check_ncm and _ncm_diverges(sped_ncm, cat_ncm):
                issues.append("NCM")
            grp_status = _status(issues, True)
        else:
            grp_status = "Nao Cadastrado"

        group_document_keys: list[str] = []
        grouped_row = {
            "status":        grp_status,
            "periodo":       str(row.get("period") or ""),
            "code":          code,
            "descricao_sped": str(row.get("description") or "").strip(),
            "fornecedor":    str(row.get("suppliers") or "").strip(),
            "cst_sped":      sped_cst,
            "cst_cad":       cat_cst,
            "cfop_sped":     sped_cfop,
            "cfop_cad":      cat_cfop,
            "aliq_sped":     sped_aliq,
            "aliq_cad":      cat_aliq,
            "cst_pis_cad":   cat_pis,
            "cst_cofins_cad": cat_cof,
            "descricao_cad": str(cat.get("descricao") or "")      if cat else "",
            "ncm_sped":      sped_ncm,
            "ncm_cad":       cat_ncm,
            "fornecedor_cad":       str(cat.get("fornecedor_nome") or "")  if cat else "",
            "cnpj_fornecedor_cad":  str(cat.get("fornecedor_cnpj") or "")  if cat else "",
            "total_operacao": str(row.get("sale_value") or ""),
            "base_icms":     str(row.get("base_icms") or ""),
            "icms_value":    str(row.get("icms_value") or ""),
            "document_keys":  group_document_keys,
        }

        for detail in row.get("launch_details", []):
            d_code = str(detail.get("code") or "").strip() or code
            d_cat  = by_code.get(d_code)
            if d_cat is None:
                d_digits = "".join(c for c in d_code if c.isdigit())
                d_cat = by_ean.get(d_digits) if d_digits else None
            if d_cat is None:
                d_cat = cat
            d_cst  = str(detail.get("cst_icms") or "").strip()
            d_cfop = str(detail.get("cfop") or "").strip()
            d_aliq = str(detail.get("icms_rate") or "").strip()
            d_ncm  = str(detail.get("ncm") or "").strip() or sped_ncm

            if d_cat:
                dc_cst  = str(d_cat.get(cst_f) or "")
                dc_cfop = str(d_cat.get(cfop_f) or "")
                dc_aliq = str(d_cat.get(aliq_f) or "")
                dc_pis  = str(d_cat.get(pis_f) or "")
                dc_cof  = str(d_cat.get(cof_f) or "")
                dc_ncm  = str(d_cat.get("ncm") or "")
                d_issues: list[str] = []
                if check_cst and _cst_diverges(d_cst, dc_cst):
                    d_issues.append("CST")
                if check_cfop:
                    if _cfop_zerado(d_cfop, dc_cfop):
                        d_issues.append("CFOP zerado")
                    elif _cfop_diverges(d_cfop, dc_cfop):
                        d_issues.append("CFOP")
                if check_aliq and _aliq_diverges(d_aliq, dc_aliq):
                    d_issues.append("Alíquota")
                if check_ncm and _ncm_diverges(d_ncm, dc_ncm):
                    d_issues.append("NCM")
                d_status = _status(d_issues, True)
            else:
                dc_cst = dc_cfop = dc_aliq = dc_pis = dc_cof = dc_ncm = ""
                d_status = "Nao Cadastrado"

            detail_document_key = normalize_document_key(str(detail.get("document_key") or "")) or str(detail.get("document_key") or "").strip()
            if d_status != "OK":
                _append_unique_document_key(group_document_keys, detail_document_key or detail.get("document_number"))

            detail_rows.append({
                "status":          d_status,
                "periodo":         str(row.get("period") or ""),
                "document_number": str(detail.get("document_number") or ""),
                "document_key":    detail_document_key,
                "document_date":   str(detail.get("document_date") or ""),
                "participant":     str(detail.get("participant_name") or "").strip(),
                "code":            d_code,
                "description":     str(detail.get("description") or "").strip(),
                "cst_sped":        d_cst,
                "cst_cad":         dc_cst,
                "cfop_sped":       d_cfop,
                "cfop_cad":        dc_cfop,
                "aliq_sped":       d_aliq,
                "aliq_cad":        dc_aliq,
                "cst_pis_cad":     dc_pis,
                "cst_cofins_cad":  dc_cof,
                "sale_value":      str(detail.get("sale_value") or ""),
                "base_icms":       str(detail.get("base_icms") or ""),
                "icms_value":      str(detail.get("icms_value") or ""),
            })

        if grp_status != "OK" and not group_document_keys:
            source_keys = row.get("document_keys") or []
            if isinstance(source_keys, (str, bytes)):
                source_keys = [source_keys]
            for key in source_keys:
                _append_unique_document_key(group_document_keys, key)

        grouped_rows.append(grouped_row)

    return grouped_rows, detail_rows
