from __future__ import annotations

from decimal import Decimal, InvalidOperation


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


def _cst_diverges(sped_val: str, cat_val: str) -> bool:
    s, c = _norm_cst(sped_val), _norm_cst(cat_val)
    return bool(s and c and s != c)


def _cfop_diverges(sped_val: str, cat_val: str) -> bool:
    s, c = _norm_cfop(sped_val), _norm_cfop(cat_val)
    return bool(s and c and s != c)


def _aliq_diverges(sped_aliq: str, cat_aliq: str) -> bool:
    s, c = _norm_aliq(sped_aliq), _norm_aliq(cat_aliq)
    return s is not None and c is not None and s != c


def _status(issues: list[str], cat_found: bool) -> str:
    if not cat_found:
        return "Nao Cadastrado"
    return "OK" if not issues else "Divergencia: " + ", ".join(dict.fromkeys(issues))


def build_confronto_data(
    rows: list[dict],
    catalog_products: list[dict],
    operation_type: str,
) -> tuple[list[dict], list[dict]]:
    """
    Compara lançamentos do SPED (com launch_details) contra o cadastro local de produtos.

    Para Entrada compara: cst_icms / cfop_entrada / aliquota_icms do cadastro.
    Para Saida compara:   cst_icms_saida / cfop_saida_empresa / aliquota_icms_saida.

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

    # Indexa cadastro por codigo_empresa (primeiro encontrado por código)
    by_code: dict[str, dict] = {}
    for prod in catalog_products:
        code = str(prod.get("codigo_empresa") or "").strip()
        if code and code not in by_code:
            by_code[code] = prod

    grouped_rows: list[dict] = []
    detail_rows: list[dict] = []

    for row in rows:
        code = str(row.get("code") or "").strip()
        sped_cst  = str(row.get("cst_icms") or "").strip()
        sped_cfop = str(row.get("cfop") or "").strip()
        sped_aliq = str(row.get("display_icms_rate") or row.get("icms_rate") or "").strip()

        cat = by_code.get(code)
        cat_cst   = str(cat.get(cst_f) or "")  if cat else ""
        cat_cfop  = str(cat.get(cfop_f) or "") if cat else ""
        cat_aliq  = str(cat.get(aliq_f) or "") if cat else ""
        cat_pis   = str(cat.get(pis_f) or "")  if cat else ""
        cat_cof   = str(cat.get(cof_f) or "")  if cat else ""

        if cat:
            issues: list[str] = []
            # Campos do SPED podem ter múltiplos valores separados por |
            if any(_cst_diverges(p.strip(), cat_cst) for p in sped_cst.split("|")):
                issues.append("CST")
            if any(_cfop_diverges(p.strip(), cat_cfop) for p in sped_cfop.split("|")):
                issues.append("CFOP")
            if _aliq_diverges(sped_aliq, cat_aliq):
                issues.append("Aliq")
            grp_status = _status(issues, True)
        else:
            grp_status = "Nao Cadastrado"

        grouped_rows.append({
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
            "ncm_sped":      str(row.get("ncm") or "").strip(),
            "ncm_cad":       str(cat.get("ncm") or "")             if cat else "",
            "fornecedor_cad": str(cat.get("fornecedor_nome") or "") if cat else "",
            "total_operacao": str(row.get("sale_value") or ""),
            "base_icms":     str(row.get("base_icms") or ""),
            "icms_value":    str(row.get("icms_value") or ""),
        })

        for detail in row.get("launch_details", []):
            d_code = str(detail.get("code") or "").strip() or code
            d_cat  = by_code.get(d_code, cat)
            d_cst  = str(detail.get("cst_icms") or "").strip()
            d_cfop = str(detail.get("cfop") or "").strip()
            d_aliq = str(detail.get("icms_rate") or "").strip()

            if d_cat:
                dc_cst  = str(d_cat.get(cst_f) or "")
                dc_cfop = str(d_cat.get(cfop_f) or "")
                dc_aliq = str(d_cat.get(aliq_f) or "")
                dc_pis  = str(d_cat.get(pis_f) or "")
                dc_cof  = str(d_cat.get(cof_f) or "")
                d_issues: list[str] = []
                if _cst_diverges(d_cst, dc_cst):
                    d_issues.append("CST")
                if _cfop_diverges(d_cfop, dc_cfop):
                    d_issues.append("CFOP")
                if _aliq_diverges(d_aliq, dc_aliq):
                    d_issues.append("Aliq")
                d_status = _status(d_issues, True)
            else:
                dc_cst = dc_cfop = dc_aliq = dc_pis = dc_cof = ""
                d_status = "Nao Cadastrado"

            detail_rows.append({
                "status":          d_status,
                "periodo":         str(row.get("period") or ""),
                "document_number": str(detail.get("document_number") or ""),
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

    return grouped_rows, detail_rows
