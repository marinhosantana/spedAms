from __future__ import annotations


LOG_HEADERS = [
    "Tipo", "Codigo", "Descricao", "CST SPED", "CFOP SPED",
    "CST Cadastro", "CFOP Cadastro", "CNPJ Fornecedor", "Acao",
]
LOG_FIELDS = [
    "tipo", "codigo", "descricao", "cst_sped", "cfop_sped",
    "cst_cad", "cfop_cad", "cnpj_fornecedor", "acao",
]


def build_rules_from_confronto(
    grouped_rows: list[dict],
    operation_type: str,
    compare_fields: frozenset[str] | set[str] | None = None,
) -> tuple[list[dict], list[dict]]:
    """
    Gera runtime_rules e log a partir dos resultados do confronto.

    Para cada produto divergente no cadastro, cria uma regra com new_cst / new_cfop.
    Para entrada, adiciona document_tax_id (CNPJ do fornecedor) se disponivel.
    Produtos nao cadastrados geram apenas uma entrada de log (nenhuma regra).

    Retorna (rules, log_entries).
    """
    op = operation_type.strip().lower()
    is_entry = op == "entrada"
    _fields = compare_fields if compare_fields is not None else frozenset({"cst", "cfop"})

    rules: list[dict] = []
    log_entries: list[dict] = []
    seen: set[str] = set()

    for row in grouped_rows:
        code = str(row.get("code") or "").strip()
        status = str(row.get("status") or "")
        descricao = str(row.get("descricao_sped") or "")
        cst_sped = str(row.get("cst_sped") or "")
        cfop_sped = str(row.get("cfop_sped") or "")
        cst_cad = str(row.get("cst_cad") or "").strip()
        cfop_cad = str(row.get("cfop_cad") or "").strip()
        cnpj_forn = str(row.get("cnpj_fornecedor_cad") or "").strip()

        if status == "OK":
            continue

        if "Cadastrado" in status:
            log_entries.append({
                "tipo": "NAO ENCONTRADO",
                "codigo": code,
                "descricao": descricao,
                "cst_sped": cst_sped,
                "cfop_sped": cfop_sped,
                "cst_cad": "",
                "cfop_cad": "",
                "cnpj_fornecedor": "",
                "acao": "Produto nao encontrado no cadastro — nenhuma regra gerada",
            })
            continue

        if "Divergencia" not in status:
            continue

        # Chave de deduplicacao: entrada diferencia por CNPJ, saida apenas por codigo
        dedup_key = f"{code}|{cnpj_forn}" if is_entry else code
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        if not cst_cad and not cfop_cad:
            log_entries.append({
                "tipo": "SEM CORRECAO",
                "codigo": code,
                "descricao": descricao,
                "cst_sped": cst_sped,
                "cfop_sped": cfop_sped,
                "cst_cad": cst_cad,
                "cfop_cad": cfop_cad,
                "cnpj_fornecedor": cnpj_forn,
                "acao": "Divergencia detectada mas cadastro sem valores para corrigir",
            })
            continue

        rule: dict = {"operation_type": op}
        if code:
            rule["match_codes"] = {code}
        if is_entry and cnpj_forn:
            digits = "".join(c for c in cnpj_forn if c.isdigit())
            if len(digits) in {11, 14}:
                rule["document_tax_id"] = digits

        parts: list[str] = []
        if "cst" in _fields and cst_cad:
            rule["new_cst"] = cst_cad
            parts.append(f"CST: {cst_sped} → {cst_cad}")
        if "cfop" in _fields and cfop_cad:
            rule["new_cfop"] = cfop_cad
            parts.append(f"CFOP: {cfop_sped} → {cfop_cad}")

        rules.append(rule)
        log_entries.append({
            "tipo": "REGRA GERADA",
            "codigo": code,
            "descricao": descricao,
            "cst_sped": cst_sped,
            "cfop_sped": cfop_sped,
            "cst_cad": cst_cad,
            "cfop_cad": cfop_cad,
            "cnpj_fornecedor": cnpj_forn,
            "acao": " | ".join(parts),
        })

    return rules, log_entries
