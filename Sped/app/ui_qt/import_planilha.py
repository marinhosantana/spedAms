from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.repositories.mysql_cadastro import MysqlCadastroRepository

# ── Campos destino agrupados ──────────────────────────────────────────────────

IGNORAR = "__ignorar__"

CAMPOS_GRUPOS: list[tuple[str, list[tuple[str, str]]]] = [
    ("Empresa", [
        ("empresa_nome", "Empresa - Nome"),
        ("empresa_cnpj", "Empresa - CNPJ"),
        ("empresa_ie", "Empresa - Insc. Estadual"),
    ]),
    ("Fornecedor", [
        ("fornecedor_nome", "Fornecedor - Nome"),
        ("fornecedor_cnpj", "Fornecedor - CNPJ"),
        ("fornecedor_uf", "Fornecedor - UF"),
        ("fornecedor_ie", "Fornecedor - Insc. Estadual"),
        ("fornecedor_regime", "Fornecedor - Regime Tributário"),
        ("fornecedor_codigo", "Fornecedor - Código"),
    ]),
    ("Classificação", [
        ("tipo_produto", "Classificação do Produto"),
    ]),
    ("Produto", [
        ("codigo_fornecedor", "Cód. Fornecedor"),
        ("descricao", "Descrição"),
        ("codigo_empresa", "Código Empresa"),
        ("ean", "EAN / Cód. Barras"),
        ("status_produto", "Status"),
        ("ncm", "NCM"),
        ("cest", "CEST"),
        ("origem_entrada", "Origem Entrada"),
        ("cfop_entrada", "CFOP Entrada"),
        ("cfop_saida", "CFOP Saída"),
        ("cst_icms", "CST ICMS Entrada"),
        ("aliquota_icms", "Alíquota ICMS %"),
        ("reducao_bc_icms", "Redução BC ICMS %"),
        ("cst_ipi", "CST IPI"),
        ("aliquota_ipi", "Alíquota IPI %"),
        ("cst_pis_cofins", "CST PIS/COFINS Entrada Empresa"),
        ("cst_pis", "CST PIS"),
        ("cst_cofins", "CST COFINS"),
        ("aliquota_pis", "Alíquota PIS %"),
        ("aliquota_cofins", "Alíquota COFINS %"),
        ("natureza_receita_entrada", "Natureza da Receita Entrada"),
        ("mva", "MVA %"),
        ("aliquota_icms_st", "Alíquota ICMS ST %"),
        ("bc_st", "Base ICMS ST"),
        ("valor_icms_st", "Valor ICMS ST"),
        ("cfop_saida_fornecedor", "CFOP Saída Fornecedor"),
        ("cst_icms_saida", "CST ICMS Saída"),
        ("cfop_saida_empresa", "CFOP Saída Empresa"),
        ("aliquota_icms_saida", "Alíquota ICMS Saída %"),
        ("cst_pis_saida", "CST PIS Saída"),
        ("cst_cofins_saida", "CST COFINS Saída"),
        ("natureza_receita_saida", "Natureza da Receita Saída"),
        ("c_classtrib", "Class. Tributária"),
        ("c_benef", "Cód. Benefício Fiscal"),
    ]),
]

ALL_CAMPOS: dict[str, str] = {k: v for _, fields in CAMPOS_GRUPOS for k, v in fields}

# ── Sinônimos para auto-mapeamento ────────────────────────────────────────────

_SINONIMOS: dict[str, list[str]] = {
    "empresa_nome":       ["empresa", "razao social", "razão social", "nome empresa", "company"],
    "empresa_cnpj":       ["cnpj empresa", "cnpj_empresa"],
    "empresa_ie":         ["ie empresa", "insc estadual empresa"],
    # Ordem importa: campos mais específicos primeiro para evitar falsos positivos
    "descricao":          ["descricao fornecedor", "descricao erp", "descricao produto",
                           "nome produto", "descricao", "descr"],
    "ean":                ["ean gtin trib", "ean gtin", "codigo barras", "cod barras",
                           "codigo ean", "barcode", "gtin", "ean13", "ean"],
    "codigo_fornecedor":  ["prod fornecedor", "codigo produto", "cod produto",
                           "codigo fornecedor", "cod forn", "referencia", "codigo"],
    "fornecedor_cnpj":    ["cnpj forn", "cnpj fornecedor", "insc federal",
                           "cpf cnpj", "cnpj"],
    "fornecedor_ie":      ["insc estadual", "ie forn"],
    "fornecedor_uf":      ["uf fornecedor", "uf forn", "estado forn"],
    "fornecedor_regime":  ["regime tributario", "regime icms", "tributacao"],
    "fornecedor_codigo":  ["codigo fornecedor", "id forn", "id fornecedor"],
    "fornecedor_nome":    ["nome fornecedor", "razao forn", "razao social",
                           "fornecedor nome", "nome forn", "fornecedor", "nome"],
    "tipo_produto":       ["tipo produto", "classificacao", "grupo produto",
                           "familia produto", "categoria", "familia", "tipo", "grupo"],
    "status_produto":     ["fora linha", "situacao", "status", "ativo"],
    "ncm":                ["ncm produto", "ncm"],
    "cest":               ["cest"],
    "origem_entrada":     ["origem entrada", "origem"],
    "cfop_entrada":       ["cfop entrada", "cfop ent"],
    "cfop_saida":         ["cfop saida", "cfop sai"],
    "cst_icms":           ["cst icms entrada", "cst icms"],
    "aliquota_icms":      ["aliquota icms", "aliq icms", "perc icms"],
    "reducao_bc_icms":    ["reducao bc icms", "reducao bc", "red bc icms"],
    "cst_ipi":            ["cst ipi"],
    "aliquota_ipi":       ["aliquota ipi", "aliq ipi"],
    "cst_pis_cofins":     ["cst pis cofins entrada", "cst piscofins", "cst pis/cofins entrada", "cst pis cofins emp"],
    "cst_pis":            ["cst pis entrada", "cst pis"],
    "cst_cofins":         ["cst cofins entrada", "cst cofins"],
    "aliquota_pis":       ["aliquota pis", "aliq pis"],
    "aliquota_cofins":    ["aliquota cofins", "aliq cofins"],
    "natureza_receita_entrada": ["natureza receita entrada", "nat receita ent", "natureza receita",
                                 "natureza da receita", "nat da receita"],
    "natureza_receita_saida":   ["natureza receita saida", "nat receita sai",
                                 "natureza da receita 1", "natureza da receita saida", "nat da receita sai"],
    "cfop_saida_fornecedor":    ["cfop saida fornecedor", "cfop sai forn", "cfop fornecedor"],
    "mva":                ["mva"],
    "aliquota_icms_st":   ["aliquota icms st", "aliq icms st", "aliquota st"],
    "valor_icms_st":      ["valor icms st", "valor st"],
    "bc_st":              ["base icms st", "base st", "bc st"],
    "empresa_nome":       ["empresa", "razao social empresa", "nome empresa"],
    "empresa_cnpj":       ["cnpj empresa"],
    "empresa_ie":         ["insc estadual empresa", "ie empresa"],
}


def _strip_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _auto_map(col_name: str) -> str:
    # Remove acentos, converte para minúsculo, troca _ por espaço, remove chars especiais
    clean = _strip_accents(col_name.lower().strip())
    clean = clean.replace("_", " ")
    normalized = re.sub(r"[^a-z0-9 ]", " ", clean)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    # 1ª passagem: match exato
    for campo, sinonimos in _SINONIMOS.items():
        for sin in sinonimos:
            if normalized == sin.replace("_", " "):
                return campo
    # 2ª passagem: match por palavras inteiras (evita "ie" bater em "especie")
    for campo, sinonimos in _SINONIMOS.items():
        for sin in sinonimos:
            sin_norm = sin.replace("_", " ")
            # só faz match parcial para sinônimos com 4+ chars
            if len(sin_norm) >= 4:
                pattern = r"\b" + re.escape(sin_norm) + r"\b"
                if re.search(pattern, normalized):
                    return campo
    return IGNORAR


# ── Perfis de mapeamento ──────────────────────────────────────────────────────

_PROFILES_FILE = Path.home() / ".spedams" / "import_profiles.json"


def _load_all_profiles() -> dict[str, dict[str, str]]:
    if _PROFILES_FILE.exists():
        try:
            return json.loads(_PROFILES_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_all_profiles(profiles: dict[str, dict[str, str]]) -> None:
    _PROFILES_FILE.parent.mkdir(parents=True, exist_ok=True)
    _PROFILES_FILE.write_text(
        json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── Worker de importação ──────────────────────────────────────────────────────

class _NoScrollCombo(QComboBox):
    """QComboBox que ignora scroll do mouse para não trocar valor ao rolar a página."""
    def wheelEvent(self, event: Any) -> None:
        event.ignore()


class _ImportWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        repository: MysqlCadastroRepository,
        environment: str,
        default_company_id: int | None,
        col_mapping: dict[str, str],
        df: Any,
    ) -> None:
        super().__init__()
        self.repository = repository
        self.environment = environment
        self.default_company_id = default_company_id
        self.col_mapping = col_mapping
        self.df = df

    def run(self) -> None:
        try:
            result = _execute_import(
                self.repository,
                self.environment,
                self.default_company_id,
                self.col_mapping,
                self.df,
                self.progress.emit,
            )
            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


def _get_val(row: Any, col_mapping: dict[str, str], campo: str) -> str:
    col = col_mapping.get(campo)
    if col is None:
        return ""
    v = row.get(col, "")
    if v is None:
        return ""
    sv = str(v).strip()
    return "" if sv.lower() in {"nan", "none", ""} else sv


def _needs_update(existing: dict, incoming: dict) -> bool:
    """True se algum campo não-vazio da planilha difere do registro já no banco."""
    if not existing or len(existing) <= 3:
        return True  # cache mínimo — assume que precisa atualizar

    def nt(v: object) -> str:
        return str(v or "").strip().lower()

    def nd(v: object) -> str:
        return "".join(c for c in str(v or "") if c.isdigit())

    for f in ("descricao", "status_produto", "codigo_empresa", "origem_entrada",
              "cfop_entrada", "cfop_saida", "cfop_saida_fornecedor", "cfop_saida_empresa",
              "cst_icms", "cst_ipi", "cst_pis", "cst_cofins",
              "cst_pis_cofins", "natureza_receita_entrada", "natureza_receita_saida",
              "cst_icms_saida", "cst_pis_saida", "cst_cofins_saida"):
        v_in = nt(incoming.get(f, ""))
        if v_in and v_in != nt(existing.get(f, "")):
            return True

    for f in ("ncm", "cest", "ean"):
        v_in = nd(incoming.get(f, ""))
        if v_in and v_in != nd(existing.get(f, "")):
            return True

    tp_in = incoming.get("tipo_produto_id")
    if tp_in and str(tp_in) != str(existing.get("tipo_produto_id") or ""):
        return True

    for f in ("aliquota_icms", "aliquota_ipi", "aliquota_pis", "aliquota_cofins", "mva"):
        v_in = nt(incoming.get(f, "")).replace(",", ".")
        v_ex = nt(existing.get(f, "")).replace(",", ".")
        if v_in and v_in != v_ex:
            return True

    return False


def _build_update_data(existing: dict, incoming: dict) -> dict:
    """Mescla dados para atualização: preserva valores do banco para campos vazios na planilha.

    Apenas campos com valor não-vazio na planilha sobrescrevem o banco.
    Isso protege dados vindos de XML/SPED de serem apagados por colunas não mapeadas.
    """
    result = dict(existing)
    for field, value in incoming.items():
        if field == "id":
            continue
        if field == "tipo_produto_id":
            # None = campo não mapeado na planilha; não apaga o tipo existente
            if value is not None:
                result[field] = value
            continue
        v_str = str(value or "").strip()
        if v_str != "":
            result[field] = value
    return result


def _build_product_cache(repo: MysqlCadastroRepository, supplier_id: int) -> dict[str, dict]:
    """Carrega todos os produtos de um fornecedor em memória para lookups O(1)."""
    cache: dict[str, dict] = {}
    for p in repo.list_supplier_products(supplier_id):
        ean = str(p.get("ean") or "").strip()
        code = str(p.get("codigo_fornecedor") or "").strip()
        if ean:
            cache[f"ean:{ean}"] = p
        if code:
            cache[f"cod:{code}"] = p
    return cache


def _execute_import(
    repo: MysqlCadastroRepository,
    environment: str,
    default_company_id: int | None,
    col_mapping: dict[str, str],
    df: Any,
    progress_cb: Any,
) -> dict:
    stats: dict = {
        "linhas": len(df),
        "novos": 0,
        "atualizados": 0,
        "sem_alteracao": 0,
        "ignorados": 0,
        "empresas_criadas": 0,
        "fornecedores_criados": 0,
        "tipos_criados": 0,
        "erros": [],
        "nao_atualizados": [],  # {"linha", "fornecedor", "codigo", "ean", "descricao", "motivo"}
    }

    BATCH_SIZE = 500  # produtos por transação

    seen_empresas: dict[tuple, int] = {}
    seen_fornecedores: dict[tuple, int] = {}
    seen_tipos: dict[str, int] = {
        t["nome"]: int(t["id"]) for t in repo.list_product_types(environment)
    }
    supplier_product_cache: dict[int, dict[str, dict]] = {}
    # Mapa lazy por empresa: {empresa_id: {codigo_empresa: cfop_entrada}}
    # Carregado na primeira vez que a empresa aparece; permite herdar CFOP
    # de outro fornecedor que já tenha o mesmo codigo_empresa configurado.
    cfop_map_by_empresa: dict[int, dict[str, str]] = {}

    total = len(df)
    notify_step = max(1, total // 200)

    # Lote acumulado: (idx_linha, fornecedor_id, forn_nome, data_dict, existing_prod|None, ean, codigo)
    pending: list[tuple] = []

    def _err(idx_l: int, forn: str, cod: str, ean_v: str, desc: str, msg: str) -> dict:
        return {"linha": idx_l + 2, "fornecedor": forn, "codigo": cod,
                "ean": ean_v, "descricao": desc[:80], "erro": msg}

    def _skip(idx_l: int, forn: str, cod: str, ean_v: str, desc: str, motivo: str) -> None:
        stats["nao_atualizados"].append(
            {"linha": idx_l + 2, "fornecedor": forn, "codigo": cod,
             "ean": ean_v, "descricao": desc[:80], "motivo": motivo}
        )

    def flush_batch() -> None:
        if not pending:
            return
        records = [
            {"supplier_id": fid, "data": d, "existing_id": ex["id"] if ex else None}
            for _, fid, _fn, d, ex, _, _ in pending
        ]
        try:
            ids = repo.bulk_upsert_supplier_products(records)
        except Exception:
            # Falha em lote: retenta linha a linha para identificar o problema
            ids = []
            for (idx_l, fid, fn, d, ex, ean_v, cod), rec in zip(pending, records):
                try:
                    new_id = repo.save_supplier_product(rec["supplier_id"], rec["data"])
                    ids.append(new_id)
                except Exception as e2:
                    ids.append(0)
                    stats["erros"].append(
                        _err(idx_l, fn, cod, ean_v, d.get("descricao", ""), str(e2))
                    )

        for (idx_l, fid, fn, d, ex, ean_v, cod), new_id in zip(pending, ids):
            cache = supplier_product_cache.get(fid, {})
            if new_id:
                entry = {"id": new_id, "ean": ean_v, "codigo_fornecedor": cod}
                if ean_v:
                    cache[f"ean:{ean_v}"] = entry
                if cod:
                    cache[f"cod:{cod}"] = entry
            if ex and ex.get("id"):
                stats["atualizados"] += 1
            else:
                stats["novos"] += 1
        pending.clear()

    for idx in range(total):
        if idx % notify_step == 0:
            progress_cb(idx + 1, total, f"Processando {idx + 1}/{total}...")
        row = df.iloc[idx].to_dict()

        try:
            def get(campo: str, _row: dict = row) -> str:
                return _get_val(_row, col_mapping, campo)

            # ── Empresa ──────────────────────────────────────────────────────
            emp_nome = get("empresa_nome")
            emp_cnpj = get("empresa_cnpj")
            emp_ie   = get("empresa_ie")

            if emp_nome:
                emp_key = (emp_nome.lower(), emp_cnpj)
                if emp_key not in seen_empresas:
                    if repo.find_company_id(environment, emp_nome, emp_cnpj) is None:
                        stats["empresas_criadas"] += 1
                    seen_empresas[emp_key] = repo.ensure_company(
                        environment, emp_nome, emp_cnpj, emp_ie
                    )
                empresa_id = seen_empresas[emp_key]
            elif default_company_id is not None:
                empresa_id = default_company_id
            else:
                stats["ignorados"] += 1
                _skip(idx, "", "", "", get("descricao"), "Sem empresa (campo empresa_nome vazio e sem empresa padrão)")
                continue

            # ── Fornecedor ───────────────────────────────────────────────────
            forn_nome = get("fornecedor_nome")
            if not forn_nome:
                stats["ignorados"] += 1
                _skip(idx, "", "", "", get("descricao"), "Sem fornecedor (campo fornecedor_nome vazio)")
                continue

            forn_cnpj   = get("fornecedor_cnpj")
            forn_uf     = get("fornecedor_uf")
            forn_ie     = get("fornecedor_ie")
            forn_regime = get("fornecedor_regime") or "LUCRO_REAL_PRESUMIDO"

            forn_key = (empresa_id, forn_nome.lower(), forn_cnpj)
            if forn_key not in seen_fornecedores:
                if repo.find_supplier_id(empresa_id, forn_nome, forn_cnpj) is None:
                    stats["fornecedores_criados"] += 1
                seen_fornecedores[forn_key] = repo.ensure_supplier(
                    empresa_id, forn_nome, forn_cnpj, forn_ie, forn_uf, forn_regime
                )
            fornecedor_id = seen_fornecedores[forn_key]

            # ── Classificação ────────────────────────────────────────────────
            tipo_nome = get("tipo_produto")
            tipo_id: int | None = None
            if tipo_nome:
                if tipo_nome not in seen_tipos:
                    stats["tipos_criados"] += 1
                    seen_tipos[tipo_nome] = repo.ensure_product_type(environment, tipo_nome)
                tipo_id = seen_tipos[tipo_nome]

            # ── Produto ──────────────────────────────────────────────────────
            codigo    = get("codigo_fornecedor")
            descricao = get("descricao")
            ean       = get("ean")

            if not codigo:
                codigo = ean or descricao[:80]
            if not codigo:
                stats["ignorados"] += 1
                _skip(idx, forn_nome, "", "", "", "Sem identificador (código, EAN e descrição todos vazios)")
                continue

            if fornecedor_id not in supplier_product_cache:
                supplier_product_cache[fornecedor_id] = _build_product_cache(repo, fornecedor_id)
            cache = supplier_product_cache[fornecedor_id]

            existing_prod = (cache.get(f"ean:{ean}") if ean else None) or cache.get(f"cod:{codigo}")

            # Monta os dados antes de comparar
            data: dict = {
                "tipo_produto_id":     tipo_id,
                "codigo_fornecedor":   codigo,
                "codigo_empresa":      get("codigo_empresa"),
                "status_produto":      get("status_produto"),
                "descricao":           descricao,
                "ean":                 ean,
                "ncm":                 get("ncm"),
                "cest":                get("cest"),
                "origem_entrada":      get("origem_entrada"),
                "cfop_entrada":        get("cfop_entrada"),
                "cfop_saida":          get("cfop_saida"),
                "cst_icms":            get("cst_icms"),
                "aliquota_icms":       get("aliquota_icms"),
                "reducao_bc_icms":     get("reducao_bc_icms"),
                "cst_ipi":             get("cst_ipi"),
                "aliquota_ipi":        get("aliquota_ipi"),
                "cst_pis_cofins":      get("cst_pis_cofins"),
                "cst_pis":             get("cst_pis"),
                "cst_cofins":          get("cst_cofins"),
                "aliquota_pis":        get("aliquota_pis"),
                "aliquota_cofins":     get("aliquota_cofins"),
                "natureza_receita_entrada": get("natureza_receita_entrada"),
                "mva":                 get("mva"),
                "aliquota_icms_st":    get("aliquota_icms_st"),
                "bc_st":               get("bc_st"),
                "valor_icms_st":       get("valor_icms_st"),
                "cfop_saida_fornecedor": get("cfop_saida_fornecedor"),
                "cst_icms_saida":      get("cst_icms_saida"),
                "cfop_saida_empresa":  get("cfop_saida_empresa"),
                "aliquota_icms_saida": get("aliquota_icms_saida"),
                "cst_pis_saida":       get("cst_pis_saida"),
                "cst_cofins_saida":    get("cst_cofins_saida"),
                "natureza_receita_saida": get("natureza_receita_saida"),
                "c_classtrib":         get("c_classtrib"),
                "c_benef":             get("c_benef"),
                "fornecedor_codigo":   get("fornecedor_codigo"),
            }

            # ── Herança de CFOP por codigo_empresa ───────────────────────────
            # Se a planilha não trouxe cfop_entrada, mas o produto tem
            # codigo_empresa e outro fornecedor da mesma empresa já tem
            # esse CFOP configurado → herda o valor.
            if not data.get("cfop_entrada"):
                cod_emp = str(data.get("codigo_empresa") or "").strip()
                if cod_emp and empresa_id is not None:
                    if empresa_id not in cfop_map_by_empresa:
                        cfop_map_by_empresa[empresa_id] = repo.get_cfop_map_by_empresa_id(empresa_id)
                    inherited = cfop_map_by_empresa[empresa_id].get(cod_emp, "")
                    if inherited:
                        data["cfop_entrada"] = inherited

            # Se produto já existe e nenhum campo mudou → pula sem gravar
            if existing_prod and not _needs_update(existing_prod, data):
                stats["sem_alteracao"] += 1
                _skip(idx, forn_nome, codigo, ean, descricao, "Sem alteração — dados da planilha já iguais ao banco")
                continue

            # Para produto existente: mescla apenas campos não-vazios da planilha,
            # preservando valores do banco (XML/SPED) nos campos não mapeados.
            effective_data = _build_update_data(existing_prod, data) if existing_prod else data

            # Reserva o slot no cache imediatamente para evitar duplicatas dentro do lote
            placeholder = {"id": None, "ean": ean, "codigo_fornecedor": codigo}
            if ean:
                cache[f"ean:{ean}"] = placeholder
            if codigo:
                cache[f"cod:{codigo}"] = placeholder

            pending.append((idx, fornecedor_id, forn_nome, effective_data, existing_prod, ean, codigo))

            if len(pending) >= BATCH_SIZE:
                flush_batch()
                progress_cb(idx + 1, total, f"Gravado {idx + 1}/{total}...")

        except Exception as exc:
            stats["erros"].append(_err(
                idx,
                locals().get("forn_nome", ""),
                locals().get("codigo", ""),
                locals().get("ean", ""),
                locals().get("descricao", ""),
                str(exc),
            ))

    flush_batch()
    progress_cb(total, total, "Finalizando...")
    return stats


# ── Worker de carregamento de arquivo ─────────────────────────────────────────

class _FileLoaderWorker(QObject):
    progress = Signal(str)
    finished = Signal(list, int, object, int, object)  # sheet_names, header_row, df_preview, total_rows, df_full
    failed = Signal(str)

    def __init__(
        self,
        file_path: Path,
        sheet: Any,
        header_row: int,
        detect_header: bool,
        read_sheets: bool,
    ) -> None:
        super().__init__()
        self._file_path = file_path
        self._sheet = sheet
        self._header_row = header_row
        self._detect_header = detect_header
        self._read_sheets = read_sheets
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            ext = self._file_path.suffix.lower()
            sheet_names: list[str] = []
            sheet = self._sheet
            header_row = self._header_row

            if self._read_sheets:
                self.progress.emit("Lendo estrutura do arquivo...")
                if ext in {".xlsx", ".xls", ".ods"}:
                    xf = pd.ExcelFile(str(self._file_path))
                    sheet_names = list(xf.sheet_names)
                    xf.close()
                    if sheet_names and (sheet == 0 or sheet not in sheet_names):
                        sheet = sheet_names[0]

            if self._cancelled:
                return

            if self._detect_header:
                self.progress.emit("Detectando cabecalho...")
                best_row = 0
                best_score = -1
                for row in range(5):
                    if self._cancelled:
                        return
                    try:
                        if ext == ".csv":
                            df = pd.read_csv(str(self._file_path), header=row, nrows=0, dtype=str)
                        else:
                            df = pd.read_excel(str(self._file_path), sheet_name=sheet, header=row, nrows=0, dtype=str)
                        total = len(df.columns)
                        if total == 0:
                            continue
                        named = sum(1 for c in df.columns if not str(c).startswith("Unnamed:"))
                        score = named / total
                        if score > best_score:
                            best_score = score
                            best_row = row
                        if score >= 0.9:
                            break
                    except Exception:
                        break
                header_row = best_row

            if self._cancelled:
                return

            self.progress.emit("Carregando preview...")
            if ext == ".csv":
                df_preview = pd.read_csv(
                    str(self._file_path), header=header_row, nrows=8, dtype=str, keep_default_na=False,
                )
            else:
                df_preview = pd.read_excel(
                    str(self._file_path), sheet_name=sheet, header=header_row, nrows=8, dtype=str,
                )

            if self._cancelled:
                return

            self.progress.emit("Contando linhas...")
            if ext == ".csv":
                df_full = pd.read_csv(str(self._file_path), header=header_row, dtype=str, keep_default_na=False)
            else:
                df_full = pd.read_excel(str(self._file_path), sheet_name=sheet, header=header_row, dtype=str)
            total_rows = len(df_full)

            if self._cancelled:
                return

            self.finished.emit(sheet_names, header_row, df_preview, total_rows, df_full)
        except Exception as exc:
            if not self._cancelled:
                self.failed.emit(str(exc))


# ── Dialog principal ──────────────────────────────────────────────────────────

class ImportPlanilhaDialog(QDialog):
    PREVIEW_ROWS = 6
    COLORS = {
        "bg": "#f3f6f9",
        "panel": "#ffffff",
        "line": "#d6e0e8",
        "head": "#edf3f7",
        "text": "#1d2730",
        "muted": "#5d6e7d",
        "accent": "#1f6fd1",
        "ok": "#1f8f5f",
        "bad": "#b42318",
        "warn": "#b56b12",
    }

    def __init__(self, repository: MysqlCadastroRepository, environment: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.environment = environment

        self._df: Any = None
        self._all_columns: list[str] = []
        self._file_path: Path | None = None
        self._mapping_combos: dict[str, QComboBox] = {}
        self._thread: QThread | None = None
        self._worker: _ImportWorker | None = None
        self._file_load_thread: QThread | None = None
        self._file_load_worker: _FileLoaderWorker | None = None
        self._profiles: dict[str, dict[str, str]] = _load_all_profiles()

        self.setWindowTitle("Importar Planilha de Produtos")
        self.setMinimumSize(1000, 680)
        if parent:
            self.setStyleSheet(parent.styleSheet())
        self.resize(1100, 740)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Cabeçalho de etapas
        self._step_header = self._build_step_header()
        root.addWidget(self._step_header)

        # Conteúdo das etapas
        self._stack = QStackedWidget()
        root.addWidget(self._stack, 1)

        self._stack.addWidget(self._build_step_arquivo())
        self._stack.addWidget(self._build_step_mapeamento())
        self._stack.addWidget(self._build_step_importar())

        # Barra de navegação
        nav = self._build_nav_bar()
        root.addWidget(nav)

        self._update_nav()

    # ── Cabeçalho de etapas ───────────────────────────────────────────────────

    def _build_step_header(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("panel")
        bar.setFixedHeight(52)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(20, 8, 20, 8)
        layout.setSpacing(0)

        self._step_labels: list[QLabel] = []
        steps = ["1. Selecionar Arquivo", "2. Correlacionar Campos", "3. Importar"]
        for i, text in enumerate(steps):
            lbl = QLabel(text)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setFixedHeight(36)
            self._step_labels.append(lbl)
            layout.addWidget(lbl, 1)
            if i < len(steps) - 1:
                sep = QLabel("›")
                sep.setAlignment(Qt.AlignCenter)
                sep.setFixedWidth(24)
                sep.setStyleSheet(f"color: {self.COLORS['muted']}; font-size: 18px;")
                layout.addWidget(sep)

        self._highlight_step(0)
        return bar

    def _highlight_step(self, active: int) -> None:
        for i, lbl in enumerate(self._step_labels):
            if i == active:
                lbl.setStyleSheet(
                    f"color: {self.COLORS['accent']}; font-weight: bold; "
                    f"background: #dbeafe; border-radius: 6px; padding: 4px 12px;"
                )
            elif i < active:
                lbl.setStyleSheet(f"color: {self.COLORS['ok']}; padding: 4px 12px;")
            else:
                lbl.setStyleSheet(f"color: {self.COLORS['muted']}; padding: 4px 12px;")

    # ── Barra de navegação ────────────────────────────────────────────────────

    def _build_nav_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("panel")
        bar.setFixedHeight(52)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(20, 8, 20, 8)
        layout.setSpacing(8)

        self._btn_anterior = QPushButton("< Anterior")
        self._btn_anterior.clicked.connect(self._go_back)

        self._btn_proximo = QPushButton("Próximo >")
        self._btn_proximo.setObjectName("primaryButton")
        self._btn_proximo.clicked.connect(self._go_next)

        self._btn_importar = QPushButton("Importar")
        self._btn_importar.setObjectName("primaryButton")
        self._btn_importar.clicked.connect(self._start_import)

        self._btn_fechar = QPushButton("Fechar")
        self._btn_fechar.clicked.connect(self.accept)
        self._btn_fechar.hide()

        layout.addStretch()
        layout.addWidget(self._btn_anterior)
        layout.addWidget(self._btn_proximo)
        layout.addWidget(self._btn_importar)
        layout.addWidget(self._btn_fechar)
        return bar

    def _update_nav(self) -> None:
        step = self._stack.currentIndex()
        self._btn_anterior.setVisible(step > 0)
        self._btn_proximo.setVisible(step < 2)
        self._btn_importar.setVisible(step == 2)
        self._highlight_step(step)

        if step == 0:
            self._btn_proximo.setEnabled(self._df is not None)
        elif step == 1:
            self._btn_proximo.setEnabled(True)

    # ── Passo 1: Selecionar Arquivo ───────────────────────────────────────────

    def _build_step_arquivo(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet(f"background: {self.COLORS['bg']};")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(16)

        # Seleção de arquivo
        file_group = QGroupBox("Planilha")
        file_layout = QGridLayout(file_group)
        file_layout.setContentsMargins(12, 12, 12, 12)
        file_layout.setSpacing(8)

        file_layout.addWidget(QLabel("Arquivo:"), 0, 0)
        self._file_label = QLabel("Nenhum arquivo selecionado")
        self._file_label.setObjectName("muted")
        self._file_label.setWordWrap(True)
        file_layout.addWidget(self._file_label, 0, 1)
        btn_select = QPushButton("Selecionar...")
        btn_select.setFixedWidth(110)
        btn_select.clicked.connect(self._select_file)
        file_layout.addWidget(btn_select, 0, 2)

        file_layout.addWidget(QLabel("Aba:"), 1, 0)
        self._sheet_combo = QComboBox()
        self._sheet_combo.setEnabled(False)
        self._sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        file_layout.addWidget(self._sheet_combo, 1, 1, 1, 2)

        # Linha do cabeçalho
        header_row_widget = QWidget()
        header_row_widget.setStyleSheet("background: transparent;")
        header_row_layout = QHBoxLayout(header_row_widget)
        header_row_layout.setContentsMargins(0, 0, 0, 0)
        header_row_layout.setSpacing(6)

        self._header_check = QCheckBox("Tem cabeçalho — linha:")
        self._header_check.setChecked(True)
        self._header_check.stateChanged.connect(self._on_header_check_changed)
        header_row_layout.addWidget(self._header_check)

        self._header_row_spin = QSpinBox()
        self._header_row_spin.setRange(0, 20)
        self._header_row_spin.setValue(0)
        self._header_row_spin.setFixedWidth(60)
        self._header_row_spin.setToolTip(
            "0 = primeira linha do arquivo\n1 = segunda linha\netc."
        )
        self._header_row_spin.valueChanged.connect(self._refresh_preview)
        header_row_layout.addWidget(self._header_row_spin)

        lbl_hint = QLabel("(0 = primeira linha)")
        lbl_hint.setStyleSheet(f"color: {self.COLORS['muted']}; background: transparent; font-size: 11px;")
        header_row_layout.addWidget(lbl_hint)
        header_row_layout.addStretch()

        file_layout.addWidget(header_row_widget, 2, 0, 1, 3)

        file_layout.setColumnStretch(1, 1)
        layout.addWidget(file_group)

        # Empresa padrão (usada quando empresa_nome não for mapeado)
        emp_group = QGroupBox("Empresa destino (quando não mapeada na planilha)")
        emp_layout = QHBoxLayout(emp_group)
        emp_layout.setContentsMargins(12, 12, 12, 12)
        emp_layout.setSpacing(8)
        emp_layout.addWidget(QLabel("Empresa:"))
        self._company_combo = QComboBox()
        self._company_combo.setMinimumWidth(320)
        self._load_company_combo()
        emp_layout.addWidget(self._company_combo, 1)
        layout.addWidget(emp_group)

        # Preview
        preview_group = QGroupBox("Preview (primeiras linhas)")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(8, 8, 8, 8)
        self._preview_table = QTableWidget()
        self._preview_table.setAlternatingRowColors(True)
        self._preview_table.setSelectionMode(QAbstractItemView.NoSelection)
        self._preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._preview_table.verticalHeader().setVisible(False)
        self._preview_table.setFixedHeight(200)
        self._preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        preview_layout.addWidget(self._preview_table)
        layout.addWidget(preview_group, 1)

        self._file_info_label = QLabel("")
        self._file_info_label.setObjectName("muted")
        layout.addWidget(self._file_info_label)

        # Barra de progresso de carregamento
        load_bar_widget = QWidget()
        load_bar_widget.setStyleSheet("background: transparent;")
        load_bar_layout = QHBoxLayout(load_bar_widget)
        load_bar_layout.setContentsMargins(0, 0, 0, 0)
        load_bar_layout.setSpacing(8)
        self._load_status_label = QLabel("")
        self._load_status_label.setStyleSheet(f"color: {self.COLORS['muted']}; font-size: 11px; background: transparent;")
        self._load_progress_bar = QProgressBar()
        self._load_progress_bar.setRange(0, 0)
        self._load_progress_bar.setFixedHeight(8)
        self._load_progress_bar.setTextVisible(False)
        load_bar_layout.addWidget(self._load_status_label)
        load_bar_layout.addWidget(self._load_progress_bar, 1)
        load_bar_widget.setVisible(False)
        self._load_bar_widget = load_bar_widget
        layout.addWidget(load_bar_widget)

        return page

    def _load_company_combo(self) -> None:
        self._company_combo.clear()
        self._company_combo.addItem("(nenhuma — usar coluna da planilha)", None)
        try:
            companies = self.repository.list_companies(self.environment)
            for c in companies:
                self._company_combo.addItem(c["nome"], c["id"])
        except Exception:
            pass

    def _select_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Planilha",
            "",
            "Planilhas (*.xlsx *.xls *.csv *.ods);;Todos (*)",
        )
        if not path:
            return
        self._file_path = Path(path)
        self._file_label.setText(str(self._file_path.name))
        self._file_label.setStyleSheet(f"color: {self.COLORS['text']};")
        self._df = None
        self._all_columns = []
        self._sheet_combo.blockSignals(True)
        self._sheet_combo.clear()
        self._sheet_combo.setEnabled(False)
        self._sheet_combo.blockSignals(False)
        self._preview_table.setRowCount(0)
        self._preview_table.setColumnCount(0)
        self._file_info_label.setText("")
        self._update_nav()
        self._start_file_load(read_sheets=True, detect_header=True)

    def _load_sheets(self) -> None:
        if self._file_path is None:
            return
        self._sheet_combo.blockSignals(True)
        self._sheet_combo.clear()
        ext = self._file_path.suffix.lower()
        if ext in {".xlsx", ".xls", ".ods"}:
            try:
                xf = pd.ExcelFile(str(self._file_path))
                for name in xf.sheet_names:
                    self._sheet_combo.addItem(name)
                self._sheet_combo.setEnabled(len(xf.sheet_names) > 1)
            except Exception as exc:
                self._file_label.setText(f"Erro ao abrir: {exc}")
                self._sheet_combo.blockSignals(False)
                return
        else:
            self._sheet_combo.addItem("(csv)")
            self._sheet_combo.setEnabled(False)
        self._sheet_combo.blockSignals(False)
        self._auto_detect_header()
        self._refresh_preview()

    def _on_sheet_changed(self, _: str) -> None:
        if self._file_path is None:
            return
        self._start_file_load(read_sheets=False, detect_header=True)

    def _on_header_check_changed(self) -> None:
        self._header_row_spin.setEnabled(self._header_check.isChecked())
        if self._file_path is not None:
            self._start_file_load(read_sheets=False, detect_header=False)

    def _start_file_load(self, read_sheets: bool, detect_header: bool) -> None:
        if self._file_path is None:
            return
        if self._file_load_worker is not None:
            self._file_load_worker.cancel()
        if self._file_load_thread is not None:
            self._file_load_thread.quit()
            self._file_load_thread.wait(300)

        ext = self._file_path.suffix.lower()
        sheet: Any = self._sheet_combo.currentText() if self._sheet_combo.isEnabled() else 0
        if sheet in {"(csv)", ""}:
            sheet = 0
        has_header = self._header_check.isChecked()
        header_row = self._header_row_spin.value() if has_header else None
        if header_row is None:
            detect_header = False
            header_row = 0

        self._load_bar_widget.setVisible(True)
        self._load_status_label.setText("Aguarde...")
        self._btn_proximo.setEnabled(False)

        self._file_load_thread = QThread(self)
        self._file_load_worker = _FileLoaderWorker(
            self._file_path, sheet, header_row, detect_header, read_sheets,
        )
        self._file_load_worker.moveToThread(self._file_load_thread)
        self._file_load_thread.started.connect(self._file_load_worker.run)
        self._file_load_worker.progress.connect(self._on_file_load_progress)
        self._file_load_worker.finished.connect(self._on_file_load_done)
        self._file_load_worker.failed.connect(self._on_file_load_failed)
        self._file_load_worker.finished.connect(self._file_load_thread.quit)
        self._file_load_worker.failed.connect(self._file_load_thread.quit)
        self._file_load_thread.finished.connect(self._file_load_worker.deleteLater)
        self._file_load_thread.finished.connect(self._file_load_thread.deleteLater)
        self._file_load_thread.finished.connect(lambda: setattr(self, "_file_load_thread", None))
        self._file_load_thread.finished.connect(lambda: setattr(self, "_file_load_worker", None))
        self._file_load_thread.start()

    def _on_file_load_progress(self, message: str) -> None:
        self._load_status_label.setText(message)

    def _on_file_load_done(self, sheet_names: list, header_row: int, df_preview: Any, total_rows: int, df_full: Any) -> None:
        self._load_bar_widget.setVisible(False)

        if sheet_names:
            self._sheet_combo.blockSignals(True)
            self._sheet_combo.clear()
            for name in sheet_names:
                self._sheet_combo.addItem(name)
            self._sheet_combo.setEnabled(len(sheet_names) > 1)
            self._sheet_combo.blockSignals(False)
        elif self._sheet_combo.count() == 0:
            ext = self._file_path.suffix.lower() if self._file_path else ""
            if ext == ".csv":
                self._sheet_combo.blockSignals(True)
                self._sheet_combo.addItem("(csv)")
                self._sheet_combo.setEnabled(False)
                self._sheet_combo.blockSignals(False)

        self._header_row_spin.blockSignals(True)
        self._header_row_spin.setValue(header_row)
        self._header_row_spin.blockSignals(False)

        df_full.columns = [str(c) for c in df_full.columns]
        self._df = df_full
        self._all_columns = list(df_full.columns)
        df_preview.columns = self._all_columns[:len(df_preview.columns)]

        self._file_info_label.setText(f"{total_rows} linhas · {len(self._all_columns)} colunas")
        self._file_info_label.setStyleSheet(f"color: {self.COLORS['muted']};")

        preview_df = df_preview.head(self.PREVIEW_ROWS)
        self._preview_table.setUpdatesEnabled(False)
        self._preview_table.setColumnCount(len(self._all_columns))
        self._preview_table.setHorizontalHeaderLabels(self._all_columns)
        self._preview_table.setRowCount(len(preview_df))
        for r, (_, row) in enumerate(preview_df.iterrows()):
            for c, col in enumerate(self._all_columns):
                v = row[col]
                text = "" if (v is None or (isinstance(v, float) and str(v) == "nan")) else str(v)
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                self._preview_table.setItem(r, c, item)
        self._preview_table.setUpdatesEnabled(True)

        self._update_nav()

    def _on_file_load_failed(self, error: str) -> None:
        self._load_bar_widget.setVisible(False)
        self._file_info_label.setText(f"Erro: {error}")
        self._file_info_label.setStyleSheet(f"color: {self.COLORS['bad']};")
        self._update_nav()

    def _auto_detect_header(self) -> None:
        """Testa linhas 0-4 e escolhe a que tem menos 'Unnamed:' como cabeçalho."""
        if self._file_path is None:
            return
        ext = self._file_path.suffix.lower()
        sheet = self._sheet_combo.currentText() if self._sheet_combo.isEnabled() else 0
        if sheet == "(csv)":
            sheet = 0
        best_row = 0
        best_score = -1
        for row in range(5):
            try:
                if ext == ".csv":
                    df = pd.read_csv(str(self._file_path), header=row, nrows=0, dtype=str)
                else:
                    df = pd.read_excel(str(self._file_path), sheet_name=sheet, header=row, nrows=0, dtype=str)
                total = len(df.columns)
                if total == 0:
                    continue
                named = sum(1 for c in df.columns if not str(c).startswith("Unnamed:"))
                score = named / total
                if score > best_score:
                    best_score = score
                    best_row = row
                if score >= 0.9:
                    break
            except Exception:
                break
        self._header_row_spin.blockSignals(True)
        self._header_row_spin.setValue(best_row)
        self._header_row_spin.blockSignals(False)

    def _refresh_preview(self) -> None:
        if self._file_path is None:
            return
        self._start_file_load(read_sheets=False, detect_header=False)

    def _read_file(self, nrows: int | None = None) -> Any:
        assert self._file_path is not None
        ext = self._file_path.suffix.lower()
        header_row = self._header_row_spin.value() if self._header_check.isChecked() else None
        sheet = self._sheet_combo.currentText() if self._sheet_combo.isEnabled() else 0
        if sheet == "(csv)":
            sheet = 0
        if ext == ".csv":
            return pd.read_csv(
                str(self._file_path), header=header_row, nrows=nrows,
                dtype=str, keep_default_na=False,
            )
        return pd.read_excel(
            str(self._file_path), sheet_name=sheet, header=header_row,
            nrows=nrows, dtype=str,
        )

    def _count_rows(self) -> int:
        try:
            full = self._read_file()
            return len(full)
        except Exception:
            return 0

    # ── Passo 2: Correlacionar Campos ─────────────────────────────────────────

    def _build_step_mapeamento(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet(f"background: {self.COLORS['bg']};")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 12, 24, 12)
        layout.setSpacing(10)

        # Instrução
        info = QLabel(
            "Para cada coluna da planilha, selecione qual campo do cadastro corresponde. "
            "Colunas marcadas como '-- Ignorar --' não serão importadas."
        )
        info.setWordWrap(True)
        info.setObjectName("muted")
        layout.addWidget(info)

        # Botões de ação do mapeamento
        actions_bar = QWidget()
        actions_bar.setStyleSheet(f"background: {self.COLORS['bg']};")
        actions_layout = QHBoxLayout(actions_bar)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)

        btn_auto = QPushButton("Sugerir Automaticamente")
        btn_auto.clicked.connect(self._apply_auto_mapping)
        actions_layout.addWidget(btn_auto)

        actions_layout.addSpacing(16)

        lbl_perfil = QLabel("Perfil:")
        lbl_perfil.setStyleSheet(f"color: {self.COLORS['text']}; background: transparent;")
        actions_layout.addWidget(lbl_perfil)

        self._profile_combo = _NoScrollCombo()
        self._profile_combo.setMinimumWidth(180)
        self._profile_combo.setMaximumWidth(260)
        self._profile_combo.setStyleSheet(
            "QComboBox { background: #ffffff; color: #1d2730; border: 1px solid #d6e0e8; "
            "border-radius: 4px; padding: 2px 6px; }"
            "QComboBox QAbstractItemView { background: #ffffff; color: #1d2730; }"
        )
        actions_layout.addWidget(self._profile_combo)

        btn_load_profile = QPushButton("Carregar")
        btn_load_profile.clicked.connect(self._load_profile)
        actions_layout.addWidget(btn_load_profile)

        btn_save_profile = QPushButton("Salvar perfil…")
        btn_save_profile.clicked.connect(self._save_profile_dialog)
        actions_layout.addWidget(btn_save_profile)

        btn_del_profile = QPushButton("Excluir")
        btn_del_profile.clicked.connect(self._delete_profile)
        actions_layout.addWidget(btn_del_profile)

        actions_layout.addStretch()
        layout.addWidget(actions_bar)

        # Scroll com grid de mapeamento
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet(
            "QScrollArea { background: #ffffff; border: 1px solid #d6e0e8; border-radius: 4px; }"
            "QScrollArea > QWidget > QWidget { background: #ffffff; }"
        )
        self._map_container = QWidget()
        self._map_container.setStyleSheet("background: #ffffff;")
        self._map_grid = QGridLayout(self._map_container)
        self._map_grid.setContentsMargins(12, 8, 12, 8)
        self._map_grid.setHorizontalSpacing(16)
        self._map_grid.setVerticalSpacing(6)
        # Cabeçalhos fixos
        for col, text in enumerate(["Coluna da Planilha", "Amostra de Valor", "Campo Destino"]):
            h = QLabel(text)
            h.setStyleSheet(f"color: {self.COLORS['muted']}; font-weight: bold; background: transparent;")
            self._map_grid.addWidget(h, 0, col)
        self._map_grid.setColumnStretch(2, 1)
        scroll.setWidget(self._map_container)
        layout.addWidget(scroll, 1)

        return page

    def _build_campo_combo(self) -> QComboBox:
        combo = _NoScrollCombo()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMaxVisibleItems(22)
        combo.setStyleSheet(
            "QComboBox { background: #ffffff; color: #1d2730; border: 1px solid #d6e0e8; "
            "border-radius: 4px; padding: 2px 6px; }"
            "QComboBox::drop-down { border: none; }"
            "QComboBox QAbstractItemView { background: #ffffff; color: #1d2730; "
            "selection-background-color: #dbeafe; selection-color: #1d2730; border: 1px solid #d6e0e8; }"
        )
        if combo.lineEdit() is not None:
            combo.lineEdit().setStyleSheet(
                "background: #ffffff; color: #1d2730; border: none;"
            )
            combo.lineEdit().setPlaceholderText("Pesquisar campo...")

        combo.addItem("-- Ignorar --", IGNORAR)
        field_labels: list[str] = ["-- Ignorar --"]

        for grp_name, fields in CAMPOS_GRUPOS:
            combo.addItem(f"── {grp_name} ──", "__sep__")
            sep_idx = combo.count() - 1
            sep_item = combo.model().item(sep_idx)
            if sep_item is not None:
                sep_item.setEnabled(False)
                sep_item.setForeground(QColor(self.COLORS["muted"]))
            for key, label in fields:
                combo.addItem(label, key)
                field_labels.append(label)

        completer = QCompleter(field_labels, combo)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        combo.setCompleter(completer)

        # Quando o usuário seleciona do popup, sincroniza o índice do combo
        def _sync_index(text: str) -> None:
            i = combo.findText(text)
            if i >= 0:
                combo.setCurrentIndex(i)

        completer.activated.connect(_sync_index)
        return combo

    def _load_step_mapeamento(self) -> None:
        # Limpa grid (mantém cabeçalhos na linha 0)
        while self._map_grid.count() > 3:
            item = self._map_grid.takeAt(3)
            if item.widget():
                item.widget().deleteLater()
        self._mapping_combos.clear()

        if not self._all_columns:
            msg = QLabel("Nenhuma coluna encontrada. Verifique se a planilha foi carregada corretamente.")
            msg.setStyleSheet(f"color: {self.COLORS['bad']}; padding: 12px; background: transparent;")
            self._map_grid.addWidget(msg, 1, 0, 1, 3)
            return

        # Pega amostras (primeira linha não-vazia de cada coluna)
        samples: dict[str, str] = {}
        if self._df is not None:
            for col in self._all_columns:
                sample = ""
                for v in self._df[col]:
                    sv = str(v).strip()
                    if sv and sv.lower() not in {"nan", "none"}:
                        sample = sv[:40]
                        break
                samples[col] = sample

        for row_i, col in enumerate(self._all_columns, start=1):
            # Nome da coluna
            col_lbl = QLabel(col)
            col_lbl.setStyleSheet(f"color: {self.COLORS['text']}; background: transparent; font-weight: 500;")
            col_lbl.setMinimumWidth(160)
            col_lbl.setMaximumWidth(220)
            col_lbl.setWordWrap(True)
            self._map_grid.addWidget(col_lbl, row_i, 0)

            # Amostra
            sample_lbl = QLabel(samples.get(col, ""))
            sample_lbl.setStyleSheet(f"color: {self.COLORS['muted']}; background: transparent;")
            sample_lbl.setMinimumWidth(120)
            sample_lbl.setMaximumWidth(200)
            self._map_grid.addWidget(sample_lbl, row_i, 1)

            # Combo destino
            combo = self._build_campo_combo()
            combo.setMinimumWidth(260)
            self._map_grid.addWidget(combo, row_i, 2)
            self._mapping_combos[col] = combo

        self._refresh_profile_combo()

        # Auto-aplica "Último utilizado" se coincidir com ≥50% das colunas
        last = self._profiles.get("Último utilizado")
        if last and self._all_columns:
            matched = sum(1 for col in self._all_columns if col in last)
            if matched >= max(1, len(self._all_columns) * 0.5):
                self._apply_profile_data(last)
                return  # pula sugestão automática — perfil já está aplicado

        self._apply_auto_mapping()

    def _apply_auto_mapping(self) -> None:
        used: set[str] = set()
        for col, combo in self._mapping_combos.items():
            suggestion = _auto_map(col)
            if suggestion != IGNORAR and suggestion not in used:
                idx = combo.findData(suggestion)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
                    used.add(suggestion)
            else:
                combo.setCurrentIndex(0)  # Ignorar

    # ── Perfis de mapeamento ──────────────────────────────────────────────────

    def _refresh_profile_combo(self) -> None:
        self._profile_combo.clear()
        for name in self._profiles:
            self._profile_combo.addItem(name)

    def _apply_profile_data(self, profile: dict[str, str]) -> None:
        for col, combo in self._mapping_combos.items():
            campo = profile.get(col)
            if campo:
                idx = combo.findData(campo)
                combo.setCurrentIndex(idx if idx >= 0 else 0)
            else:
                combo.setCurrentIndex(0)

    def _load_profile(self) -> None:
        name = self._profile_combo.currentText()
        if not name or name not in self._profiles:
            return
        self._apply_profile_data(self._profiles[name])

    def _save_profile_dialog(self) -> None:
        from PySide6.QtWidgets import QInputDialog
        default = (
            self._profile_combo.currentText()
            or (self._file_path.stem if self._file_path else "Perfil")
        )
        name, ok = QInputDialog.getText(
            self, "Salvar Perfil de Mapeamento", "Nome do perfil:", text=default
        )
        if ok and name.strip():
            self._persist_profile(name.strip())

    def _persist_profile(self, name: str) -> None:
        profile: dict[str, str] = {}
        for col, combo in self._mapping_combos.items():
            key = combo.currentData()
            if key and key not in {IGNORAR, "__sep__"}:
                profile[col] = key
        self._profiles[name] = profile
        _save_all_profiles(self._profiles)
        self._refresh_profile_combo()
        idx = self._profile_combo.findText(name)
        if idx >= 0:
            self._profile_combo.setCurrentIndex(idx)

    def _delete_profile(self) -> None:
        name = self._profile_combo.currentText()
        if not name or name not in self._profiles:
            return
        from PySide6.QtWidgets import QMessageBox
        if (
            QMessageBox.question(self, "Excluir perfil", f"Excluir o perfil '{name}'?")
            == QMessageBox.Yes
        ):
            del self._profiles[name]
            _save_all_profiles(self._profiles)
            self._refresh_profile_combo()

    def _get_col_mapping(self) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for col, combo in self._mapping_combos.items():
            key = combo.currentData()
            if key and key not in {IGNORAR, "__sep__"}:
                mapping[key] = col
        return mapping

    # ── Passo 3: Importar ─────────────────────────────────────────────────────

    def _build_step_importar(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet(f"background: {self.COLORS['bg']};")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(10)

        self._import_status_label = QLabel("Pronto para importar.")
        self._import_status_label.setWordWrap(True)
        layout.addWidget(self._import_status_label)

        self._import_progress = QProgressBar()
        self._import_progress.setMinimum(0)
        self._import_progress.setMaximum(100)
        self._import_progress.setValue(0)
        self._import_progress.setFixedHeight(20)
        layout.addWidget(self._import_progress)

        # Abas: Log e Erros
        self._import_tabs = QTabWidget()
        self._import_tabs.setStyleSheet(
            "QTabWidget::pane { background: #ffffff; border: 1px solid #d6e0e8; border-radius: 4px; }"
            "QTabBar::tab { background: #edf3f7; color: #5d6e7d; padding: 5px 14px; "
            "border: 1px solid #d6e0e8; border-bottom: none; border-radius: 4px 4px 0 0; margin-right: 2px; }"
            "QTabBar::tab:selected { background: #ffffff; color: #1d2730; font-weight: bold; }"
        )

        self._import_log = QTextEdit()
        self._import_log.setReadOnly(True)
        self._import_log.setStyleSheet("background: #ffffff; border: none;")
        self._import_log.setPlaceholderText("Log da importação aparecerá aqui...")
        self._import_tabs.addTab(self._import_log, "Log")

        # Aba de erros com tabela
        err_page = QWidget()
        err_page.setStyleSheet("background: #ffffff;")
        err_layout = QVBoxLayout(err_page)
        err_layout.setContentsMargins(8, 8, 8, 8)
        err_layout.setSpacing(6)

        self._error_table = QTableWidget()
        self._error_table.setColumnCount(6)
        self._error_table.setHorizontalHeaderLabels(
            ["Linha", "Fornecedor", "Código", "EAN", "Descrição", "Erro"]
        )
        self._error_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._error_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._error_table.setAlternatingRowColors(True)
        self._error_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self._error_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self._error_table.setColumnWidth(0, 55)
        self._error_table.setColumnWidth(1, 160)
        self._error_table.setColumnWidth(2, 100)
        self._error_table.setColumnWidth(3, 110)
        self._error_table.setStyleSheet(
            "QTableWidget { background: #ffffff; border: none; }"
            "QHeaderView::section { background: #edf3f7; color: #1d2730; "
            "font-weight: bold; border: 1px solid #d6e0e8; padding: 4px; }"
        )
        err_layout.addWidget(self._error_table, 1)

        self._btn_export_errors = QPushButton("Exportar erros (CSV)…")
        self._btn_export_errors.setFixedWidth(190)
        self._btn_export_errors.clicked.connect(self._export_errors)
        self._btn_export_errors.hide()
        err_layout.addWidget(self._btn_export_errors, 0, Qt.AlignRight)

        self._import_tabs.addTab(err_page, "Erros (0)")

        # Aba de não atualizados (sem alteração + ignorados)
        skip_page = QWidget()
        skip_page.setStyleSheet("background: #ffffff;")
        skip_layout = QVBoxLayout(skip_page)
        skip_layout.setContentsMargins(8, 8, 8, 8)
        skip_layout.setSpacing(6)

        self._skip_table = QTableWidget()
        self._skip_table.setColumnCount(6)
        self._skip_table.setHorizontalHeaderLabels(
            ["Linha", "Fornecedor", "Código", "EAN", "Descrição", "Motivo"]
        )
        self._skip_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._skip_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._skip_table.setAlternatingRowColors(True)
        self._skip_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self._skip_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self._skip_table.setColumnWidth(0, 55)
        self._skip_table.setColumnWidth(1, 160)
        self._skip_table.setColumnWidth(2, 100)
        self._skip_table.setColumnWidth(3, 110)
        self._skip_table.setStyleSheet(
            "QTableWidget { background: #ffffff; border: none; }"
            "QHeaderView::section { background: #edf3f7; color: #1d2730; "
            "font-weight: bold; border: 1px solid #d6e0e8; padding: 4px; }"
        )
        skip_layout.addWidget(self._skip_table, 1)

        self._btn_export_skipped = QPushButton("Exportar não atualizados (CSV)…")
        self._btn_export_skipped.setFixedWidth(230)
        self._btn_export_skipped.clicked.connect(self._export_skipped)
        self._btn_export_skipped.hide()
        skip_layout.addWidget(self._btn_export_skipped, 0, Qt.AlignRight)

        self._import_tabs.addTab(skip_page, "Não Atualizados (0)")
        layout.addWidget(self._import_tabs, 1)

        self._import_summary = QLabel("")
        self._import_summary.setWordWrap(True)
        layout.addWidget(self._import_summary)

        return page

    def _prepare_import_step(self) -> None:
        mapping = self._get_col_mapping()
        mapped_count = len(mapping)
        total_rows = len(self._df) if self._df is not None else 0

        lines = [f"<b>Arquivo:</b> {self._file_path.name if self._file_path else '?'}"]
        lines.append(f"<b>Linhas a processar:</b> {total_rows}")
        lines.append(f"<b>Campos mapeados:</b> {mapped_count}")
        lines.append("")

        if "fornecedor_nome" not in mapping:
            lines.append(
                "<span style='color:#b42318'>⚠ Campo <b>Fornecedor - Nome</b> não mapeado. "
                "Linhas sem fornecedor serão ignoradas.</span>"
            )
        if "codigo_fornecedor" not in mapping and "ean" not in mapping:
            lines.append(
                "<span style='color:#b56b12'>⚠ Nem <b>Cód. Fornecedor</b> nem <b>EAN</b> "
                "foram mapeados. Produtos podem ser duplicados.</span>"
            )

        lines.append("")
        lines.append("Clique em <b>Importar</b> para iniciar.")
        self._import_status_label.setText("<br>".join(lines))
        self._import_progress.setValue(0)
        self._import_log.clear()
        self._import_summary.clear()

    # ── Navegação ─────────────────────────────────────────────────────────────

    def _go_next(self) -> None:
        step = self._stack.currentIndex()
        if step == 0:
            if self._df is None:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "Aguarde", "O arquivo ainda esta sendo carregado. Aguarde e tente novamente.")
                return
            self._stack.setCurrentIndex(1)
            try:
                self._load_step_mapeamento()
            except Exception as exc:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.critical(self, "Erro ao carregar mapeamento", str(exc))
        elif step == 1:
            self._stack.setCurrentIndex(2)
            self._prepare_import_step()
        self._update_nav()

    def _go_back(self) -> None:
        step = self._stack.currentIndex()
        if step > 0:
            if step == 2:
                # Reseta o passo 3 para permitir nova importação com mapeamento corrigido
                self._import_progress.setValue(0)
                self._import_log.clear()
                self._import_summary.clear()
                self._error_table.setRowCount(0)
                self._import_tabs.setTabText(1, "Erros (0)")
                self._skip_table.setRowCount(0)
                self._import_tabs.setTabText(2, "Não Atualizados (0)")
                self._import_tabs.setCurrentIndex(0)
                self._btn_export_errors.hide()
                self._btn_export_skipped.hide()
                self._import_status_label.setText("Pronto para importar.")
                self._btn_fechar.hide()
                self._btn_importar.show()
                self._btn_importar.setEnabled(True)
            self._stack.setCurrentIndex(step - 1)
            self._update_nav()

    # ── Importação ────────────────────────────────────────────────────────────

    def _start_import(self) -> None:
        if self._df is None:
            return

        self._btn_importar.setEnabled(False)
        self._btn_anterior.setEnabled(False)
        self._import_log.clear()
        self._import_progress.setValue(0)
        self._import_status_label.setText("Importando...")

        col_mapping = self._get_col_mapping()
        company_id = self._company_combo.currentData()

        self._thread = QThread(self)
        self._worker = _ImportWorker(
            self.repository,
            self.environment,
            company_id,
            col_mapping,
            self._df,
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_import_finished)
        self._worker.failed.connect(self._on_import_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._import_progress.setValue(int(current * 100 / total))
        self._import_status_label.setText(msg)

    def _on_import_finished(self, stats: dict) -> None:
        self._import_progress.setValue(100)
        self._btn_fechar.show()
        self._btn_importar.hide()
        self._btn_anterior.setEnabled(True)

        try:
            self._persist_profile("Último utilizado")
        except Exception:
            pass

        ok_color  = self.COLORS["ok"]
        bad_color = self.COLORS["bad"]

        erros = stats["erros"]
        n_erros = len(erros)

        summary_parts = [
            f"<b style='color:{ok_color}'>Importação concluída!</b>",
            "",
            f"<b>Total de linhas:</b> {stats['linhas']}",
            f"<b>Novos produtos:</b> {stats['novos']}",
            f"<b>Atualizados:</b> {stats['atualizados']}",
            f"<b>Sem alteração:</b> {stats.get('sem_alteracao', 0)}",
            f"<b>Ignorados:</b> {stats['ignorados']}",
            f"<b>Empresas criadas:</b> {stats['empresas_criadas']}",
            f"<b>Fornecedores criados:</b> {stats['fornecedores_criados']}",
            f"<b>Classificações criadas:</b> {stats['tipos_criados']}",
        ]
        if n_erros:
            summary_parts.append(
                f"<b style='color:{bad_color}'>Erros: {n_erros}</b>"
            )
        self._import_summary.setText("<br>".join(summary_parts))

        # Popula tabela de erros
        self._error_table.setRowCount(0)
        if n_erros:
            self._error_table.setRowCount(n_erros)
            for r, e in enumerate(erros):
                if isinstance(e, dict):
                    vals = [
                        str(e.get("linha", "")),
                        e.get("fornecedor", ""),
                        e.get("codigo", ""),
                        e.get("ean", ""),
                        e.get("descricao", ""),
                        e.get("erro", ""),
                    ]
                else:
                    vals = ["", "", "", "", "", str(e)]
                for c, v in enumerate(vals):
                    item = QTableWidgetItem(v)
                    if c == 5:  # coluna Erro em vermelho
                        item.setForeground(QColor(self.COLORS["bad"]))
                    self._error_table.setItem(r, c, item)

            self._import_tabs.setTabText(1, f"Erros ({n_erros})")
            self._import_tabs.setCurrentIndex(1)
            self._btn_export_errors.show()
        else:
            self._import_tabs.setTabText(1, "Erros (0)")

        # Popula aba de não atualizados
        nao_atualizados = stats.get("nao_atualizados", [])
        n_skip = len(nao_atualizados)
        self._skip_table.setRowCount(0)
        if n_skip:
            self._skip_table.setRowCount(n_skip)
            warn_color = QColor("#b56b12")
            muted_color = QColor("#5d6e7d")
            for r, s in enumerate(nao_atualizados):
                vals = [
                    str(s.get("linha", "")),
                    s.get("fornecedor", ""),
                    s.get("codigo", ""),
                    s.get("ean", ""),
                    s.get("descricao", ""),
                    s.get("motivo", ""),
                ]
                for c, v in enumerate(vals):
                    item = QTableWidgetItem(v)
                    if c == 5:
                        color = warn_color if "Ignorado" in v or "Sem " in v else muted_color
                        item.setForeground(color)
                    self._skip_table.setItem(r, c, item)
            self._import_tabs.setTabText(2, f"Não Atualizados ({n_skip})")
            self._btn_export_skipped.show()
        else:
            self._import_tabs.setTabText(2, "Não Atualizados (0)")

        sem_alt = stats.get("sem_alteracao", 0)
        self._import_status_label.setText(
            f"<b style='color:{ok_color}'>Pronto:</b> "
            f"{stats['novos']} novos · {stats['atualizados']} atualizados · "
            f"{sem_alt} sem alteração · {stats['ignorados']} ignorados"
            + (f" · <b style='color:{bad_color}'>{n_erros} erros</b>" if n_erros else "")
        )

    def _export_errors(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar erros", "erros_importacao.csv",
            "CSV (*.csv);;Excel (*.xlsx)"
        )
        if not path:
            return
        rows = []
        for r in range(self._error_table.rowCount()):
            rows.append({
                "Linha":      self._error_table.item(r, 0).text() if self._error_table.item(r, 0) else "",
                "Fornecedor": self._error_table.item(r, 1).text() if self._error_table.item(r, 1) else "",
                "Código":     self._error_table.item(r, 2).text() if self._error_table.item(r, 2) else "",
                "EAN":        self._error_table.item(r, 3).text() if self._error_table.item(r, 3) else "",
                "Descrição":  self._error_table.item(r, 4).text() if self._error_table.item(r, 4) else "",
                "Erro":       self._error_table.item(r, 5).text() if self._error_table.item(r, 5) else "",
            })
        try:
            df_err = pd.DataFrame(rows)
            if path.endswith(".xlsx"):
                df_err.to_excel(path, index=False)
            else:
                df_err.to_csv(path, index=False, encoding="utf-8-sig")
            from PySide6.QtWidgets import QMessageBox
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportado")
            msg.setText(f"Arquivo salvo em:\n{path}\n\nDeseja abrir o arquivo?")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                import os
                os.startfile(path)
        except Exception as exc:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Erro ao exportar", str(exc))

    def _export_skipped(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar não atualizados", "nao_atualizados_importacao.csv",
            "CSV (*.csv);;Excel (*.xlsx)"
        )
        if not path:
            return
        rows = []
        for r in range(self._skip_table.rowCount()):
            rows.append({
                "Linha":      self._skip_table.item(r, 0).text() if self._skip_table.item(r, 0) else "",
                "Fornecedor": self._skip_table.item(r, 1).text() if self._skip_table.item(r, 1) else "",
                "Código":     self._skip_table.item(r, 2).text() if self._skip_table.item(r, 2) else "",
                "EAN":        self._skip_table.item(r, 3).text() if self._skip_table.item(r, 3) else "",
                "Descrição":  self._skip_table.item(r, 4).text() if self._skip_table.item(r, 4) else "",
                "Motivo":     self._skip_table.item(r, 5).text() if self._skip_table.item(r, 5) else "",
            })
        try:
            df_skip = pd.DataFrame(rows)
            if path.endswith(".xlsx"):
                df_skip.to_excel(path, index=False)
            else:
                df_skip.to_csv(path, index=False, encoding="utf-8-sig")
            from PySide6.QtWidgets import QMessageBox
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportado")
            msg.setText(f"Arquivo salvo em:\n{path}\n\nDeseja abrir o arquivo?")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                import os
                os.startfile(path)
        except Exception as exc:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Erro ao exportar", str(exc))

    def _on_import_failed(self, msg: str) -> None:
        self._import_progress.setValue(0)
        self._import_status_label.setText(
            f"<b style='color:{self.COLORS['bad']}'>Falha na importação:</b> {msg}"
        )
        self._import_log.append(f"ERRO FATAL: {msg}")
        self._btn_importar.setEnabled(True)
        self._btn_anterior.setEnabled(True)
