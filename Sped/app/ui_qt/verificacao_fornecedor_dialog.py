"""
Verificação de Produtos por Fornecedor — verificação em lote.

Compara todos os produtos de todos os participantes/fornecedores encontrados
no SPED contra o catálogo da empresa. Mostra o que está e o que não está
cadastrado por fornecedor, com filtro e exportação.
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QBrush, QColor, QPalette
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStyle,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.repositories.mysql_cadastro import MysqlCadastroRepository

# ── Colunas ───────────────────────────────────────────────────────────────────

_MAIN_HEADERS = [
    "Status",
    "Fornecedor SPED",
    "Fornecedor Cadastro",
    "Código",
    "Descrição SPED",
    "NCM",
    "CST ICMS",
    "CFOP",
    "Alíq. ICMS %",
    "Período(s)",
    "Qtd. Docs",
]
_MAIN_FIELDS = [
    "status",
    "participant_name",
    "supplier_db_name",
    "code",
    "description",
    "ncm",
    "cst_icms",
    "cfop",
    "icms_rate",
    "periodos_str",
    "qtd_docs",
]

_EXPORT_EXTRA_HEADERS = [
    "Descrição Cadastro",
    "NCM Cadastro",
    "CST ICMS Cadastro",
    "CFOP Cadastro",
    "Alíq. ICMS Cadastro",
]
_EXPORT_EXTRA_FIELDS = [
    "descricao_cad",
    "ncm_cad",
    "cst_cad",
    "cfop_cad",
    "aliq_cad",
]

# ── Cores ─────────────────────────────────────────────────────────────────────

_BG_NAO   = QColor("#ffeb9c")   # amarelo — não cadastrado
_FG_NAO   = QColor("#7d4a00")
_BG_OK    = QColor("#c6efce")   # verde — já cadastrado
_FG_OK    = QColor("#276221")
_BG_SEM   = QColor("#d9e8f5")   # azul claro — sem correspondência no cadastro
_FG_SEM   = QColor("#1a4a7a")
_ROW_EVEN = QColor("#f3f6f9")
_ROW_ODD  = QColor("#ffffff")

_STATUS_NAO = "Não Cadastrado"
_STATUS_OK  = "Já Cadastrado"
_STATUS_SEM = "Sem Correspondência"


def _status_colors(status: str) -> tuple[QColor, QColor]:
    if status == _STATUS_OK:
        return _BG_OK, _FG_OK
    if status == _STATUS_SEM:
        return _BG_SEM, _FG_SEM
    return _BG_NAO, _FG_NAO


def _db_aliq_pct(v: object) -> str:
    """Converte fração de banco (0.18) para percentual display (18,00)."""
    s = str(v or "").strip()
    if not s:
        return ""
    try:
        result = round(Decimal(s.replace(",", ".")) * 100, 2)
        return f"{result:.2f}".replace(".", ",")
    except (InvalidOperation, Exception):
        return s


def _best_match_supplier(participant_name: str, suppliers: list[dict]) -> dict | None:
    """Retorna o fornecedor do cadastro com nome mais próximo do participante SPED."""
    pname_lower = participant_name.lower().strip()
    pname_words = set(pname_lower.split())
    best: dict | None = None
    best_score = 0
    for s in suppliers:
        sname = str(s.get("nome") or "").lower().strip()
        if not sname:
            continue
        if sname == pname_lower:
            return s
        swords = set(sname.split())
        score = len(pname_words & swords)
        if pname_lower in sname or sname in pname_lower:
            score += 5
        if score > best_score:
            best_score = score
            best = s
    return best if best_score >= 2 else None


# ── Delegate de pintura ───────────────────────────────────────────────────────

class _BrushDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index) -> None:
        painter.save()
        is_selected = bool(option.state & QStyle.State_Selected)
        bg = index.data(Qt.BackgroundRole)
        if is_selected:
            painter.fillRect(option.rect, option.palette.color(QPalette.Highlight))
        elif bg is not None:
            color = bg.color() if isinstance(bg, QBrush) else bg
            painter.fillRect(option.rect, color)
        if is_selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            fg = index.data(Qt.ForegroundRole)
            if fg is not None:
                painter.setPen(fg.color() if isinstance(fg, QBrush) else fg)
            else:
                painter.setPen(option.palette.color(QPalette.Text))
        font = index.data(Qt.FontRole)
        if font is not None:
            painter.setFont(font)
        painter.drawText(
            option.rect.adjusted(4, 0, -4, 0),
            Qt.AlignVCenter | Qt.AlignLeft,
            str(index.data(Qt.DisplayRole) or ""),
        )
        painter.restore()

    def sizeHint(self, option, index):
        hint = super().sizeHint(option, index)
        hint.setHeight(max(hint.height(), 22))
        return hint


# ── Worker: verificação em lote de todos os fornecedores ─────────────────────

class _VerificarTodosWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(list)
    failed   = Signal(str)

    def __init__(
        self,
        rows: list[dict],
        operation_type: str,
        empresa_id: int,
        repository: MysqlCadastroRepository,
    ) -> None:
        super().__init__()
        self._rows           = rows
        self._operation_type = operation_type
        self._empresa_id     = empresa_id
        self._repo           = repository

    def run(self) -> None:
        try:
            # Campos de catálogo dependem da operação
            if self._operation_type == "Entrada":
                cst_f  = "cst_icms"
                cfop_f = "cfop_entrada"
                aliq_f = "aliquota_icms"
            else:
                cst_f  = "cst_icms_saida"
                cfop_f = "cfop_saida_empresa"
                aliq_f = "aliquota_icms_saida"

            # 1. Agrupa produtos do SPED por participante
            by_participant: dict[str, dict[str, dict]] = {}
            for row in self._rows:
                period = str(row.get("period") or "")
                for detail in row.get("launch_details", []):
                    participant = str(detail.get("participant_name") or "").strip()
                    if not participant:
                        continue
                    code = str(detail.get("code") or "").strip()
                    if not code:
                        continue
                    if participant not in by_participant:
                        by_participant[participant] = {}
                    prods = by_participant[participant]
                    if code not in prods:
                        prods[code] = {
                            "code":            code,
                            "description":     str(detail.get("description") or "").strip(),
                            "ncm":             str(detail.get("ncm") or "").strip(),
                            "cst_icms":        str(detail.get("cst_icms") or "").strip(),
                            "cfop":            str(detail.get("cfop") or "").strip(),
                            "icms_rate":       str(detail.get("icms_rate") or "").strip(),
                            "aliquota_pis":    str(detail.get("aliquota_pis") or "").strip(),
                            "aliquota_cofins": str(detail.get("aliquota_cofins") or "").strip(),
                            "periodos":        set(),
                            "qtd_docs":        0,
                            "document_keys":   [],
                        }
                    prod = prods[code]
                    if period:
                        prod["periodos"].add(period)
                    prod["qtd_docs"] += 1
                    if detail.get("document_key"):
                        key = str(detail["document_key"]).strip()
                        if key and key not in prod["document_keys"]:
                            prod["document_keys"].append(key)

            # 2. Carrega todos os fornecedores da empresa
            all_suppliers = self._repo.list_suppliers(self._empresa_id)

            # 3. Para cada participante, encontra o fornecedor e compara produtos
            all_rows: list[dict] = []
            participants = list(by_participant.keys())
            total = len(participants)

            catalog_cache: dict[int, tuple[dict, dict]] = {}

            for idx, participant in enumerate(participants):
                self.progress.emit(idx, total, f"Verificando: {participant[:50]}...")

                supplier = _best_match_supplier(participant, all_suppliers)

                sped_prods = by_participant[participant]

                if supplier is None:
                    # Sem correspondência no cadastro
                    for code, prod in sped_prods.items():
                        periodos = sorted(prod["periodos"])
                        all_rows.append({
                            **prod,
                            "participant_name":  participant,
                            "supplier_id":       None,
                            "supplier_db_name":  _STATUS_SEM,
                            "status":            _STATUS_SEM,
                            "periodos_str":      ", ".join(periodos),
                            "descricao_cad":     "",
                            "ncm_cad":           "",
                            "cst_cad":           "",
                            "cfop_cad":          "",
                            "aliq_cad":          "",
                        })
                    continue

                supplier_id = int(supplier["id"])
                supplier_name = str(supplier.get("nome") or "")

                # Carrega catálogo do fornecedor (com cache)
                if supplier_id not in catalog_cache:
                    catalog = self._repo.list_supplier_products(supplier_id)
                    cat_by_code: dict[str, dict] = {}
                    cat_by_ean:  dict[str, dict] = {}
                    for p in catalog:
                        for kf in ("codigo_empresa", "codigo_fornecedor"):
                            k = str(p.get(kf) or "").strip()
                            if k and k not in cat_by_code:
                                cat_by_code[k] = p
                        ean = "".join(c for c in str(p.get("ean") or "") if c.isdigit())
                        if ean and ean not in cat_by_ean:
                            cat_by_ean[ean] = p
                    catalog_cache[supplier_id] = (cat_by_code, cat_by_ean)
                else:
                    cat_by_code, cat_by_ean = catalog_cache[supplier_id]

                for code, prod in sped_prods.items():
                    cat = cat_by_code.get(code)
                    if cat is None:
                        digits = "".join(c for c in code if c.isdigit())
                        cat = cat_by_ean.get(digits) if digits else None

                    periodos = sorted(prod["periodos"])
                    base = {
                        **prod,
                        "participant_name": participant,
                        "supplier_id":      supplier_id,
                        "supplier_db_name": supplier_name,
                        "periodos_str":     ", ".join(periodos),
                    }

                    if cat is None:
                        base["status"]       = _STATUS_NAO
                        base["descricao_cad"] = ""
                        base["ncm_cad"]      = ""
                        base["cst_cad"]      = ""
                        base["cfop_cad"]     = ""
                        base["aliq_cad"]     = ""
                    else:
                        base["status"]        = _STATUS_OK
                        base["descricao_cad"] = str(cat.get("descricao") or "")
                        base["ncm_cad"]       = str(cat.get("ncm") or "")
                        base["cst_cad"]       = str(cat.get(cst_f) or "")
                        base["cfop_cad"]      = str(cat.get(cfop_f) or "")
                        base["aliq_cad"]      = _db_aliq_pct(cat.get(aliq_f))

                    all_rows.append(base)

            self.progress.emit(total, total, "Concluído.")
            self.finished.emit(all_rows)
        except Exception as exc:
            self.failed.emit(str(exc))


# ── Worker: cadastrar produtos para múltiplos fornecedores ───────────────────

class _CadastrarTodosWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(dict)
    failed   = Signal(str)

    def __init__(
        self,
        produtos: list[dict],
        operation_type: str,
        repository: MysqlCadastroRepository,
    ) -> None:
        super().__init__()
        self._produtos       = produtos
        self._operation_type = operation_type
        self._repo           = repository

    def run(self) -> None:
        try:
            stats = {"cadastrados": 0, "ignorados": 0, "erros": []}
            total = len(self._produtos)

            for idx, prod in enumerate(self._produtos):
                supplier_id = prod.get("supplier_id")
                code        = str(prod.get("code") or "").strip()
                self.progress.emit(idx, total, f"Cadastrando {idx + 1}/{total} — {code}...")

                if not code or supplier_id is None:
                    stats["ignorados"] += 1
                    continue

                desc     = str(prod.get("description") or "").strip()
                ncm      = str(prod.get("ncm") or "").strip()
                cst      = str(prod.get("cst_icms") or "").strip()
                cfop     = str(prod.get("cfop") or "").strip()
                aliq_pct = str(prod.get("icms_rate") or "").strip()

                try:
                    aliq_frac = str(round(
                        Decimal(aliq_pct.replace(",", ".")) / 100, 4
                    )) if aliq_pct else "0"
                except InvalidOperation:
                    aliq_frac = "0"

                data: dict = {
                    "codigo_fornecedor": code,
                    "codigo_empresa":    code,
                    "ean":               code,
                    "descricao":         desc,
                    "ncm":              ncm,
                    "status_produto":    "ATIVO",
                }
                if self._operation_type == "Entrada":
                    data["cst_icms"]      = cst
                    data["cfop_entrada"]  = cfop
                    data["aliquota_icms"] = aliq_frac
                else:
                    data["cst_icms_saida"]     = cst
                    data["cfop_saida_empresa"]  = cfop
                    data["aliquota_icms_saida"] = aliq_frac

                try:
                    self._repo.save_supplier_product(supplier_id, data)
                    stats["cadastrados"] += 1
                except ValueError as exc:
                    if "Ja existe produto com o mesmo Codigo do Fornecedor" in str(exc):
                        existing = self._repo.fetch_supplier_product_by_fornecedor_code(
                            supplier_id, code
                        )
                        if existing and not str(existing.get("codigo_empresa") or "").strip():
                            data["id"] = existing["id"]
                            try:
                                self._repo.save_supplier_product(supplier_id, data)
                                stats["cadastrados"] += 1
                            except Exception as exc2:
                                stats["erros"].append({"code": code, "erro": str(exc2)})
                        else:
                            stats["erros"].append({
                                "code": code,
                                "erro": "Produto já existe para este fornecedor.",
                            })
                    else:
                        stats["erros"].append({"code": code, "erro": str(exc)})
                except Exception as exc:
                    stats["erros"].append({"code": code, "erro": str(exc)})

            self.progress.emit(total, total, "Concluído.")
            self.finished.emit(stats)
        except Exception as exc:
            self.failed.emit(str(exc))


# ── Diálogo principal ─────────────────────────────────────────────────────────

class VerificacaoFornecedorDialog(QDialog):
    """
    Verifica todos os fornecedores/participantes do SPED de uma só vez,
    mostrando quais produtos estão e quais não estão cadastrados no catálogo.
    Permite filtrar por fornecedor/status e exportar para Excel.
    """

    def __init__(
        self,
        rows: list[dict],
        operation_type: str,
        repository: MysqlCadastroRepository,
        environment: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._rows           = rows or []
        self._operation_type = operation_type
        self._repo           = repository
        self._env            = environment
        self._empresa_id: int | None = None
        self._all_rows: list[dict] = []       # todos os resultados
        self._filter_status    = "Todos"
        self._filter_supplier  = ""           # "" = todos
        self._filter_btns: dict[str, QPushButton] = {}
        self._v_thread: QThread | None = None
        self._v_worker: _VerificarTodosWorker | None = None
        self._c_thread: QThread | None = None
        self._c_worker: _CadastrarTodosWorker | None = None

        participante = "Fornecedor" if operation_type == "Entrada" else "Cliente"

        self.setWindowTitle(f"Verificação de Produtos por {participante} — {operation_type}s")
        self.setMinimumSize(1200, 700)
        if parent:
            self.setStyleSheet(parent.styleSheet())
        self.resize(1450, 850)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        title = QLabel(f"Verificação de Produtos por {participante} — {operation_type}s")
        title.setObjectName("sectionTitle")
        root.addWidget(title)

        # ── Linha 1: empresa + verificar ──────────────────────────────────────
        bar1 = QHBoxLayout()
        bar1.addWidget(QLabel("Empresa:"))
        self._company_combo = QComboBox()
        self._company_combo.setMinimumWidth(280)
        self._load_companies()
        bar1.addWidget(self._company_combo, 1)

        self._btn_verificar = QPushButton("Verificar Todos")
        self._btn_verificar.setObjectName("primaryButton")
        self._btn_verificar.clicked.connect(self._run_verificar)
        bar1.addWidget(self._btn_verificar)
        bar1.addStretch()
        root.addLayout(bar1)

        # ── Métricas ───────────────────────────────────────────────────────────
        metrics_row = QHBoxLayout()
        self._metric_labels: dict[str, QLabel] = {}
        for key, label in (
            ("fornecedores", f"Fornecedores no SPED"),
            ("total",        "Total Produtos"),
            ("nao_cad",      "Não Cadastrados"),
            ("ja_cad",       "Já Cadastrados"),
            ("sem_corr",     "Sem Correspondência"),
        ):
            card = QWidget()
            card.setObjectName("metricCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 6, 10, 6)
            cl.setSpacing(2)
            cl.addWidget(QLabel(label))
            val = QLabel("-")
            val.setObjectName("metricValue")
            cl.addWidget(val)
            self._metric_labels[key] = val
            metrics_row.addWidget(card)
        metrics_row.addStretch()
        root.addLayout(metrics_row)

        self._status_label = QLabel("")
        self._status_label.setObjectName("muted")
        root.addWidget(self._status_label)

        # ── Filtros de status ─────────────────────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.addWidget(QLabel("Filtrar:"))
        _filter_styles = {
            "Todos":              "QPushButton:checked { background-color: #2563eb; color: white; font-weight: bold; }",
            _STATUS_NAO:          "QPushButton:checked { background-color: #ffeb9c; color: #7d4a00; font-weight: bold; }",
            _STATUS_OK:           "QPushButton:checked { background-color: #c6efce; color: #276221; font-weight: bold; }",
            _STATUS_SEM:          "QPushButton:checked { background-color: #d9e8f5; color: #1a4a7a; font-weight: bold; }",
        }
        for mode, css in _filter_styles.items():
            btn = QPushButton(mode)
            btn.setCheckable(True)
            btn.setChecked(mode == "Todos")
            btn.setStyleSheet(css)
            btn.clicked.connect(lambda _checked, m=mode: self._set_status_filter(m))
            self._filter_btns[mode] = btn
            filter_bar.addWidget(btn)

        filter_bar.addSpacing(20)
        filter_bar.addWidget(QLabel(f"{participante}:"))
        self._supplier_filter_combo = QComboBox()
        self._supplier_filter_combo.setMinimumWidth(260)
        self._supplier_filter_combo.addItem(f"Todos os {participante}s", "")
        self._supplier_filter_combo.currentIndexChanged.connect(self._on_supplier_filter_changed)
        filter_bar.addWidget(self._supplier_filter_combo, 1)
        filter_bar.addStretch()
        root.addLayout(filter_bar)

        # ── Tabela principal ───────────────────────────────────────────────────
        self._table = self._make_table(_MAIN_HEADERS)
        root.addWidget(self._table, 1)

        # ── Barra inferior ─────────────────────────────────────────────────────
        bottom = QHBoxLayout()

        self._btn_cadastrar = QPushButton("Cadastrar Não Cadastrados (visíveis)")
        self._btn_cadastrar.setObjectName("primaryButton")
        self._btn_cadastrar.setEnabled(False)
        self._btn_cadastrar.clicked.connect(self._cadastrar)
        bottom.addWidget(self._btn_cadastrar)

        self._btn_export = QPushButton("Exportar Excel")
        self._btn_export.setEnabled(False)
        self._btn_export.clicked.connect(self._export)
        bottom.addWidget(self._btn_export)

        self._btn_atualizar_codigo = QPushButton("Atualizar Código Forn. (XML)")
        self._btn_atualizar_codigo.setEnabled(False)
        self._btn_atualizar_codigo.clicked.connect(self._abrir_atualizacao_codigo)
        bottom.addWidget(self._btn_atualizar_codigo)

        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.hide()
        bottom.addWidget(self._progress, 1)

        self._progress_label = QLabel("")
        self._progress_label.hide()
        bottom.addWidget(self._progress_label)

        bottom.addStretch()
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(self.reject)
        bottom.addWidget(btn_fechar)
        root.addLayout(bottom)

    # ── Helpers de UI ─────────────────────────────────────────────────────────

    def _make_table(self, headers: list[str]) -> QTableWidget:
        t = QTableWidget()
        t.setColumnCount(len(headers))
        t.setHorizontalHeaderLabels(headers)
        t.setAlternatingRowColors(False)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.verticalHeader().setVisible(False)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        t.horizontalHeader().setStretchLastSection(True)
        t.setItemDelegate(_BrushDelegate(t))
        return t

    def _load_companies(self) -> None:
        try:
            companies = self._repo.list_companies_for_confronto(self._env)
            for c in companies:
                label = str(c["nome"] or "").strip() or str(c["cnpj"] or "")
                self._company_combo.addItem(label, str(c["cnpj"] or ""))
            if len(companies) == 1:
                self._company_combo.setCurrentIndex(0)
        except Exception:
            pass

    def _resolve_empresa_id(self) -> int | None:
        cnpj = str(self._company_combo.currentData() or "").strip()
        name = self._company_combo.currentText()
        if not cnpj:
            return None
        try:
            return self._repo.find_company_id(self._env, name, cnpj)
        except Exception:
            return None

    # ── Verificação em lote ───────────────────────────────────────────────────

    def _run_verificar(self) -> None:
        if not self._rows:
            QMessageBox.warning(self, "Verificar", "Não há dados no SPED. Processe o SPED primeiro.")
            return

        empresa_id = self._resolve_empresa_id()
        if empresa_id is None:
            QMessageBox.warning(self, "Verificar", "Selecione uma empresa.")
            return

        self._empresa_id = empresa_id
        self._btn_verificar.setEnabled(False)
        self._btn_cadastrar.setEnabled(False)
        self._btn_export.setEnabled(False)
        self._progress.setValue(0)
        self._progress.show()
        self._progress_label.setText("Iniciando...")
        self._progress_label.show()
        self._status_label.setText("")

        self._v_thread = QThread(self)
        self._v_worker = _VerificarTodosWorker(
            self._rows, self._operation_type, empresa_id, self._repo
        )
        self._v_worker.moveToThread(self._v_thread)
        self._v_thread.started.connect(self._v_worker.run)
        self._v_worker.progress.connect(self._on_v_progress)
        self._v_worker.finished.connect(self._on_v_done)
        self._v_worker.failed.connect(self._on_v_failed)
        self._v_worker.finished.connect(self._v_thread.quit)
        self._v_worker.failed.connect(self._v_thread.quit)
        self._v_thread.finished.connect(self._v_worker.deleteLater)
        self._v_thread.finished.connect(self._v_thread.deleteLater)
        self._v_thread.finished.connect(lambda: setattr(self, "_v_thread", None))
        self._v_thread.finished.connect(lambda: setattr(self, "_v_worker", None))
        self._v_thread.start()

    def _on_v_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._progress.setValue(int(current / total * 100))
        self._progress_label.setText(msg)

    def _on_v_done(self, all_rows: list[dict]) -> None:
        self._progress.hide()
        self._progress_label.hide()
        self._btn_verificar.setEnabled(True)
        self._all_rows = all_rows

        # Atualiza combo de filtro por fornecedor
        self._supplier_filter_combo.blockSignals(True)
        self._supplier_filter_combo.clear()
        participante = "Fornecedor" if self._operation_type == "Entrada" else "Cliente"
        self._supplier_filter_combo.addItem(f"Todos os {participante}s", "")
        participants = sorted({r["participant_name"] for r in all_rows}, key=str.lower)
        for p in participants:
            self._supplier_filter_combo.addItem(p, p)
        self._supplier_filter_combo.blockSignals(False)

        self._filter_status   = "Todos"
        self._filter_supplier = ""
        for m, btn in self._filter_btns.items():
            btn.setChecked(m == "Todos")
        self._supplier_filter_combo.setCurrentIndex(0)

        self._fill_table(all_rows)
        self._update_metrics(all_rows)
        self._update_filter_btns(all_rows)
        self._apply_filter()

        nao = sum(1 for r in all_rows if r["status"] == _STATUS_NAO)
        ja  = sum(1 for r in all_rows if r["status"] == _STATUS_OK)
        sem = sum(1 for r in all_rows if r["status"] == _STATUS_SEM)
        self._status_label.setText(
            f"Verificação concluída: {len(all_rows)} produto(s) — "
            f"{nao} não cadastrado(s) · {ja} já cadastrado(s) · {sem} sem correspondência"
        )
        self._btn_export.setEnabled(bool(all_rows))
        self._btn_atualizar_codigo.setEnabled(bool(all_rows))
        self._btn_cadastrar.setEnabled(nao > 0)

    def _on_v_failed(self, msg: str) -> None:
        self._progress.hide()
        self._progress_label.hide()
        self._btn_verificar.setEnabled(True)
        QMessageBox.critical(self, "Verificar", f"Erro ao verificar:\n{msg}")

    # ── Preenchimento da tabela ───────────────────────────────────────────────

    def _fill_table(self, rows: list[dict]) -> None:
        self._table.setUpdatesEnabled(False)
        self._table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            status = str(row.get("status") or "")
            bg_status, fg_status = _status_colors(status)
            row_bg = _ROW_EVEN if r % 2 == 0 else _ROW_ODD
            for c, field in enumerate(_MAIN_FIELDS):
                val = row.get(field)
                text = str(val if val is not None else "")
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                if c == 0:
                    item.setBackground(bg_status)
                    item.setForeground(fg_status)
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                else:
                    item.setBackground(row_bg)
                self._table.setItem(r, c, item)
        self._table.resizeColumnsToContents()
        self._table.setUpdatesEnabled(True)

    # ── Filtros ───────────────────────────────────────────────────────────────

    def _row_visible(self, row: dict) -> bool:
        if self._filter_status != "Todos" and row.get("status") != self._filter_status:
            return False
        if self._filter_supplier and row.get("participant_name") != self._filter_supplier:
            return False
        return True

    def _apply_filter(self) -> None:
        self._table.setUpdatesEnabled(False)
        visible_nao = 0
        for r, row in enumerate(self._all_rows):
            visible = self._row_visible(row)
            self._table.setRowHidden(r, not visible)
            if visible and row.get("status") == _STATUS_NAO:
                visible_nao += 1
        self._table.setUpdatesEnabled(True)
        self._btn_cadastrar.setEnabled(visible_nao > 0)

    def _set_status_filter(self, mode: str) -> None:
        self._filter_status = mode
        for m, btn in self._filter_btns.items():
            btn.setChecked(m == mode)
        self._apply_filter()

    def _on_supplier_filter_changed(self) -> None:
        self._filter_supplier = str(self._supplier_filter_combo.currentData() or "")
        self._apply_filter()

    def _update_filter_btns(self, rows: list[dict]) -> None:
        total = len(rows)
        nao   = sum(1 for r in rows if r["status"] == _STATUS_NAO)
        ja    = sum(1 for r in rows if r["status"] == _STATUS_OK)
        sem   = sum(1 for r in rows if r["status"] == _STATUS_SEM)
        self._filter_btns["Todos"].setText(f"Todos ({total})")
        self._filter_btns[_STATUS_NAO].setText(f"Não Cadastrados ({nao})")
        self._filter_btns[_STATUS_OK].setText(f"Já Cadastrados ({ja})")
        self._filter_btns[_STATUS_SEM].setText(f"Sem Correspondência ({sem})")

    def _update_metrics(self, rows: list[dict]) -> None:
        fornecedores = len({r["participant_name"] for r in rows})
        nao  = sum(1 for r in rows if r["status"] == _STATUS_NAO)
        ja   = sum(1 for r in rows if r["status"] == _STATUS_OK)
        sem  = sum(1 for r in rows if r["status"] == _STATUS_SEM)
        self._metric_labels["fornecedores"].setText(str(fornecedores))
        self._metric_labels["total"].setText(str(len(rows)))
        self._metric_labels["nao_cad"].setText(str(nao))
        self._metric_labels["ja_cad"].setText(str(ja))
        self._metric_labels["sem_corr"].setText(str(sem))

    # ── Cadastrar ─────────────────────────────────────────────────────────────

    def _cadastrar(self) -> None:
        # Coleta apenas os "Não Cadastrado" visíveis com supplier_id conhecido
        to_register = [
            row for r, row in enumerate(self._all_rows)
            if not self._table.isRowHidden(r)
            and row.get("status") == _STATUS_NAO
            and row.get("supplier_id") is not None
        ]
        if not to_register:
            QMessageBox.information(
                self, "Cadastrar",
                "Nenhum produto não cadastrado visível com fornecedor correspondente."
            )
            return

        fornecedores_envolvidos = len({r["supplier_db_name"] for r in to_register})
        resp = QMessageBox.question(
            self, "Cadastrar Produtos",
            f"{len(to_register)} produto(s) serão cadastrados em {fornecedores_envolvidos} fornecedor(es).\n\n"
            "Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        self._btn_cadastrar.setEnabled(False)
        self._btn_verificar.setEnabled(False)
        self._btn_export.setEnabled(False)
        self._progress.setValue(0)
        self._progress.show()
        self._progress_label.setText("Iniciando...")
        self._progress_label.show()

        self._c_thread = QThread(self)
        self._c_worker = _CadastrarTodosWorker(to_register, self._operation_type, self._repo)
        self._c_worker.moveToThread(self._c_thread)
        self._c_thread.started.connect(self._c_worker.run)
        self._c_worker.progress.connect(self._on_c_progress)
        self._c_worker.finished.connect(self._on_c_done)
        self._c_worker.failed.connect(self._on_c_failed)
        self._c_worker.finished.connect(self._c_thread.quit)
        self._c_worker.failed.connect(self._c_thread.quit)
        self._c_thread.finished.connect(self._c_worker.deleteLater)
        self._c_thread.finished.connect(self._c_thread.deleteLater)
        self._c_thread.finished.connect(lambda: setattr(self, "_c_thread", None))
        self._c_thread.finished.connect(lambda: setattr(self, "_c_worker", None))
        self._c_thread.start()

    def _on_c_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._progress.setValue(int(current / total * 100))
        self._progress_label.setText(msg)

    def _on_c_done(self, stats: dict) -> None:
        self._progress.hide()
        self._progress_label.hide()
        self._btn_verificar.setEnabled(True)
        erros = stats.get("erros", [])
        txt = (
            f"Cadastrados: {stats['cadastrados']} | "
            f"Ignorados: {stats['ignorados']} | "
            f"Erros: {len(erros)}"
        )
        if erros:
            detalhe = "\n".join(f"  [{e['code']}] {e['erro']}" for e in erros[:10])
            txt += f"\n\n{detalhe}"
        QMessageBox.information(self, "Cadastro Concluído", txt)
        # Re-verifica para atualizar os resultados
        self._run_verificar()

    def _on_c_failed(self, msg: str) -> None:
        self._progress.hide()
        self._progress_label.hide()
        self._btn_verificar.setEnabled(True)
        self._btn_cadastrar.setEnabled(True)
        QMessageBox.critical(self, "Erro no Cadastro", f"Erro ao cadastrar:\n{msg}")

    # ── Exportar ──────────────────────────────────────────────────────────────

    def _export(self) -> None:
        if not self._all_rows:
            return

        import re
        op = self._operation_type.lower()
        default_name = f"verificacao_fornecedor_{op}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Excel", default_name, "Arquivo Excel (*.xlsx)"
        )
        if not path:
            return

        try:
            import os
            from app.exporters.workbook_exporter import write_simple_excel_workbook
            from app.exporters.excel_base import (
                EXCEL_STYLE_STATUS_OK, EXCEL_STYLE_STATUS_NAO, EXCEL_STYLE_STATUS_DIV,
                EXCEL_STYLE_ROW_OK, EXCEL_STYLE_ROW_NAO, EXCEL_STYLE_ROW_DIV,
            )

            # Usa "Sem Correspondência" com o estilo de DIV (divergência)
            _STATUS_STYLE = {
                _STATUS_OK:  (EXCEL_STYLE_STATUS_OK,  EXCEL_STYLE_ROW_OK),
                _STATUS_NAO: (EXCEL_STYLE_STATUS_NAO, EXCEL_STYLE_ROW_NAO),
                _STATUS_SEM: (EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV),
            }

            # Exporta apenas os rows visíveis (respeita filtro atual)
            visible_rows = [
                row for r, row in enumerate(self._all_rows)
                if not self._table.isRowHidden(r)
            ]

            all_fields  = _MAIN_FIELDS + _EXPORT_EXTRA_FIELDS
            all_headers = _MAIN_HEADERS + _EXPORT_EXTRA_HEADERS

            data   = [[str(row.get(f) or "") for f in all_fields] for row in visible_rows]
            styles = []
            for row in visible_rows:
                status = str(row.get("status") or "")
                s, r = _STATUS_STYLE.get(status, (EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV))
                styles.append([s] + [r] * (len(all_fields) - 1))

            write_simple_excel_workbook(
                Path(path),
                [(
                    f"Verif. Fornecedor {self._operation_type}",
                    all_headers,
                    data,
                    {"row_style_ids": styles, "include_total": False},
                )],
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportar")
            msg.setText(f"Exportação concluída.\n\nDeseja abrir o arquivo?\n{path}")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Não", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                os.startfile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Exportar", f"Erro ao exportar:\n{exc}")

    # ── Atualizar Código Fornecedor via XML ────────────────────────────────────

    def _abrir_atualizacao_codigo(self) -> None:
        if self._empresa_id is None:
            QMessageBox.warning(self, "Verificação", "Execute a verificação antes.")
            return
        dlg = _AtualizarCodigoFornecedorDialog(
            rows=self._rows,
            empresa_id=self._empresa_id,
            repository=self._repo,
            parent=self,
        )
        dlg.exec()


# ── Worker: escanear XMLs e propor atualizações de codigo_fornecedor ──────────

class _ScanarXMLsWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(list)
    failed   = Signal(str)

    _TOL_QTY = Decimal("0.001")
    _TOL_VAL = Decimal("0.01")

    def __init__(
        self,
        rows: list[dict],
        empresa_id: int,
        repository: MysqlCadastroRepository,
        xml_folder: str,
    ) -> None:
        super().__init__()
        self._rows       = rows
        self._empresa_id = empresa_id
        self._repo       = repository
        self._xml_folder = xml_folder

    def run(self) -> None:
        try:
            from app.parsers.compare_xml import parse_compare_xml_file

            # 1. Carrega catálogo da empresa; só interessa produtos auto-preenchidos
            #    (codigo_fornecedor == codigo_empresa → foram cadastrados pelo sistema)
            self.progress.emit(0, 100, "Carregando catálogo da empresa...")
            catalog = self._repo.get_catalog_products_by_empresa_id(self._empresa_id)
            auto_by_code: dict[str, dict] = {}
            for p in catalog:
                cfe = str(p.get("codigo_empresa") or "").strip()
                cff = str(p.get("codigo_fornecedor") or "").strip()
                if not cfe or cfe != cff:
                    continue
                if cfe not in auto_by_code:
                    auto_by_code[cfe] = p
                ean = "".join(c for c in str(p.get("ean") or "") if c.isdigit())
                if ean and ean not in auto_by_code:
                    auto_by_code[ean] = p

            if not auto_by_code:
                self.finished.emit([])
                return

            # 2. Coleta itens SPED para produtos auto-preenchidos, agrupados por chave NFe
            sped_items_by_key: dict[str, list[dict]] = {}
            for row in self._rows:
                for detail in row.get("launch_details", []):
                    code    = str(detail.get("code") or "").strip()
                    doc_key = str(detail.get("document_key") or "").strip()
                    if not code or len(doc_key) != 44:
                        continue
                    cat_prod = auto_by_code.get(code)
                    if cat_prod is None:
                        digits = "".join(c for c in code if c.isdigit())
                        cat_prod = auto_by_code.get(digits) if digits else None
                    if cat_prod is None:
                        continue
                    qty = detail.get("quantity")
                    val = detail.get("sale_value")
                    if qty is None or val is None:
                        continue
                    sped_items_by_key.setdefault(doc_key, []).append({
                        "code":             code,
                        "product_id":       int(cat_prod["id"]),
                        "fornecedor_id":    int(cat_prod["fornecedor_id"]),
                        "code_old":         str(cat_prod.get("codigo_fornecedor") or code),
                        "description_sped": str(detail.get("description") or "").strip(),
                        "ncm_sped":         str(detail.get("ncm") or cat_prod.get("ncm") or ""),
                        "participant_name":  str(detail.get("participant_name") or "").strip(),
                        "fornecedor_nome":  str(cat_prod.get("fornecedor_nome") or "").strip(),
                        "quantity":         Decimal(str(qty)),
                        "sale_value":       Decimal(str(val)),
                        "document_key":     doc_key,
                    })

            if not sped_items_by_key:
                self.finished.emit([])
                return

            # 3. Escaneia XMLs da pasta, indexa por chave de acesso
            needed_keys = set(sped_items_by_key.keys())
            xml_files   = sorted(Path(self._xml_folder).rglob("*.xml"))
            total_files = len(xml_files)
            xml_by_key: dict[str, object] = {}

            for i, xml_file in enumerate(xml_files):
                if not needed_keys:
                    break
                self.progress.emit(i, total_files, f"Lendo: {xml_file.name[:50]}")
                try:
                    invoice = parse_compare_xml_file(xml_file)
                    if invoice and invoice.key in needed_keys:
                        xml_by_key[invoice.key] = invoice
                        needed_keys.discard(invoice.key)
                except Exception:
                    pass

            # 4. Matching SPED ↔ XML por quantidade + valor dentro da mesma NFe
            proposals: list[dict] = []
            seen: set[tuple] = set()  # evita duplicatas (product_id, code_new)

            for doc_key, sped_items in sped_items_by_key.items():
                invoice = xml_by_key.get(doc_key)
                if invoice is None:
                    continue

                for sitem in sped_items:
                    sped_qty = sitem["quantity"]
                    sped_val = sitem["sale_value"]

                    matches = [
                        xi for xi in invoice.items
                        if abs(Decimal(str(xi.quantity)) - sped_qty) <= self._TOL_QTY
                        and abs(Decimal(str(xi.value)) - sped_val) <= self._TOL_VAL
                    ]
                    if not matches:
                        continue

                    # Pega o primeiro match XML (se houver mais de um item com mesma
                    # qty+valor na NF é inconclusivo no nível do documento, mas a
                    # ambiguidade real é detectada depois em nível de fornecedor)
                    xml_item = matches[0]
                    new_code = str(xml_item.code or "").strip()
                    if not new_code or new_code == sitem["code_old"]:
                        continue

                    key_uniq = (sitem["product_id"], new_code)
                    if key_uniq in seen:
                        continue
                    seen.add(key_uniq)

                    proposals.append({
                        "selected":         True,   # será revisado no pós-processamento
                        "confidence":       "Único",
                        "product_id":       sitem["product_id"],
                        "fornecedor_id":    sitem["fornecedor_id"],
                        "fornecedor_nome":  sitem["fornecedor_nome"],
                        "code_old":         sitem["code_old"],
                        "code_new":         new_code,
                        "description_sped": sitem["description_sped"],
                        "description_xml":  str(xml_item.description or "").strip(),
                        "ean_xml":          str(xml_item.ean or ""),
                        "ncm_sped":         sitem["ncm_sped"],
                        "ncm_xml":          str(xml_item.ncm or ""),
                        "quantity":         str(sped_qty),
                        "value":            str(sped_val),
                        "document_key":     doc_key,
                    })

            # Pós-processamento: ambiguidade real = mesmo code_new para produtos
            # diferentes do MESMO fornecedor (cada fornecedor tem seu espaço de códigos,
            # mas se ele repetiu o código para produtos distintos não dá para saber qual é qual)
            from collections import defaultdict
            forn_code_products: dict[tuple[int, str], set[int]] = defaultdict(set)
            for p in proposals:
                forn_code_products[(p["fornecedor_id"], p["code_new"])].add(p["product_id"])

            for p in proposals:
                if len(forn_code_products[(p["fornecedor_id"], p["code_new"])]) > 1:
                    p["confidence"] = "Ambíguo"
                    p["selected"]   = False

            self.finished.emit(proposals)
        except Exception as exc:
            import traceback
            self.failed.emit(str(exc) + "\n" + traceback.format_exc())


# ── Worker: aplicar as atualizações aprovadas ─────────────────────────────────

class _AplicarCodigosWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(dict)
    failed   = Signal(str)

    def __init__(
        self,
        proposals: list[dict],
        repository: MysqlCadastroRepository,
    ) -> None:
        super().__init__()
        self._proposals = proposals
        self._repo      = repository

    def run(self) -> None:
        try:
            stats  = {"atualizados": 0, "erros": []}
            to_do  = [p for p in self._proposals if p.get("selected")]
            total  = len(to_do)
            for idx, proposal in enumerate(to_do):
                self.progress.emit(idx, total, f"Atualizando {idx + 1}/{total}...")
                try:
                    self._repo.update_produto_codigo_fornecedor(
                        proposal["product_id"],
                        proposal["code_new"],
                    )
                    stats["atualizados"] += 1
                except Exception as exc:
                    stats["erros"].append({"code": proposal["code_old"], "erro": str(exc)})
            self.progress.emit(total, total, "Concluído.")
            self.finished.emit(stats)
        except Exception as exc:
            self.failed.emit(str(exc))


# ── Diálogo: preview e aplicação de atualizações de codigo_fornecedor ─────────

class _AtualizarCodigoFornecedorDialog(QDialog):
    """
    Escaneia uma pasta de XMLs das NFe, faz matching com os itens do SPED
    por quantidade + valor, e propõe atualizações de codigo_fornecedor no
    catálogo. Exibe preview com seleção antes de gravar qualquer coisa.
    """

    _PROPOSAL_HEADERS = [
        "✓", "Confiança", "Fornecedor", "Cód. Atual", "Cód. Novo (XML)",
        "Descrição SPED", "Descrição XML", "EAN XML", "NCM SPED", "NCM XML",
        "Qtd.", "Valor", "Chave NFe",
    ]
    _PROPOSAL_FIELDS = [
        None, "confidence", "fornecedor_nome", "code_old", "code_new",
        "description_sped", "description_xml", "ean_xml", "ncm_sped", "ncm_xml",
        "quantity", "value", "document_key",
    ]

    _BG_UNICO  = QColor("#c6efce")
    _FG_UNICO  = QColor("#276221")
    _BG_AMBIG  = QColor("#ffeb9c")
    _FG_AMBIG  = QColor("#7d4a00")

    def __init__(
        self,
        rows: list[dict],
        empresa_id: int,
        repository: MysqlCadastroRepository,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._rows       = rows
        self._empresa_id = empresa_id
        self._repo       = repository
        self._proposals: list[dict] = []
        self._s_thread: QThread | None = None
        self._s_worker: _ScanarXMLsWorker | None = None
        self._a_thread: QThread | None = None
        self._a_worker: _AplicarCodigosWorker | None = None

        self.setWindowTitle("Atualizar Código Fornecedor via XML das NFe")
        self.setMinimumSize(1200, 640)
        if parent:
            self.setStyleSheet(parent.styleSheet())
        self.resize(1400, 750)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        title = QLabel("Atualizar codigo_fornecedor via XML das NFe")
        title.setObjectName("sectionTitle")
        root.addWidget(title)

        info = QLabel(
            "<b>Regra de segurança:</b> Somente produtos onde "
            "<code>codigo_fornecedor = codigo_empresa</code> (cadastrados automaticamente) "
            "entram no escopo.<br>"
            "Matching: <b>chave NFe + quantidade + valor do item</b>. "
            "Ambíguo = mesmo código para produtos diferentes do mesmo fornecedor "
            "(ficam <b>desmarcados</b> por padrão)."
        )
        info.setWordWrap(True)
        root.addWidget(info)

        # Pasta de XMLs
        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("Pasta de XMLs:"))
        self._folder_edit = QLineEdit()
        self._folder_edit.setPlaceholderText("Selecione a pasta com os arquivos .xml das NFe...")
        self._folder_edit.setReadOnly(True)
        folder_row.addWidget(self._folder_edit, 1)
        btn_browse = QPushButton("Procurar...")
        btn_browse.clicked.connect(self._browse_folder)
        folder_row.addWidget(btn_browse)
        self._btn_scan = QPushButton("Escanear e Propor")
        self._btn_scan.setObjectName("primaryButton")
        self._btn_scan.setEnabled(False)
        self._btn_scan.clicked.connect(self._run_scan)
        folder_row.addWidget(self._btn_scan)
        root.addLayout(folder_row)

        self._status_label = QLabel("")
        self._status_label.setObjectName("muted")
        root.addWidget(self._status_label)

        # Seleção rápida
        sel_row = QHBoxLayout()
        for label, fn in (
            ("Selecionar Todos",     lambda: self._set_all(True)),
            ("Desmarcar Ambíguos",   self._deselect_ambiguous),
            ("Desmarcar Todos",      lambda: self._set_all(False)),
        ):
            btn = QPushButton(label)
            btn.clicked.connect(fn)
            sel_row.addWidget(btn)
        sel_row.addStretch()
        self._count_label = QLabel("")
        sel_row.addWidget(self._count_label)
        root.addLayout(sel_row)

        # Tabela de preview
        self._table = QTableWidget()
        self._table.setColumnCount(len(self._PROPOSAL_HEADERS))
        self._table.setHorizontalHeaderLabels(self._PROPOSAL_HEADERS)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setAlternatingRowColors(False)
        self._table.setItemDelegate(_BrushDelegate(self._table))
        self._table.cellClicked.connect(self._on_cell_clicked)
        root.addWidget(self._table, 1)

        self._progress = QProgressBar()
        self._progress.hide()
        root.addWidget(self._progress)

        bottom = QHBoxLayout()
        self._btn_apply = QPushButton("Aplicar Atualizações Selecionadas")
        self._btn_apply.setObjectName("primaryButton")
        self._btn_apply.setEnabled(False)
        self._btn_apply.clicked.connect(self._apply)
        bottom.addWidget(self._btn_apply)
        bottom.addStretch()
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(self.reject)
        bottom.addWidget(btn_fechar)
        root.addLayout(bottom)

    # ── Pasta ─────────────────────────────────────────────────────────────────

    def _browse_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Selecionar pasta com XMLs das NFe", ""
        )
        if folder:
            self._folder_edit.setText(folder)
            self._btn_scan.setEnabled(True)

    # ── Scan ──────────────────────────────────────────────────────────────────

    def _run_scan(self) -> None:
        folder = self._folder_edit.text().strip()
        if not folder:
            return
        self._btn_scan.setEnabled(False)
        self._btn_apply.setEnabled(False)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.show()
        self._status_label.setText("Escaneando XMLs...")

        self._s_thread = QThread(self)
        self._s_worker = _ScanarXMLsWorker(self._rows, self._empresa_id, self._repo, folder)
        self._s_worker.moveToThread(self._s_thread)
        self._s_thread.started.connect(self._s_worker.run)
        self._s_worker.progress.connect(self._on_scan_progress)
        self._s_worker.finished.connect(self._on_scan_done)
        self._s_worker.failed.connect(self._on_scan_failed)
        self._s_worker.finished.connect(self._s_thread.quit)
        self._s_worker.failed.connect(self._s_thread.quit)
        self._s_thread.finished.connect(self._s_worker.deleteLater)
        self._s_thread.finished.connect(self._s_thread.deleteLater)
        self._s_thread.finished.connect(lambda: setattr(self, "_s_thread", None))
        self._s_thread.finished.connect(lambda: setattr(self, "_s_worker", None))
        self._s_thread.start()

    def _on_scan_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._progress.setValue(int(current / total * 100))
        self._status_label.setText(msg)

    def _on_scan_done(self, proposals: list[dict]) -> None:
        self._progress.hide()
        self._btn_scan.setEnabled(True)
        self._proposals = proposals
        self._fill_table()
        n_uniq = sum(1 for p in proposals if p["confidence"] == "Único")
        n_amb  = sum(1 for p in proposals if p["confidence"] == "Ambíguo")
        if not proposals:
            self._status_label.setText(
                "Nenhuma proposta encontrada. Verifique se os XMLs correspondem às NFes do SPED "
                "e se há produtos com codigo_fornecedor = codigo_empresa."
            )
        else:
            self._status_label.setText(
                f"{len(proposals)} proposta(s) — {n_uniq} únicas (selecionadas) · "
                f"{n_amb} ambíguas (desmarcadas). Revise e clique em Aplicar."
            )
        self._update_count()

    def _on_scan_failed(self, msg: str) -> None:
        self._progress.hide()
        self._btn_scan.setEnabled(True)
        QMessageBox.critical(self, "Erro ao Escanear", f"Erro:\n{msg}")

    # ── Tabela de preview ─────────────────────────────────────────────────────

    def _fill_table(self) -> None:
        self._table.setUpdatesEnabled(False)
        self._table.setRowCount(len(self._proposals))
        for r, p in enumerate(self._proposals):
            is_unico = p["confidence"] == "Único"
            row_bg   = _ROW_EVEN if r % 2 == 0 else _ROW_ODD

            # Coluna 0: checkbox
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            chk.setCheckState(Qt.Checked if p["selected"] else Qt.Unchecked)
            chk.setBackground(row_bg)
            self._table.setItem(r, 0, chk)

            for c, field in enumerate(self._PROPOSAL_FIELDS[1:], start=1):
                val  = str(p.get(field) or "")
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                if c == 1:  # coluna Confiança
                    bg = self._BG_UNICO if is_unico else self._BG_AMBIG
                    fg = self._FG_UNICO if is_unico else self._FG_AMBIG
                    item.setBackground(bg)
                    item.setForeground(fg)
                    font = item.font(); font.setBold(True); item.setFont(font)
                elif c == 4:  # Cód. Novo (XML) — destaca em verde
                    item.setBackground(self._BG_UNICO if is_unico else self._BG_AMBIG)
                else:
                    item.setBackground(row_bg)
                self._table.setItem(r, c, item)

        self._table.resizeColumnsToContents()
        self._table.setUpdatesEnabled(True)

    def _on_cell_clicked(self, row: int, col: int) -> None:
        if col != 0 or row >= len(self._proposals):
            return
        item = self._table.item(row, 0)
        if item is None:
            return
        new_state = Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked
        item.setCheckState(new_state)
        self._proposals[row]["selected"] = (new_state == Qt.Checked)
        self._update_count()

    def _set_all(self, state: bool) -> None:
        for r, p in enumerate(self._proposals):
            p["selected"] = state
            item = self._table.item(r, 0)
            if item:
                item.setCheckState(Qt.Checked if state else Qt.Unchecked)
        self._update_count()

    def _deselect_ambiguous(self) -> None:
        for r, p in enumerate(self._proposals):
            if p["confidence"] == "Ambíguo":
                p["selected"] = False
                item = self._table.item(r, 0)
                if item:
                    item.setCheckState(Qt.Unchecked)
        self._update_count()

    def _update_count(self) -> None:
        selected = sum(1 for p in self._proposals if p.get("selected"))
        self._count_label.setText(f"{selected}/{len(self._proposals)} selecionados")
        self._btn_apply.setEnabled(selected > 0)

    # ── Aplicar ───────────────────────────────────────────────────────────────

    def _apply(self) -> None:
        to_do = [p for p in self._proposals if p.get("selected")]
        if not to_do:
            return
        resp = QMessageBox.question(
            self, "Aplicar Atualizações",
            f"{len(to_do)} produto(s) terão o campo <b>codigo_fornecedor</b> atualizado.\n\n"
            "Essa operação só altera produtos onde codigo_fornecedor = codigo_empresa.\n"
            "Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        self._btn_apply.setEnabled(False)
        self._btn_scan.setEnabled(False)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.show()

        self._a_thread = QThread(self)
        self._a_worker = _AplicarCodigosWorker(to_do, self._repo)
        self._a_worker.moveToThread(self._a_thread)
        self._a_thread.started.connect(self._a_worker.run)
        self._a_worker.progress.connect(self._on_apply_progress)
        self._a_worker.finished.connect(self._on_apply_done)
        self._a_worker.failed.connect(self._on_apply_failed)
        self._a_worker.finished.connect(self._a_thread.quit)
        self._a_worker.failed.connect(self._a_thread.quit)
        self._a_thread.finished.connect(self._a_worker.deleteLater)
        self._a_thread.finished.connect(self._a_thread.deleteLater)
        self._a_thread.finished.connect(lambda: setattr(self, "_a_thread", None))
        self._a_thread.finished.connect(lambda: setattr(self, "_a_worker", None))
        self._a_thread.start()

    def _on_apply_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._progress.setValue(int(current / total * 100))
        self._status_label.setText(msg)

    def _on_apply_done(self, stats: dict) -> None:
        self._progress.hide()
        self._btn_scan.setEnabled(True)
        erros = stats.get("erros", [])
        txt = f"Atualizados: {stats['atualizados']} | Erros: {len(erros)}"
        if erros:
            txt += "\n" + "\n".join(f"  [{e['code']}] {e['erro']}" for e in erros[:10])
        QMessageBox.information(self, "Atualização Concluída", txt)
        applied_ids = {p["product_id"] for p in self._proposals if p.get("selected")}
        self._proposals = [p for p in self._proposals if p["product_id"] not in applied_ids]
        self._fill_table()
        self._update_count()
        self._status_label.setText(
            f"Atualizados: {stats['atualizados']}. "
            f"Propostas restantes: {len(self._proposals)}"
        )

    def _on_apply_failed(self, msg: str) -> None:
        self._progress.hide()
        self._btn_scan.setEnabled(True)
        self._btn_apply.setEnabled(True)
        QMessageBox.critical(self, "Erro ao Aplicar", f"Erro:\n{msg}")
