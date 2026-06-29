from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QBrush, QColor, QPalette
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStyle,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.services.confronto_sped_cadastro import build_confronto_data
from app.services.confronto_rules_builder import (
    LOG_FIELDS, LOG_HEADERS, build_rules_from_confronto,
)

# ── Colunas ───────────────────────────────────────────────────────────────────

def _build_grouped_headers(operation_type: str) -> list[str]:
    ent = operation_type.lower() == "entrada"
    participante = "Fornecedor (SPED)"  if ent else "Cliente (SPED)"
    cst_cad      = "CST ICMS Ent. (Cad.)" if ent else "CST ICMS Sai. (Cad.)"
    cfop_cad     = "CFOP Entrada (Cad.)"  if ent else "CFOP Saida (Cad.)"
    aliq_cad     = "Aliq. ICMS % Ent. (Cad.)" if ent else "Aliq. ICMS % Sai. (Cad.)"
    return [
        "Status", "Periodo", "Cod. Produto", "Descricao SPED", participante,
        "CST ICMS (SPED)", cst_cad,
        "CFOP (SPED)", cfop_cad,
        "Aliq. ICMS % (SPED)", aliq_cad,
        "Aliq. PIS % (SPED)", "Aliq. PIS % (Cad.)",
        "Aliq. COFINS % (SPED)", "Aliq. COFINS % (Cad.)",
        "Descricao Cadastro", "NCM (SPED)", "NCM (Cad.)", "Fornecedor (Cad.)",
        "Total Operacao", "Base ICMS", "Valor ICMS",
    ]

GROUPED_FIELDS = [
    "status", "periodo", "code", "descricao_sped", "fornecedor",
    "cst_sped", "cst_cad",
    "cfop_sped", "cfop_cad",
    "aliq_sped", "aliq_cad",
    "aliq_pis_sped", "aliq_pis_cad",
    "aliq_cofins_sped", "aliq_cofins_cad",
    "descricao_cad", "ncm_sped", "ncm_cad", "fornecedor_cad",
    "total_operacao", "base_icms", "icms_value",
]

def _build_detail_headers(operation_type: str) -> list[str]:
    ent = operation_type.lower() == "entrada"
    participante = "Fornecedor (SPED)"  if ent else "Cliente (SPED)"
    cst_cad      = "CST ICMS Ent. (Cad.)" if ent else "CST ICMS Sai. (Cad.)"
    cfop_cad     = "CFOP Entrada (Cad.)"  if ent else "CFOP Saida (Cad.)"
    aliq_cad     = "Aliq. ICMS % Ent. (Cad.)" if ent else "Aliq. ICMS % Sai. (Cad.)"
    return [
        "Status", "Periodo", "Num. Doc.", "Chave NFe", "Data", participante,
        "Cod. Produto", "Descricao SPED",
        "CST ICMS (SPED)", cst_cad,
        "CFOP (SPED)", cfop_cad,
        "Aliq. ICMS % (SPED)", aliq_cad,
        "Aliq. PIS % (SPED)", "Aliq. PIS % (Cad.)",
        "Aliq. COFINS % (SPED)", "Aliq. COFINS % (Cad.)",
        "Valor Operacao", "Base ICMS", "Valor ICMS",
    ]

DETAIL_FIELDS = [
    "status", "periodo", "document_number", "document_key", "document_date", "participant",
    "code", "description",
    "cst_sped", "cst_cad",
    "cfop_sped", "cfop_cad",
    "aliq_sped", "aliq_cad",
    "aliq_pis_sped", "aliq_pis_cad",
    "aliq_cofins_sped", "aliq_cofins_cad",
    "sale_value", "base_icms", "icms_value",
]

# Cores da célula Status (fundo + texto)
_BG_OK  = QColor("#c6efce")   # verde claro
_BG_DIV = QColor("#ffc7ce")   # vermelho claro
_BG_NAO = QColor("#ffeb9c")   # amarelo/laranja claro
_BG_ZER = QColor("#ffe0b2")   # laranja suave — CFOP zerado no cadastro

_FG_OK  = QColor("#276221")   # verde escuro
_FG_DIV = QColor("#9c0006")   # vermelho escuro
_FG_NAO = QColor("#7d4a00")   # laranja escuro
_FG_ZER = QColor("#5d2e00")   # marrom escuro

# Cores de fundo das linhas (alternadas, aplicadas manualmente)
_ROW_EVEN = QColor("#f3f6f9")
_ROW_ODD  = QColor("#ffffff")


# ── Delegate que garante setBackground() mesmo com stylesheet herdado ─────────

class _BrushDelegate(QStyledItemDelegate):
    """Pinta background/foreground dos itens a partir dos roles, ignorando CSS."""

    def paint(self, painter, option, index) -> None:
        painter.save()
        is_selected = bool(option.state & QStyle.State_Selected)

        # Fundo
        bg = index.data(Qt.BackgroundRole)
        if is_selected:
            painter.fillRect(option.rect, option.palette.color(QPalette.Highlight))
        elif bg is not None:
            color = bg.color() if isinstance(bg, QBrush) else bg
            painter.fillRect(option.rect, color)

        # Texto
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


# ── Worker: cadastrar produtos não cadastrados ────────────────────────────────

class _CadastrarWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        nao_cad: list[dict],
        operation_type: str,
        company_cnpj: str,
        company_name: str,
        fallback_supplier_id: int | None,
        repository: MysqlCadastroRepository,
        environment: str,
    ) -> None:
        super().__init__()
        self._nao_cad = nao_cad
        self._operation_type = operation_type
        self._company_cnpj = company_cnpj
        self._company_name = company_name
        self._fallback_supplier_id = fallback_supplier_id
        self._repo = repository
        self._env = environment

    def run(self) -> None:
        try:
            stats = {"cadastrados": 0, "ignorados": 0, "erros": []}
            total = len(self._nao_cad)

            # Resolve empresa_id
            empresa_id = self._repo.find_company_id(self._env, self._company_name, self._company_cnpj)
            if empresa_id is None:
                self.failed.emit(f"Empresa não encontrada: {self._company_name} / {self._company_cnpj}")
                return

            for idx, grp in enumerate(self._nao_cad):
                self.progress.emit(idx, total, f"Cadastrando {idx + 1}/{total}...")
                code        = str(grp.get("code") or "").strip()
                descricao   = str(grp.get("descricao_sped") or "").strip()
                ncm         = str(grp.get("ncm_sped") or "").strip()
                cst_sped    = str(grp.get("cst_sped") or "").strip().split("|")[0].strip()
                cfop_sped   = str(grp.get("cfop_sped") or "").strip().split("|")[0].strip()
                aliq_sped   = str(grp.get("aliq_sped") or "").strip()
                forn_sped   = str(grp.get("fornecedor") or "").strip()

                if not code:
                    stats["ignorados"] += 1
                    continue

                # Alíquota: SPED tem percentual (18.00), DB armazena fração (0.18)
                try:
                    aliq_frac = str(round(Decimal(aliq_sped.replace(",", ".")) / 100, 4)) if aliq_sped else "0"
                except InvalidOperation:
                    aliq_frac = "0"

                # Tenta encontrar fornecedor pelo nome do SPED
                forn_nomes = [p.strip() for p in forn_sped.split("|") if p.strip()]
                supplier_id: int | None = None
                for forn_nome in forn_nomes:
                    supplier_id = self._repo.find_supplier_id(empresa_id, forn_nome, "")
                    if supplier_id:
                        break
                if not supplier_id and forn_nomes:
                    # Cria o fornecedor automaticamente com o primeiro nome do SPED
                    supplier_id = self._repo.ensure_supplier(empresa_id, forn_nomes[0])
                    stats.setdefault("fornecedores_criados", 0)
                    stats["fornecedores_criados"] += 1
                if not supplier_id:
                    supplier_id = self._fallback_supplier_id
                if not supplier_id:
                    # Sem fornecedor no SPED e sem fallback → usa a própria empresa como fornecedor
                    supplier_id = self._repo.ensure_supplier(empresa_id, self._company_name, self._company_cnpj)
                    stats.setdefault("fornecedores_criados", 0)
                    stats["fornecedores_criados"] += 1

                data: dict = {
                    "codigo_fornecedor": code,
                    "codigo_empresa":    code,
                    "ean":               code,
                    "descricao":         descricao,
                    "ncm":               ncm,
                    "status_produto":    "ATIVO",
                }
                if self._operation_type == "Entrada":
                    data["cst_icms"]     = cst_sped
                    data["cfop_entrada"] = cfop_sped
                    data["aliquota_icms"] = aliq_frac
                else:
                    data["cst_icms_saida"]      = cst_sped
                    data["cfop_saida_empresa"]   = cfop_sped
                    data["aliquota_icms_saida"]  = aliq_frac

                try:
                    self._repo.save_supplier_product(supplier_id, data)
                    stats["cadastrados"] += 1
                except ValueError as exc:
                    if "Ja existe produto com o mesmo Codigo do Fornecedor" in str(exc):
                        # Produto já existe com codigo_fornecedor mas sem codigo_empresa —
                        # atualiza apenas o codigo_empresa para ficar visível no confronto.
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
                            ja_tem = str(existing.get("codigo_empresa") or "") if existing else "?"
                            stats["erros"].append({
                                "code": code,
                                "erro": f"Ja cadastrado (codigo_empresa='{ja_tem}'). Edite manualmente se necessario.",
                            })
                    else:
                        stats["erros"].append({"code": code, "erro": str(exc)})
                except Exception as exc:
                    stats["erros"].append({"code": code, "erro": str(exc)})

            self.progress.emit(total, total, "Concluído.")
            self.finished.emit(stats)
        except Exception as exc:
            self.failed.emit(str(exc))


# ── Dialog: selecionar produtos e fornecedor fallback ────────────────────────

class _CadastrarNaoCadastradosDialog(QDialog):
    def __init__(
        self,
        nao_cad: list[dict],
        operation_type: str,
        company_cnpj: str,
        company_name: str,
        repository: MysqlCadastroRepository,
        environment: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._nao_cad = nao_cad
        self._operation_type = operation_type
        self._company_cnpj = company_cnpj
        self._company_name = company_name
        self._repository = repository
        self._environment = environment
        self._thread: QThread | None = None
        self._worker: _CadastrarWorker | None = None

        self.setWindowTitle("Cadastrar Produtos Não Cadastrados")
        self.setMinimumSize(900, 580)
        if parent:
            self.setStyleSheet(parent.styleSheet())
        self.resize(1050, 640)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        info = QLabel(
            f"<b>{len(nao_cad)}</b> produto(s) sem cadastro encontrado(s) no confronto.<br>"
            "Serão cadastrados com <b>codigo_empresa = codigo_fornecedor = EAN = código do SPED</b>.<br>"
            "O fornecedor é resolvido pelo nome do SPED: se existir, usa; se não existir, <b>cria automaticamente</b>.<br>"
            "O fornecedor de fallback abaixo é usado quando o produto não tem nome de fornecedor no SPED.<br>"
            "Se nenhum fallback for selecionado, a <b>própria empresa é usada como fornecedor</b> (ex.: energia, serviços próprios)."
        )
        info.setWordWrap(True)
        root.addWidget(info)

        # Fornecedor fallback
        forn_box = QGroupBox("Fornecedor Fallback (Quando Não Encontrado pelo Nome do SPED)")
        fb = QHBoxLayout(forn_box)
        fb.addWidget(QLabel("Fornecedor:"))
        self._forn_combo = QComboBox()
        self._forn_combo.setMinimumWidth(300)
        self._forn_combo.addItem("(Nenhum — ignorar se não achar)", None)
        self._load_suppliers()
        fb.addWidget(self._forn_combo, 1)
        root.addWidget(forn_box)

        # Tabela de produtos
        table_box = QGroupBox("Produtos a Cadastrar")
        tb = QVBoxLayout(table_box)
        self._table = QTableWidget()
        self._table.setColumnCount(7)
        self._table.setHorizontalHeaderLabels([
            "Código SPED", "Descrição SPED", "NCM", "CST ICMS", "CFOP", "Alíq. ICMS", "Fornecedor (SPED)",
        ])
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self._table.setAlternatingRowColors(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        for grp in nao_cad:
            r = self._table.rowCount()
            self._table.insertRow(r)
            cfop_sped = str(grp.get("cfop_sped") or "").strip().split("|")[0].strip()
            for c, v in enumerate([
                grp.get("code", ""),
                grp.get("descricao_sped", ""),
                grp.get("ncm_sped", ""),
                str(grp.get("cst_sped", "")).split("|")[0].strip(),
                cfop_sped,
                grp.get("aliq_sped", ""),
                grp.get("fornecedor", ""),
            ]):
                self._table.setItem(r, c, QTableWidgetItem(str(v or "")))
        tb.addWidget(self._table)
        root.addWidget(table_box, 1)

        # Progresso
        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.hide()
        root.addWidget(self._progress)

        self._status_lbl = QLabel("")
        self._status_lbl.setWordWrap(True)
        root.addWidget(self._status_lbl)

        # Botões
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._btn_cadastrar = QPushButton("Cadastrar Todos")
        self._btn_cadastrar.setObjectName("primaryButton")
        self._btn_cadastrar.clicked.connect(self._start)
        btn_row.addWidget(self._btn_cadastrar)
        self._btn_fechar = QPushButton("Fechar")
        self._btn_fechar.clicked.connect(self.reject)
        btn_row.addWidget(self._btn_fechar)
        root.addLayout(btn_row)

    def _load_suppliers(self) -> None:
        try:
            empresa_id = self._repository.find_company_id(
                self._environment, self._company_name, self._company_cnpj
            )
            if empresa_id is None:
                return
            suppliers = self._repository.list_suppliers(empresa_id)
            for s in suppliers:
                label = str(s.get("nome") or "")
                cnpj  = str(s.get("cnpj") or "")
                if cnpj:
                    label += f" ({cnpj})"
                self._forn_combo.addItem(label, int(s["id"]))
        except Exception:
            pass

    def _start(self) -> None:
        self._btn_cadastrar.setEnabled(False)
        self._progress.show()
        self._progress.setValue(0)
        self._status_lbl.setText("Iniciando...")

        fallback_id: int | None = self._forn_combo.currentData()

        self._thread = QThread(self)
        self._worker = _CadastrarWorker(
            self._nao_cad,
            self._operation_type,
            self._company_cnpj,
            self._company_name,
            fallback_id,
            self._repository,
            self._environment,
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_progress(self, current: int, total: int, msg: str) -> None:
        if total > 0:
            self._progress.setValue(int(current / total * 100))
        self._status_lbl.setText(msg)

    def _on_finished(self, stats: dict) -> None:
        if self._thread:
            self._thread.quit()
        self._progress.setValue(100)
        erros = stats.get("erros", [])
        txt = (
            f"Produtos cadastrados: <b>{stats['cadastrados']}</b> | "
            f"Fornecedores criados: <b>{stats.get('fornecedores_criados', 0)}</b> | "
            f"Ignorados: <b>{stats['ignorados']}</b> | "
            f"Erros: <b>{len(erros)}</b>"
        )
        if erros:
            detalhe = "\n".join(f"  [{e['code']}] {e['erro']}" for e in erros[:20])
            txt += f"<br><pre>{detalhe}</pre>"
        self._status_lbl.setText(txt)
        self._btn_fechar.setText("Fechar e Regerar")
        self._btn_fechar.clicked.disconnect()
        self._btn_fechar.clicked.connect(lambda: self.accept())

    def _on_failed(self, msg: str) -> None:
        if self._thread:
            self._thread.quit()
        self._progress.setValue(0)
        self._status_lbl.setText(f"<b style='color:red'>Erro:</b> {msg}")
        self._btn_cadastrar.setEnabled(True)


# ── Worker assíncrono ─────────────────────────────────────────────────────────

class _ConfrontoWorker(QObject):
    finished = Signal(list, list)
    failed = Signal(str)

    def __init__(
        self,
        rows: list[dict],
        catalog: list[dict],
        operation_type: str,
        compare_fields: frozenset,
    ) -> None:
        super().__init__()
        self._rows = rows
        self._catalog = catalog
        self._operation_type = operation_type
        self._compare_fields = compare_fields

    def run(self) -> None:
        try:
            grouped, detail = build_confronto_data(
                self._rows, self._catalog, self._operation_type, self._compare_fields
            )
            self.finished.emit(grouped, detail)
        except Exception as exc:
            self.failed.emit(str(exc))


# ── Dialog ────────────────────────────────────────────────────────────────────

class ConfrontoDialog(QDialog):
    def __init__(
        self,
        rows: list[dict],
        operation_type: str,
        repository: MysqlCadastroRepository,
        environment: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._rows = rows
        self._operation_type = operation_type
        self._repository = repository
        self._environment = environment
        self._grouped_rows: list[dict] = []
        self._detail_rows: list[dict] = []
        self._companies: list[dict] = []
        self._thread: QThread | None = None
        self._worker: _ConfrontoWorker | None = None
        self._filter_mode: str = "Todos"
        self._filter_btns: dict[str, QPushButton] = {}
        self._runtime_rules: list[dict] = []
        self._log_entries: list[dict] = []
        self.accepted_rules: list[dict] = []  # lido pelo pai após exec()
        self._catalog_by_code: dict[str, dict] = {}
        self._catalog_by_ean: dict[str, dict] = {}
        self._catalog_company_name: str = ""
        self._catalog_company_cnpj: str = ""

        self.setWindowTitle(f"Confronto SPED x Cadastro — {operation_type}s")
        self.setMinimumSize(1200, 700)
        if parent:
            self.setStyleSheet(parent.styleSheet())
        self.resize(1400, 820)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Título
        title = QLabel(f"Confronto SPED x Cadastro de Produtos — {operation_type}s")
        title.setObjectName("sectionTitle")
        root.addWidget(title)

        # ── Linha 1: empresa + ações principais ──────────────────────────────
        bar1 = QHBoxLayout()
        bar1.addWidget(QLabel("Empresa:"))
        self._company_combo = QComboBox()
        self._company_combo.setMinimumWidth(220)
        self._load_companies()
        bar1.addWidget(self._company_combo, 1)

        self._btn_gerar = QPushButton("Gerar Confronto")
        self._btn_gerar.setObjectName("primaryButton")
        self._btn_gerar.clicked.connect(self._run)
        bar1.addWidget(self._btn_gerar)

        self._btn_cadastrar_nc = QPushButton("Cadastrar Nao Cadastrados")
        self._btn_cadastrar_nc.setObjectName("primaryButton")
        self._btn_cadastrar_nc.clicked.connect(self._cadastrar_nao_cadastrados)
        bar1.addWidget(self._btn_cadastrar_nc)

        self._btn_rules = QPushButton("Gerar Regras Dinamicas")
        self._btn_rules.clicked.connect(self._gerar_regras)
        bar1.addWidget(self._btn_rules)

        self._btn_reprocess = QPushButton("Reprocessar SPED com Regras")
        self._btn_reprocess.setObjectName("primaryButton")
        self._btn_reprocess.setEnabled(False)
        self._btn_reprocess.clicked.connect(self._reprocessar_sped)
        bar1.addWidget(self._btn_reprocess)

        bar1.addStretch()
        root.addLayout(bar1)

        # ── Linha 2: exportações e relatórios ─────────────────────────────────
        bar2 = QHBoxLayout()

        self._btn_export = QPushButton("Exportar Excel")
        self._btn_export.clicked.connect(self._export)
        bar2.addWidget(self._btn_export)

        self._btn_export_modelo = QPushButton("Exportar Modelo Produto")
        self._btn_export_modelo.clicked.connect(self._export_modelo_produto)
        bar2.addWidget(self._btn_export_modelo)

        self._btn_report = QPushButton("Relatorio Cliente")
        self._btn_report.clicked.connect(self._export_client_report)
        bar2.addWidget(self._btn_report)

        self._btn_pdf = QPushButton("Diagnostico PDF")
        self._btn_pdf.clicked.connect(self._export_diagnostico_pdf)
        bar2.addWidget(self._btn_pdf)

        bar2.addStretch()
        root.addLayout(bar2)

        # Painel de seleção: o que comparar
        from app.services.confronto_sped_cadastro import COMPARE_FIELD_LABELS, DEFAULT_COMPARE_FIELDS
        sel_box = QGroupBox("O que comparar:")
        sel_box.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #d6e0e8; border-radius: 4px; "
            "margin-top: 6px; padding: 4px 8px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }"
        )
        sel_layout = QHBoxLayout(sel_box)
        sel_layout.setContentsMargins(8, 4, 8, 4)
        sel_layout.setSpacing(16)
        self._compare_checks: dict[str, QCheckBox] = {}
        for field_key, field_label in COMPARE_FIELD_LABELS.items():
            chk = QCheckBox(field_label)
            chk.setChecked(field_key in DEFAULT_COMPARE_FIELDS)
            self._compare_checks[field_key] = chk
            sel_layout.addWidget(chk)
        sel_layout.addStretch()
        root.addWidget(sel_box)

        # Métricas
        metrics_row = QHBoxLayout()
        self._metric_labels: dict[str, QLabel] = {}
        for key, label in (
            ("total", "Total Produtos"),
            ("ok", "OK"),
            ("divergencia", "Com Divergencia"),
            ("nao_cad", "Nao Cadastrado"),
            ("detalhes", "Total Detalhes"),
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

        # Filtros rápidos
        filter_bar = QHBoxLayout()
        filter_bar.addWidget(QLabel("Filtrar:"))
        _filter_styles = {
            "Todos":              "QPushButton:checked { background-color: #2563eb; color: white; font-weight: bold; }",
            "Divergentes":        "QPushButton:checked { background-color: #ffc7ce; color: #9c0006; font-weight: bold; }",
            "CFOP zerado no cad.":"QPushButton:checked { background-color: #ffe0b2; color: #5d2e00; font-weight: bold; }",
            "Nao Cadastrados":    "QPushButton:checked { background-color: #ffeb9c; color: #7d4a00; font-weight: bold; }",
            "OK":                 "QPushButton:checked { background-color: #c6efce; color: #276221; font-weight: bold; }",
        }
        for mode, css in _filter_styles.items():
            btn = QPushButton(mode)
            btn.setCheckable(True)
            btn.setChecked(mode == "Todos")
            btn.setStyleSheet(css)
            btn.clicked.connect(lambda _checked, m=mode: self._set_filter(m))
            self._filter_btns[mode] = btn
            filter_bar.addWidget(btn)
        filter_bar.addStretch()
        root.addLayout(filter_bar)

        # Abas
        self._tabs = QTabWidget()
        self._grouped_table = self._make_table(_build_grouped_headers(operation_type))
        self._detail_table  = self._make_table(_build_detail_headers(operation_type))
        self._log_table     = self._make_table(LOG_HEADERS)
        self._tabs.addTab(self._grouped_table, "Agrupado (por Produto)")
        self._tabs.addTab(self._detail_table,  "Detalhado (por Item)")
        self._tabs.addTab(self._log_table,     "Log de Regras")
        self._log_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._log_table.customContextMenuRequested.connect(self._log_context_menu)
        root.addWidget(self._tabs, 1)

    # ── Helpers ───────────────────────────────────────────────────────────────

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
            companies = self._repository.list_companies_for_confronto(self._environment)
            for c in companies:
                label = str(c["nome"] or "").strip() or str(c["cnpj"] or "")
                self._company_combo.addItem(label, str(c["cnpj"] or ""))
            if len(companies) == 1:
                self._company_combo.setCurrentIndex(0)
        except Exception:
            pass

    def _export_filename(self, prefix: str) -> str:
        import re
        cnpj = re.sub(r"\D", "", self._catalog_company_cnpj or "")
        name = re.sub(r"[^\w\s-]", "", self._catalog_company_name or "").strip()
        name = re.sub(r"\s+", "_", name)[:40]
        op   = self._operation_type.lower()
        parts = [prefix, op]
        if cnpj:
            parts.append(cnpj)
        if name:
            parts.append(name)
        return "_".join(parts) + ".xlsx"

    def _status_colors(self, status: str) -> tuple[QColor, QColor]:
        """Retorna (bg_color, fg_color) para o status dado."""
        if status == "OK":
            return _BG_OK, _FG_OK
        if "Cadastrado" in status:
            return _BG_NAO, _FG_NAO
        if status == "Divergencia: CFOP zerado":
            return _BG_ZER, _FG_ZER
        return _BG_DIV, _FG_DIV

    # ── Gerar confronto ───────────────────────────────────────────────────────

    def _run(self) -> None:
        if not self._rows:
            QMessageBox.warning(self, "Confronto", "Nao ha dados de consulta para confrontar. Processe o SPED primeiro.")
            return
        company_cnpj = str(self._company_combo.currentData() or "").strip()
        if not company_cnpj:
            QMessageBox.warning(self, "Confronto", "Selecione uma empresa.")
            return

        self._btn_gerar.setEnabled(False)
        self._status_label.setText("Carregando cadastro de produtos...")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        try:
            catalog = self._repository.get_catalog_products_by_company_cnpj(self._environment, company_cnpj)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            self._btn_gerar.setEnabled(True)
            self._status_label.setText("")
            QMessageBox.critical(self, "Confronto", f"Erro ao carregar cadastro:\n{exc}")
            return

        self._status_label.setText(f"Comparando {len(self._rows)} produto(s) contra {len(catalog)} produto(s) do cadastro...")
        self._catalog_company_name = self._company_combo.currentText()
        self._catalog_company_cnpj = str(self._company_combo.currentData() or "").strip()
        self._catalog_by_code = {}
        self._catalog_by_ean = {}
        for _prod in catalog:
            _code = str(_prod.get("codigo_empresa") or _prod.get("codigo_fornecedor") or "").strip()
            if _code and _code not in self._catalog_by_code:
                self._catalog_by_code[_code] = _prod
            _ean = "".join(c for c in str(_prod.get("ean") or "") if c.isdigit())
            if _ean and _ean not in self._catalog_by_ean:
                self._catalog_by_ean[_ean] = _prod

        compare_fields = frozenset(k for k, chk in self._compare_checks.items() if chk.isChecked())
        if not compare_fields:
            QMessageBox.warning(self, "Confronto", "Selecione ao menos um campo para comparar.")
            self._btn_gerar.setEnabled(True)
            QApplication.restoreOverrideCursor()
            return

        self._thread = QThread(self)
        self._worker = _ConfrontoWorker(self._rows, catalog, self._operation_type, compare_fields)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(lambda: setattr(self, "_thread", None))
        self._thread.finished.connect(lambda: setattr(self, "_worker", None))
        self._thread.start()

    def _on_done(self, grouped: list[dict], detail: list[dict]) -> None:
        QApplication.restoreOverrideCursor()
        self._btn_gerar.setEnabled(True)
        self._grouped_rows = grouped
        self._detail_rows  = detail
        # Preenche tabelas UMA VEZ com todos os dados
        self._fill_table(self._grouped_table, grouped, GROUPED_FIELDS)
        self._fill_table(self._detail_table,  detail,  DETAIL_FIELDS)
        self._update_metrics(grouped, detail)
        self._update_filter_btns(grouped)
        self._filter_mode = "Todos"
        for m, btn in self._filter_btns.items():
            btn.setChecked(m == "Todos")
        # Filtro inicial = Todos → apenas garante que tudo está visível
        self._apply_filter()
        total_ok  = sum(1 for r in grouped if r.get("status") == "OK")
        total_div = sum(1 for r in grouped if "Divergencia" in str(r.get("status") or ""))
        total_zer = sum(1 for r in grouped if "CFOP zerado" in str(r.get("status") or ""))
        total_nao = sum(1 for r in grouped if "Cadastrado" in str(r.get("status") or ""))
        self._status_label.setText(
            f"Confronto gerado: {len(grouped)} produtos — "
            f"{total_ok} OK · {total_div} com divergencia "
            f"({total_zer} CFOP zerado no cad.) · {total_nao} nao cadastrados"
        )

    def _on_failed(self, msg: str) -> None:
        QApplication.restoreOverrideCursor()
        self._btn_gerar.setEnabled(True)
        self._status_label.setText("")
        QMessageBox.critical(self, "Confronto", f"Erro ao confrontar:\n{msg}")

    def _set_filter(self, mode: str) -> None:
        self._filter_mode = mode
        for m, btn in self._filter_btns.items():
            btn.setChecked(m == mode)
        self._apply_filter()

    @staticmethod
    def _row_visible(row: dict, mode: str) -> bool:
        status = str(row.get("status") or "")
        if mode == "OK":
            return status == "OK"
        if mode == "Divergentes":
            return "Divergencia" in status and "CFOP zerado" not in status
        if mode == "CFOP zerado no cad.":
            return "CFOP zerado" in status
        if mode == "Nao Cadastrados":
            return "Cadastrado" in status
        return True  # Todos

    def _apply_filter(self) -> None:
        mode = self._filter_mode
        for t, rows in (
            (self._grouped_table, self._grouped_rows),
            (self._detail_table,  self._detail_rows),
        ):
            t.setUpdatesEnabled(False)
            for r, row in enumerate(rows):
                t.setRowHidden(r, not self._row_visible(row, mode))
            t.setUpdatesEnabled(True)

    def _update_filter_btns(self, grouped: list[dict]) -> None:
        total = len(grouped)
        div   = sum(1 for r in grouped if "Divergencia" in str(r.get("status") or ""))
        zer   = sum(1 for r in grouped if "CFOP zerado" in str(r.get("status") or ""))
        nao   = sum(1 for r in grouped if "Cadastrado"  in str(r.get("status") or ""))
        ok    = sum(1 for r in grouped if r.get("status") == "OK")
        self._filter_btns["Todos"].setText(f"Todos ({total})")
        self._filter_btns["Divergentes"].setText(f"Divergentes ({div})")
        self._filter_btns["CFOP zerado no cad."].setText(f"CFOP zerado no cad. ({zer})")
        self._filter_btns["Nao Cadastrados"].setText(f"Nao Cadastrados ({nao})")
        self._filter_btns["OK"].setText(f"OK ({ok})")

    def _fill_table(self, table: QTableWidget, rows: list[dict], fields: list[str]) -> None:
        table.setUpdatesEnabled(False)
        table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            status = str(row.get("status") or "")
            bg_status, fg_status = self._status_colors(status)
            row_bg = _ROW_EVEN if r % 2 == 0 else _ROW_ODD
            for c, field in enumerate(fields):
                val = str(row.get(field) or "")
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                if c == 0:
                    item.setBackground(bg_status)
                    item.setForeground(fg_status)
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                else:
                    item.setBackground(row_bg)
                table.setItem(r, c, item)
        table.resizeColumnsToContents()
        table.setUpdatesEnabled(True)

    def _update_metrics(self, grouped: list[dict], detail: list[dict]) -> None:
        ok  = sum(1 for r in grouped if r.get("status") == "OK")
        div = sum(1 for r in grouped if "Divergencia" in str(r.get("status") or ""))
        nao = sum(1 for r in grouped if "Cadastrado" in str(r.get("status") or ""))
        for key, val in (
            ("total", len(grouped)),
            ("ok", ok),
            ("divergencia", div),
            ("nao_cad", nao),
            ("detalhes", len(detail)),
        ):
            lbl = self._metric_labels.get(key)
            if lbl:
                lbl.setText(str(val))

    # ── Regras dinâmicas ─────────────────────────────────────────────────────

    def _gerar_regras(self) -> None:
        if not self._grouped_rows:
            QMessageBox.warning(self, "Regras", "Gere o confronto antes de criar regras.")
            return
        compare_fields = frozenset(k for k, chk in self._compare_checks.items() if chk.isChecked())
        self._runtime_rules, self._log_entries = build_rules_from_confronto(
            self._grouped_rows, self._operation_type, compare_fields
        )
        # Vincula índice da regra a cada entrada "REGRA GERADA" para edição via menu
        rule_idx = 0
        for entry in self._log_entries:
            if entry["tipo"] == "REGRA GERADA":
                entry["_rule_idx"] = rule_idx
                rule_idx += 1
        self._fill_log_table(self._log_entries)
        self._tabs.setCurrentWidget(self._log_table)
        n_rules = len(self._runtime_rules)
        n_nao = sum(1 for e in self._log_entries if e["tipo"] == "NAO ENCONTRADO")
        n_sem = sum(1 for e in self._log_entries if e["tipo"] == "SEM CORRECAO")
        self._btn_reprocess.setEnabled(n_rules > 0)
        QMessageBox.information(
            self, "Regras Geradas",
            f"{n_rules} regra(s) gerada(s) para correcao automatica.\n"
            f"{n_nao} produto(s) nao encontrado(s) no cadastro.\n"
            f"{n_sem} produto(s) com divergencia mas sem valor de correcao no cadastro.\n\n"
            + ("Clique em 'Reprocessar SPED com Regras' para aplicar." if n_rules > 0
               else "Nenhuma regra gerada — verifique o cadastro de produtos."),
        )

    def _fill_log_table(self, entries: list[dict]) -> None:
        _TYPE_COLORS = {
            "REGRA GERADA":   (_BG_OK,  _FG_OK),
            "NAO ENCONTRADO": (_BG_NAO, _FG_NAO),
            "SEM CORRECAO":   (_BG_DIV, _FG_DIV),
            "REGRA EXCLUIDA": (QColor("#e0e0e0"), QColor("#888888")),
        }
        self._log_table.setUpdatesEnabled(False)
        self._log_table.setRowCount(len(entries))
        for r, entry in enumerate(entries):
            tipo = str(entry.get("tipo") or "")
            bg, fg = _TYPE_COLORS.get(tipo, (_ROW_EVEN, _BG_OK))
            row_bg = _ROW_EVEN if r % 2 == 0 else _ROW_ODD
            for c, field in enumerate(LOG_FIELDS):
                val = str(entry.get(field) or "")
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                if c == 0:
                    item.setBackground(bg)
                    item.setForeground(fg)
                    font = item.font(); font.setBold(True); item.setFont(font)
                else:
                    item.setBackground(row_bg)
                self._log_table.setItem(r, c, item)
        self._log_table.setUpdatesEnabled(True)

    # ── Menu de contexto do Log de Regras ─────────────────────────────────────

    def _log_context_menu(self, pos) -> None:
        row = self._log_table.rowAt(pos.y())
        if row < 0 or row >= len(self._log_entries):
            return
        entry = self._log_entries[row]
        rule_idx = entry.get("_rule_idx")
        excluida = entry.get("tipo") == "REGRA EXCLUIDA"

        menu = QMenu(self)
        act_ver  = menu.addAction("Ver Regra")
        act_edit = menu.addAction("Editar Regra") if rule_idx is not None and not excluida else None
        if rule_idx is not None and not excluida:
            menu.addSeparator()
            act_del = menu.addAction("Excluir Regra")
        else:
            act_del = None

        action = menu.exec(self._log_table.viewport().mapToGlobal(pos))
        if action == act_ver:
            self._view_or_edit_rule(entry, read_only=True)
        elif act_edit and action == act_edit:
            self._view_or_edit_rule(entry, read_only=False)
        elif act_del and action == act_del:
            self._delete_rule(row, entry)

    def _view_or_edit_rule(self, entry: dict, read_only: bool) -> None:
        rule_idx = entry.get("_rule_idx")
        rule = self._runtime_rules[rule_idx] if rule_idx is not None else None

        dlg = QDialog(self)
        dlg.setWindowTitle("Detalhes da Regra" if read_only else "Editar Regra")
        dlg.setMinimumWidth(440)
        layout = QVBoxLayout(dlg)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        layout.addLayout(form)

        def ro(val: object) -> QLineEdit:
            f = QLineEdit(str(val or ""))
            f.setReadOnly(True)
            f.setStyleSheet("background: #f5f5f5;")
            return f

        form.addRow("Tipo:", ro(entry.get("tipo")))
        form.addRow("Codigo:", ro(entry.get("codigo")))
        form.addRow("Descricao:", ro(entry.get("descricao")))
        form.addRow("CST no SPED:", ro(entry.get("cst_sped")))
        form.addRow("CFOP no SPED:", ro(entry.get("cfop_sped")))

        if rule is not None:
            codes = ", ".join(sorted(rule.get("match_codes", set())))
            form.addRow("Codigo(s):", ro(codes))

            fld_cst = QLineEdit(str(rule.get("new_cst") or ""))
            fld_cst.setReadOnly(read_only)
            form.addRow("Novo CST:", fld_cst)

            fld_cfop = QLineEdit(str(rule.get("new_cfop") or ""))
            fld_cfop.setReadOnly(read_only)
            form.addRow("Novo CFOP:", fld_cfop)

            fld_cnpj = QLineEdit(str(rule.get("document_tax_id") or ""))
            fld_cnpj.setReadOnly(read_only)
            form.addRow("CNPJ Fornecedor:", fld_cnpj)
        else:
            form.addRow("Acao:", ro(entry.get("acao")))

        if read_only:
            btns = QDialogButtonBox(QDialogButtonBox.Close)
            btns.rejected.connect(dlg.reject)
        else:
            btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
            btns.accepted.connect(dlg.accept)
            btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        if dlg.exec() == QDialog.Accepted and not read_only and rule is not None:
            cst  = fld_cst.text().strip()
            cfop = fld_cfop.text().strip()
            cnpj = "".join(c for c in fld_cnpj.text() if c.isdigit())

            if cst:
                rule["new_cst"] = cst
            else:
                rule.pop("new_cst", None)

            if cfop:
                rule["new_cfop"] = cfop
            else:
                rule.pop("new_cfop", None)

            if cnpj and len(cnpj) in {11, 14}:
                rule["document_tax_id"] = cnpj
            else:
                rule.pop("document_tax_id", None)

            # Atualiza exibição do log
            parts = []
            if cst:
                parts.append(f"CST: {entry.get('cst_sped')} → {cst}")
            if cfop:
                parts.append(f"CFOP: {entry.get('cfop_sped')} → {cfop}")
            entry["cst_cad"]  = cst
            entry["cfop_cad"] = cfop
            entry["cnpj_fornecedor"] = fld_cnpj.text().strip()
            entry["acao"] = " | ".join(parts) if parts else "Sem alteracoes definidas"
            self._fill_log_table(self._log_entries)

    def _delete_rule(self, row: int, entry: dict) -> None:
        rule_idx = entry.get("_rule_idx")
        if rule_idx is None:
            return
        resp = QMessageBox.question(
            self, "Excluir Regra",
            f"Excluir regra do produto '{entry.get('codigo')} — {entry.get('descricao')}'?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        self._runtime_rules.pop(rule_idx)

        # Reajusta índices das entradas subsequentes
        for e in self._log_entries:
            idx = e.get("_rule_idx")
            if idx is not None and idx > rule_idx:
                e["_rule_idx"] = idx - 1

        entry.pop("_rule_idx", None)
        entry["tipo"] = "REGRA EXCLUIDA"
        entry["acao"] = "Regra excluida manualmente"

        self._fill_log_table(self._log_entries)
        self._btn_reprocess.setEnabled(len(self._runtime_rules) > 0)

    def _reprocessar_sped(self) -> None:
        if not self._runtime_rules:
            QMessageBox.warning(self, "Reprocessar", "Gere as regras antes de reprocessar.")
            return
        n = len(self._runtime_rules)
        resp = QMessageBox.question(
            self, "Reprocessar SPED",
            f"{n} regra(s) serao aplicadas ao SPED.\n"
            "As regras do Confronto serao combinadas com as regras dinamicas ja configuradas.\n\n"
            "Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return
        self.accepted_rules = list(self._runtime_rules)
        self.accept()

    # ── Relatório Cliente ─────────────────────────────────────────────────────

    def _export_client_report(self) -> None:
        if not self._grouped_rows:
            QMessageBox.warning(self, "Relatorio", "Gere o confronto antes de exportar o relatorio.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Relatorio Cliente",
            self._export_filename("relatorio_confronto"),
            "Arquivo Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            import datetime as dt
            import os
            from app.exporters.workbook_exporter import write_simple_excel_workbook
            from app.exporters.excel_base import (
                EXCEL_STYLE_DEFAULT, EXCEL_STYLE_HEADER, EXCEL_STYLE_HEADER_BLUE,
                EXCEL_STYLE_STATUS_OK, EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_STATUS_NAO,
                EXCEL_STYLE_ROW_OK,   EXCEL_STYLE_ROW_DIV,   EXCEL_STYLE_ROW_NAO,
            )

            grouped = self._grouped_rows
            total    = len(grouped)
            ok_rows  = [r for r in grouped if r.get("status") == "OK"]
            div_rows = [r for r in grouped if "Divergencia" in str(r.get("status") or "")]
            nao_rows = [r for r in grouped if "Cadastrado"  in str(r.get("status") or "")]
            ok, div, nao = len(ok_rows), len(div_rows), len(nao_rows)

            def pct(n: int) -> str:
                return f"{n} ({n / total * 100:.1f}%)" if total > 0 else str(n)

            periods = sorted({str(r.get("periodo") or "") for r in grouped if r.get("periodo")})
            period_range = (
                f"{periods[0]} a {periods[-1]}" if len(periods) > 1
                else (periods[0] if periods else "-")
            )
            company_name = self._company_combo.currentText()
            today = dt.date.today().strftime("%d/%m/%Y")

            cst_only  = sum(1 for r in div_rows if "CST"  in str(r.get("status")) and "CFOP" not in str(r.get("status")))
            cfop_only = sum(1 for r in div_rows if "CFOP" in str(r.get("status")) and "CST"  not in str(r.get("status")))
            both      = sum(1 for r in div_rows if "CST"  in str(r.get("status")) and "CFOP" in str(r.get("status")))

            # ── Aba 1: Resumo ──────────────────────────────────────────────────
            SECAO  = [EXCEL_STYLE_HEADER_BLUE, EXCEL_STYLE_HEADER_BLUE]
            DADO   = [EXCEL_STYLE_DEFAULT,      EXCEL_STYLE_DEFAULT]
            VAZIO  = [EXCEL_STYLE_DEFAULT,      EXCEL_STYLE_DEFAULT]

            sum_rows:   list[list] = []
            sum_styles: list[list] = []

            def _sec(titulo: str) -> None:
                sum_rows.append([titulo, ""]); sum_styles.append(SECAO)

            def _row(campo: str, valor: object) -> None:
                sum_rows.append([campo, str(valor)]); sum_styles.append(DADO)

            def _blank() -> None:
                sum_rows.append(["", ""]); sum_styles.append(VAZIO)

            _sec("IDENTIFICACAO")
            _row("Empresa",            company_name)
            _row("Periodo Analisado",  period_range)
            _row("Tipo de Operacao",   f"{self._operation_type}s")
            _row("Data de Geracao",    today)
            _blank()

            _sec("RESULTADO GERAL")
            _row("Total de Produtos no SPED",         total)
            _row("Produtos OK (sem divergencia)",      pct(ok))
            _row("Com Divergencia de CST e/ou CFOP",  pct(div))
            _row("Nao Cadastrados no Catalogo",        pct(nao))
            _blank()

            _sec("TIPOS DE DIVERGENCIA")
            _row("Somente CST divergente",   cst_only)
            _row("Somente CFOP divergente",  cfop_only)
            _row("CST e CFOP divergentes",   both)

            # ── Aba 2: Divergencias ────────────────────────────────────────────
            DIV_HEADERS = [
                "Tipo de Divergencia", "Codigo", "Descricao",
                "Fornecedor/Emitente",
                "CST no SPED", "CST no Catalogo",
                "CFOP no SPED", "CFOP no Catalogo",
                "Aliq. ICMS SPED", "Aliq. ICMS Catalogo",
                "Descricao no Catalogo",
            ]
            DIV_FIELDS = [
                "_div_tipo", "code", "descricao_sped",
                "fornecedor",
                "cst_sped", "cst_cad",
                "cfop_sped", "cfop_cad",
                "aliq_sped", "aliq_cad",
                "descricao_cad",
            ]

            def _tipo_div(status: str) -> str:
                has_cst  = "CST"  in status
                has_cfop = "CFOP" in status
                if has_cst and has_cfop: return "CST e CFOP"
                if has_cst:              return "CST"
                if has_cfop:             return "CFOP"
                return ""

            div_data:   list[list] = []
            div_styles: list[list] = []
            for r in div_rows:
                enriched = dict(r, _div_tipo=_tipo_div(str(r.get("status") or "")))
                div_data.append([str(enriched.get(f) or "") for f in DIV_FIELDS])
                div_styles.append(
                    [EXCEL_STYLE_STATUS_DIV] + [EXCEL_STYLE_ROW_DIV] * (len(DIV_FIELDS) - 1)
                )

            # ── Aba 3: Nao Cadastrados ─────────────────────────────────────────
            NAO_HEADERS = [
                "Codigo", "Descricao SPED", "Fornecedor/Emitente",
                "CST no SPED", "CFOP no SPED", "Aliq. ICMS SPED",
                "Observacao",
            ]
            NAO_FIELDS = [
                "code", "descricao_sped", "fornecedor",
                "cst_sped", "cfop_sped", "aliq_sped",
                "_obs",
            ]

            nao_data:   list[list] = []
            nao_styles: list[list] = []
            for r in nao_rows:
                enriched = dict(r, _obs="Produto nao encontrado no catalogo da empresa")
                nao_data.append([str(enriched.get(f) or "") for f in NAO_FIELDS])
                nao_styles.append(
                    [EXCEL_STYLE_STATUS_NAO] + [EXCEL_STYLE_ROW_NAO] * (len(NAO_FIELDS) - 1)
                )

            # ── Aba 4: OK (referencia) ─────────────────────────────────────────
            ok_data:   list[list] = []
            ok_styles: list[list] = []
            for r in ok_rows:
                ok_data.append([str(r.get(f) or "") for f in [
                    "code", "descricao_sped", "fornecedor",
                    "cst_sped", "cfop_sped", "aliq_sped",
                ]])
                ok_styles.append(
                    [EXCEL_STYLE_STATUS_OK] + [EXCEL_STYLE_ROW_OK] * 5
                )

            write_simple_excel_workbook(
                Path(path),
                [
                    (
                        "Resumo",
                        ["Campo", "Resultado"],
                        sum_rows,
                        {"row_style_ids": sum_styles, "include_total": False},
                    ),
                    (
                        f"Divergencias ({div})",
                        DIV_HEADERS,
                        div_data,
                        {"row_style_ids": div_styles, "include_total": False},
                    ),
                    (
                        f"Nao Cadastrados ({nao})",
                        NAO_HEADERS,
                        nao_data,
                        {"row_style_ids": nao_styles, "include_total": False},
                    ),
                    (
                        f"OK ({ok})",
                        ["Codigo", "Descricao SPED", "Fornecedor/Emitente",
                         "CST no SPED", "CFOP no SPED", "Aliq. ICMS SPED"],
                        ok_data,
                        {"row_style_ids": ok_styles, "include_total": False},
                    ),
                ],
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Relatorio Cliente")
            msg.setText(f"Relatorio gerado.\n\nDeseja abrir o arquivo?\n{path}")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                os.startfile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Relatorio Cliente", f"Erro ao gerar relatorio:\n{exc}")

    # ── Diagnóstico PDF ───────────────────────────────────────────────────────

    def _export_diagnostico_pdf(self) -> None:
        if not self._grouped_rows:
            QMessageBox.warning(self, "Diagnostico PDF", "Gere o confronto antes de exportar o diagnostico.")
            return
        import re
        import os
        cnpj = re.sub(r"\D", "", self._catalog_company_cnpj or "")
        op   = self._operation_type.lower()
        default_name = f"diagnostico_confronto_{op}_{cnpj}.pdf" if cnpj else f"diagnostico_confronto_{op}.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Diagnostico PDF",
            default_name,
            "PDF (*.pdf)",
        )
        if not path:
            return
        try:
            from app.exporters.confronto_pdf_exporter import generate_confronto_pdf
            generate_confronto_pdf(
                grouped_rows=self._grouped_rows,
                company_name=self._catalog_company_name,
                company_cnpj=self._catalog_company_cnpj,
                operation_type=self._operation_type,
                output_path=path,
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Diagnostico PDF")
            msg.setText(f"PDF gerado com sucesso.\n\nDeseja abrir o arquivo?\n{path}")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                os.startfile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Diagnostico PDF", f"Erro ao gerar PDF:\n{exc}")

    # ── Exportar Modelo Produto ───────────────────────────────────────────────

    def _export_modelo_produto(self) -> None:
        if not self._grouped_rows:
            QMessageBox.warning(self, "Exportar Modelo Produto", "Gere o confronto antes de exportar.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Exportacao de Produtos",
            self._export_filename("confronto_produtos"),
            "Arquivo Excel (*.xlsx)",
        )
        if not path:
            return

        try:
            import os
            from app.exporters.workbook_exporter import write_simple_excel_workbook

            HEADERS = [
                "Status",
                "ID",
                "Empresa",
                "Fornecedor",
                "UF",
                "Classificação",
                "Cod. Forn.",
                "Cod. Empresa",
                "Descrição",
                "EAN",
                "NCM",
                "CEST",
                "Origem (entrada)",
                "CST ICMS (entrada)",
                "% Red BC ICMS",
                "CFOP saída fornecedor",
                "% ICMS (entrada)",
                "CFOP entrada empresa",
                "CST IPI",
                "% IPI",
                "CST PIS (entrada)",
                "CST PIS_COFINS (ENTRADA EMPRESA)",
                "% PIS",
                "CST COFINS (entrada)",
                "% COFINS",
                "Natureza da receita",
                "MVA",
                "Valor ICMS-ST",
                "cClassTrib",
                "cBenef",
                "Origem (saída)",
                "CST ICMS (saída)",
                "CFOP saída empresa",
                "% ICMS (saída)",
                "CST PIS (saída)",
                "CST COFINS (saída)",
                "Natureza da receita",
                "Chave NFe origem",
            ]

            def _pct(v: object) -> str:
                s = str(v or "").strip()
                if not s:
                    return ""
                try:
                    result = round(float(s.replace(",", ".")) * 100, 2)
                    return f"{result:.2f}".replace(".", ",")
                except (ValueError, TypeError):
                    return s

            def _doc_keys(grp: dict) -> list[str]:
                keys = grp.get("document_keys") or []
                if isinstance(keys, (str, bytes)):
                    keys = [keys]
                result: list[str] = []
                for k in keys:
                    t = str(k or "").strip()
                    if t and t not in result:
                        result.append(t)
                return result

            # Calcula o máximo de chaves para definir quantas colunas extras criar
            max_keys = max((_doc_keys(g) for g in self._grouped_rows), key=len, default=[])
            max_keys_count = len(max_keys)
            key_headers = [f"Chave NFe SPED {i}" for i in range(1, max_keys_count + 1)]
            HEADERS.extend(key_headers)

            n_fixed = len(HEADERS) - max_keys_count  # colunas fixas sem as chaves

            rows: list[list[object]] = []
            for grp in self._grouped_rows:
                code = str(grp.get("code") or "").strip()
                cat = self._catalog_by_code.get(code)
                if cat is None:
                    code_digits = "".join(c for c in code if c.isdigit())
                    cat = self._catalog_by_ean.get(code_digits) if code_digits else None

                status = str(grp.get("status") or "")
                sped_keys = _doc_keys(grp)
                key_values = sped_keys + [""] * max(0, max_keys_count - len(sped_keys))

                if cat:
                    row: list[object] = [
                        status,
                        str(cat.get("id") or ""),
                        self._catalog_company_name,
                        str(grp.get("fornecedor") or ""),
                        str(cat.get("fornecedor_uf") or ""),
                        str(cat.get("tipo_produto") or ""),
                        str(cat.get("codigo_fornecedor") or "") or code,
                        str(cat.get("codigo_empresa") or ""),
                        str(cat.get("descricao") or ""),
                        str(cat.get("ean") or ""),
                        str(cat.get("ncm") or ""),
                        str(cat.get("cest") or ""),
                        str(cat.get("origem_entrada") or ""),
                        str(cat.get("cst_icms") or ""),
                        _pct(cat.get("reducao_bc_icms")),
                        str(cat.get("cfop_saida_fornecedor") or ""),
                        _pct(cat.get("aliquota_icms")),
                        str(cat.get("cfop_entrada") or ""),
                        str(cat.get("cst_ipi") or ""),
                        _pct(cat.get("aliquota_ipi")),
                        str(cat.get("cst_pis") or ""),
                        str(cat.get("cst_pis_cofins") or ""),
                        _pct(cat.get("aliquota_pis")),
                        str(cat.get("cst_cofins") or ""),
                        _pct(cat.get("aliquota_cofins")),
                        str(cat.get("natureza_receita_entrada") or ""),
                        _pct(cat.get("mva")),
                        str(cat.get("valor_icms_st") or ""),
                        str(cat.get("c_classtrib") or ""),
                        str(cat.get("c_benef") or ""),
                        str(cat.get("origem_saida") or ""),
                        str(cat.get("cst_icms_saida") or ""),
                        str(cat.get("cfop_saida_empresa") or ""),
                        _pct(cat.get("aliquota_icms_saida")),
                        str(cat.get("cst_pis_saida") or ""),
                        str(cat.get("cst_cofins_saida") or ""),
                        str(cat.get("natureza_receita_saida") or ""),
                        str(cat.get("chave_nfe_origem") or ""),
                    ] + key_values
                else:
                    blanks = [""] * (n_fixed - 18)
                    row = [
                        status,
                        "",
                        self._catalog_company_name,
                        str(grp.get("fornecedor") or ""),
                        "",
                        "",
                        code,
                        "",
                        str(grp.get("descricao_sped") or ""),
                        "",
                        str(grp.get("ncm_sped") or ""),
                        "",
                        "",
                        str(grp.get("cst_sped") or ""),
                        "",
                        "",
                        str(grp.get("aliq_sped") or ""),
                        "",
                    ] + blanks + key_values
                rows.append(row)

            # ── Aba 2: POR_DOCUMENTO (uma linha por NF/documento) ─────────────
            DOC_HEADERS = [
                "Status",
                "Período",
                "Num. Doc.",
                "Chave NFe",
                "Data",
                "Fornecedor/Emitente",
                "Cod. Produto",
                "Descrição SPED",
                "ID",
                "Empresa",
                "Fornecedor (Cad.)",
                "UF",
                "Classificação",
                "Cod. Forn.",
                "Cod. Empresa",
                "Descrição (Cad.)",
                "EAN",
                "NCM",
                "CEST",
                "Origem (entrada)",
                "CST ICMS (entrada)",
                "% Red BC ICMS",
                "CFOP saída fornecedor",
                "% ICMS (entrada)",
                "CFOP entrada empresa",
                "CST IPI",
                "% IPI",
                "CST PIS (entrada)",
                "CST PIS_COFINS (ENTRADA EMPRESA)",
                "% PIS",
                "CST COFINS (entrada)",
                "% COFINS",
                "Natureza da receita",
                "MVA",
                "Valor ICMS-ST",
                "cClassTrib",
                "cBenef",
                "Origem (saída)",
                "CST ICMS (saída)",
                "CFOP saída empresa",
                "% ICMS (saída)",
                "CST PIS (saída)",
                "CST COFINS (saída)",
                "Natureza da receita (saída)",
                "Chave NFe origem",
                "Valor Operação",
                "Base ICMS",
                "Valor ICMS",
            ]

            _n_doc_prefix   = 8   # colunas de cabeçalho do documento
            _n_doc_monetary = 3   # Valor Operação, Base ICMS, Valor ICMS
            _n_doc_catalog  = len(DOC_HEADERS) - _n_doc_prefix - _n_doc_monetary

            doc_rows: list[list[object]] = []
            for det in self._detail_rows:
                d_code = str(det.get("code") or "").strip()
                d_cat = self._catalog_by_code.get(d_code)
                if d_cat is None:
                    d_digits = "".join(c for c in d_code if c.isdigit())
                    d_cat = self._catalog_by_ean.get(d_digits) if d_digits else None

                d_prefix: list[object] = [
                    str(det.get("status") or ""),
                    str(det.get("periodo") or ""),
                    str(det.get("document_number") or ""),
                    str(det.get("document_key") or ""),
                    str(det.get("document_date") or ""),
                    str(det.get("participant") or ""),
                    d_code,
                    str(det.get("description") or ""),
                ]
                d_monetary: list[object] = [
                    str(det.get("sale_value") or ""),
                    str(det.get("base_icms") or ""),
                    str(det.get("icms_value") or ""),
                ]

                if d_cat:
                    d_catalog: list[object] = [
                        str(d_cat.get("id") or ""),
                        self._catalog_company_name,
                        str(d_cat.get("fornecedor_nome") or ""),
                        str(d_cat.get("fornecedor_uf") or ""),
                        str(d_cat.get("tipo_produto") or ""),
                        str(d_cat.get("codigo_fornecedor") or "") or d_code,
                        str(d_cat.get("codigo_empresa") or ""),
                        str(d_cat.get("descricao") or ""),
                        str(d_cat.get("ean") or ""),
                        str(d_cat.get("ncm") or ""),
                        str(d_cat.get("cest") or ""),
                        str(d_cat.get("origem_entrada") or ""),
                        str(d_cat.get("cst_icms") or ""),
                        _pct(d_cat.get("reducao_bc_icms")),
                        str(d_cat.get("cfop_saida_fornecedor") or ""),
                        _pct(d_cat.get("aliquota_icms")),
                        str(d_cat.get("cfop_entrada") or ""),
                        str(d_cat.get("cst_ipi") or ""),
                        _pct(d_cat.get("aliquota_ipi")),
                        str(d_cat.get("cst_pis") or ""),
                        str(d_cat.get("cst_pis_cofins") or ""),
                        _pct(d_cat.get("aliquota_pis")),
                        str(d_cat.get("cst_cofins") or ""),
                        _pct(d_cat.get("aliquota_cofins")),
                        str(d_cat.get("natureza_receita_entrada") or ""),
                        _pct(d_cat.get("mva")),
                        str(d_cat.get("valor_icms_st") or ""),
                        str(d_cat.get("c_classtrib") or ""),
                        str(d_cat.get("c_benef") or ""),
                        str(d_cat.get("origem_saida") or ""),
                        str(d_cat.get("cst_icms_saida") or ""),
                        str(d_cat.get("cfop_saida_empresa") or ""),
                        _pct(d_cat.get("aliquota_icms_saida")),
                        str(d_cat.get("cst_pis_saida") or ""),
                        str(d_cat.get("cst_cofins_saida") or ""),
                        str(d_cat.get("natureza_receita_saida") or ""),
                        str(d_cat.get("chave_nfe_origem") or ""),
                    ]
                else:
                    d_catalog = [""] * _n_doc_catalog

                doc_rows.append(d_prefix + d_catalog + d_monetary)

            write_simple_excel_workbook(
                Path(path),
                [
                    ("BASE_COMPLETA", HEADERS, rows, {"include_total": False}),
                    ("POR_DOCUMENTO", DOC_HEADERS, doc_rows, {"include_total": False}),
                ],
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportar Modelo Produto")
            msg.setText(f"Exportacao concluida.\n\nDeseja abrir o arquivo?\n{path}")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                os.startfile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Exportar Modelo Produto", f"Erro ao exportar:\n{exc}")

    # ── Exportar ─────────────────────────────────────────────────────────────

    def _export(self) -> None:
        if not self._grouped_rows and not self._detail_rows:
            QMessageBox.warning(self, "Exportar", "Gere o confronto antes de exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Excel",
            self._export_filename("confronto"),
            "Arquivo Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            from app.exporters.workbook_exporter import write_simple_excel_workbook
            from app.exporters.excel_base import (
                EXCEL_STYLE_STATUS_OK, EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_STATUS_NAO,
                EXCEL_STYLE_ROW_OK, EXCEL_STYLE_ROW_DIV, EXCEL_STYLE_ROW_NAO,
            )

            from app.exporters.excel_base import EXCEL_STYLE_CURRENCY

            MONEY_FIELDS = {"total_operacao", "sale_value", "base_icms", "icms_value"}

            def _decimal_value(value: object) -> Decimal | str:
                if isinstance(value, Decimal):
                    return value.quantize(Decimal("0.01"))
                text = str(value or "").strip()
                if not text:
                    return ""
                normalized = text.replace("R$", "").replace(" ", "")
                if "," in normalized and "." in normalized:
                    normalized = normalized.replace(".", "").replace(",", ".")
                else:
                    normalized = normalized.replace(",", ".")
                try:
                    return Decimal(normalized).quantize(Decimal("0.01"))
                except InvalidOperation:
                    return text

            def _export_value(row: dict, field: str) -> object:
                value = row.get(field)
                if field in MONEY_FIELDS:
                    return _decimal_value(value)
                return str(value or "")

            def _document_keys_for_row(row: dict) -> list[str]:
                keys = row.get("document_keys") or []
                if isinstance(keys, (str, bytes)):
                    keys = [keys]
                result: list[str] = []
                for key in keys:
                    text = str(key or "").strip()
                    if text and text not in result:
                        result.append(text)
                return result

            def _row_styles(rows: list[dict], fields: list[str]) -> list[list[int]]:
                result = []
                ncols = len(fields)
                for row in rows:
                    status = str(row.get("status") or "")
                    if status == "OK":
                        s, r = EXCEL_STYLE_STATUS_OK, EXCEL_STYLE_ROW_OK
                    elif "Cadastrado" in status:
                        s, r = EXCEL_STYLE_STATUS_NAO, EXCEL_STYLE_ROW_NAO
                    else:
                        s, r = EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV
                    styles = [s] + [r] * (ncols - 1)
                    for index, field in enumerate(fields):
                        if field in MONEY_FIELDS:
                            styles[index] = EXCEL_STYLE_CURRENCY
                    result.append(styles)
                return result

            _LOG_STYLE_MAP = {
                "REGRA GERADA":   (EXCEL_STYLE_STATUS_OK,  EXCEL_STYLE_ROW_OK),
                "NAO ENCONTRADO": (EXCEL_STYLE_STATUS_NAO, EXCEL_STYLE_ROW_NAO),
                "SEM CORRECAO":   (EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV),
            }
            def _log_row_styles(entries: list[dict]) -> list[list[int]]:
                result = []
                for entry in entries:
                    s, r = _LOG_STYLE_MAP.get(str(entry.get("tipo") or ""), (EXCEL_STYLE_STATUS_DIV, EXCEL_STYLE_ROW_DIV))
                    result.append([s] + [r] * (len(LOG_FIELDS) - 1))
                return result

            max_group_keys = max((len(_document_keys_for_row(r)) for r in self._grouped_rows), default=0)
            grouped_key_headers = [f"Chave NFe {index}" for index in range(1, max_group_keys + 1)]
            GROUPED_HEADERS = _build_grouped_headers(self._operation_type)
            DETAIL_HEADERS  = _build_detail_headers(self._operation_type)
            grouped_export_headers = GROUPED_HEADERS + grouped_key_headers
            grouped_export_fields = GROUPED_FIELDS + [f"_document_key_{index}" for index in range(1, max_group_keys + 1)]

            grp_data = []
            for row in self._grouped_rows:
                values = [_export_value(row, field) for field in GROUPED_FIELDS]
                keys = _document_keys_for_row(row)
                values.extend(keys + [""] * max(0, max_group_keys - len(keys)))
                grp_data.append(values)
            det_data = [[_export_value(r, f) for f in DETAIL_FIELDS] for r in self._detail_rows]
            log_data = [[str(e.get(f) or "") for f in LOG_FIELDS]     for e in self._log_entries]
            write_simple_excel_workbook(
                Path(path),
                [
                    (
                        f"Agrupado {self._operation_type}",
                        grouped_export_headers,
                        grp_data,
                        {"row_style_ids": _row_styles(self._grouped_rows, grouped_export_fields), "include_total": False},
                    ),
                    (
                        f"Detalhado {self._operation_type}",
                        DETAIL_HEADERS,
                        det_data,
                        {"row_style_ids": _row_styles(self._detail_rows, DETAIL_FIELDS), "include_total": False},
                    ),
                    (
                        "Log de Regras",
                        LOG_HEADERS,
                        log_data,
                        {"row_style_ids": _log_row_styles(self._log_entries), "include_total": False},
                    ),
                ],
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportar")
            msg.setText(f"Exportacao concluida.\n\nDeseja abrir o arquivo?\n{path}")
            btn_sim = msg.addButton("Sim", QMessageBox.YesRole)
            msg.addButton("Nao", QMessageBox.NoRole)
            msg.exec()
            if msg.clickedButton() == btn_sim:
                import os
                os.startfile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Exportar", f"Erro ao exportar:\n{exc}")

    # ── Cadastrar Nao Cadastrados ─────────────────────────────────────────────

    def _cadastrar_nao_cadastrados(self) -> None:
        if not self._grouped_rows:
            QMessageBox.warning(self, "Cadastrar", "Gere o confronto antes de cadastrar produtos.")
            return

        nao_cad = [r for r in self._grouped_rows if r.get("status") == "Nao Cadastrado"]
        if not nao_cad:
            QMessageBox.information(self, "Cadastrar", "Nenhum produto 'Nao Cadastrado' encontrado no confronto.")
            return

        dlg = _CadastrarNaoCadastradosDialog(
            nao_cad,
            self._operation_type,
            self._catalog_company_cnpj,
            self._catalog_company_name,
            self._repository,
            self._environment,
            self,
        )
        if dlg.exec() == QDialog.Accepted:
            # Recarrega o confronto para refletir os produtos recém-cadastrados
            self._run()
