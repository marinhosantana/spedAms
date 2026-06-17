from __future__ import annotations

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
    QMessageBox,
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

GROUPED_HEADERS = [
    "Status", "Periodo", "Cod. Produto", "Descricao SPED", "Fornecedor/Cliente",
    "CST ICMS (SPED)", "CST ICMS (Cad.)",
    "CFOP (SPED)", "CFOP (Cad.)",
    "Aliq ICMS (SPED)", "Aliq ICMS (Cad.)",
    "Descricao Cadastro", "NCM SPED", "NCM Cad.", "Fornecedor Cadastro",
    "Total Operacao", "Base ICMS", "Valor ICMS",
]
GROUPED_FIELDS = [
    "status", "periodo", "code", "descricao_sped", "fornecedor",
    "cst_sped", "cst_cad",
    "cfop_sped", "cfop_cad",
    "aliq_sped", "aliq_cad",
    "descricao_cad", "ncm_sped", "ncm_cad", "fornecedor_cad",
    "total_operacao", "base_icms", "icms_value",
]

DETAIL_HEADERS = [
    "Status", "Periodo", "Num. Doc.", "Data", "Fornecedor/Cliente",
    "Cod. Produto", "Descricao",
    "CST ICMS (SPED)", "CST ICMS (Cad.)",
    "CFOP (SPED)", "CFOP (Cad.)",
    "Aliq ICMS (SPED)", "Aliq ICMS (Cad.)",
    "Valor Operacao", "Base ICMS", "Valor ICMS",
]
DETAIL_FIELDS = [
    "status", "periodo", "document_number", "document_date", "participant",
    "code", "description",
    "cst_sped", "cst_cad",
    "cfop_sped", "cfop_cad",
    "aliq_sped", "aliq_cad",
    "sale_value", "base_icms", "icms_value",
]

# Cores da célula Status (fundo + texto)
_BG_OK  = QColor("#c6efce")   # verde claro
_BG_DIV = QColor("#ffc7ce")   # vermelho claro
_BG_NAO = QColor("#ffeb9c")   # amarelo/laranja claro

_FG_OK  = QColor("#276221")   # verde escuro
_FG_DIV = QColor("#9c0006")   # vermelho escuro
_FG_NAO = QColor("#7d4a00")   # laranja escuro

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


# ── Worker assíncrono ─────────────────────────────────────────────────────────

class _ConfrontoWorker(QObject):
    finished = Signal(list, list)
    failed = Signal(str)

    def __init__(
        self,
        rows: list[dict],
        catalog: list[dict],
        operation_type: str,
    ) -> None:
        super().__init__()
        self._rows = rows
        self._catalog = catalog
        self._operation_type = operation_type

    def run(self) -> None:
        try:
            grouped, detail = build_confronto_data(self._rows, self._catalog, self._operation_type)
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

        # Barra de empresa + ações
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Empresa:"))
        self._company_combo = QComboBox()
        self._company_combo.setMinimumWidth(300)
        self._load_companies()
        bar.addWidget(self._company_combo, 1)

        self._btn_gerar = QPushButton("Gerar Confronto")
        self._btn_gerar.setObjectName("primaryButton")
        self._btn_gerar.clicked.connect(self._run)
        bar.addWidget(self._btn_gerar)

        self._btn_export = QPushButton("Exportar Excel")
        self._btn_export.clicked.connect(self._export)
        bar.addWidget(self._btn_export)

        self._btn_report = QPushButton("Relatorio Cliente")
        self._btn_report.clicked.connect(self._export_client_report)
        bar.addWidget(self._btn_report)

        self._btn_rules = QPushButton("Gerar Regras Dinamicas")
        self._btn_rules.clicked.connect(self._gerar_regras)
        bar.addWidget(self._btn_rules)

        self._btn_reprocess = QPushButton("Reprocessar SPED com Regras")
        self._btn_reprocess.setObjectName("primaryButton")
        self._btn_reprocess.setEnabled(False)
        self._btn_reprocess.clicked.connect(self._reprocessar_sped)
        bar.addWidget(self._btn_reprocess)

        bar.addStretch()
        root.addLayout(bar)

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
            "Todos":           "QPushButton:checked { background-color: #2563eb; color: white; font-weight: bold; }",
            "Divergentes":     "QPushButton:checked { background-color: #ffc7ce; color: #9c0006; font-weight: bold; }",
            "Nao Cadastrados": "QPushButton:checked { background-color: #ffeb9c; color: #7d4a00; font-weight: bold; }",
            "OK":              "QPushButton:checked { background-color: #c6efce; color: #276221; font-weight: bold; }",
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
        self._grouped_table = self._make_table(GROUPED_HEADERS)
        self._detail_table  = self._make_table(DETAIL_HEADERS)
        self._log_table     = self._make_table(LOG_HEADERS)
        self._tabs.addTab(self._grouped_table, "Agrupado (por Produto)")
        self._tabs.addTab(self._detail_table,  "Detalhado (por Item)")
        self._tabs.addTab(self._log_table,     "Log de Regras")
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
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        t.horizontalHeader().setStretchLastSection(True)
        # Delegate garante que setBackground() não seja sobrescrito pelo stylesheet pai
        t.setItemDelegate(_BrushDelegate(t))
        return t

    def _load_companies(self) -> None:
        try:
            self._companies = self._repository.list_companies(self._environment)
            for c in self._companies:
                self._company_combo.addItem(c["nome"], c["id"])
        except Exception:
            pass

    def _status_colors(self, status: str) -> tuple[QColor, QColor]:
        """Retorna (bg_color, fg_color) para o status dado."""
        if status == "OK":
            return _BG_OK, _FG_OK
        if "Cadastrado" in status:
            return _BG_NAO, _FG_NAO
        return _BG_DIV, _FG_DIV

    # ── Gerar confronto ───────────────────────────────────────────────────────

    def _run(self) -> None:
        if not self._rows:
            QMessageBox.warning(self, "Confronto", "Nao ha dados de consulta para confrontar. Processe o SPED primeiro.")
            return
        company_id = self._company_combo.currentData()
        if company_id is None:
            QMessageBox.warning(self, "Confronto", "Selecione uma empresa.")
            return

        company = next((c for c in self._companies if c["id"] == company_id), None)
        company_cnpj = str(company.get("cnpj") or "") if company else ""

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

        self._thread = QThread(self)
        self._worker = _ConfrontoWorker(self._rows, catalog, self._operation_type)
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
        self._update_metrics(grouped, detail)
        self._update_filter_btns(grouped)
        self._filter_mode = "Todos"
        for m, btn in self._filter_btns.items():
            btn.setChecked(m == "Todos")
        self._apply_filter()
        total_ok  = sum(1 for r in grouped if r.get("status") == "OK")
        total_div = sum(1 for r in grouped if "Divergencia" in str(r.get("status") or ""))
        total_nao = sum(1 for r in grouped if "Cadastrado" in str(r.get("status") or ""))
        self._status_label.setText(
            f"Confronto gerado: {len(grouped)} produtos — "
            f"{total_ok} OK · {total_div} com divergencia · {total_nao} nao cadastrados"
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

    def _apply_filter(self) -> None:
        mode = self._filter_mode
        if mode == "OK":
            grp = [r for r in self._grouped_rows if r.get("status") == "OK"]
            det = [r for r in self._detail_rows  if r.get("status") == "OK"]
        elif mode == "Divergentes":
            grp = [r for r in self._grouped_rows if "Divergencia" in str(r.get("status") or "")]
            det = [r for r in self._detail_rows  if "Divergencia" in str(r.get("status") or "")]
        elif mode == "Nao Cadastrados":
            grp = [r for r in self._grouped_rows if "Cadastrado" in str(r.get("status") or "")]
            det = [r for r in self._detail_rows  if "Cadastrado" in str(r.get("status") or "")]
        else:
            grp = self._grouped_rows
            det = self._detail_rows
        self._fill_table(self._grouped_table, grp, GROUPED_FIELDS)
        self._fill_table(self._detail_table,  det,  DETAIL_FIELDS)

    def _update_filter_btns(self, grouped: list[dict]) -> None:
        total = len(grouped)
        div   = sum(1 for r in grouped if "Divergencia" in str(r.get("status") or ""))
        nao   = sum(1 for r in grouped if "Cadastrado"  in str(r.get("status") or ""))
        ok    = sum(1 for r in grouped if r.get("status") == "OK")
        self._filter_btns["Todos"].setText(f"Todos ({total})")
        self._filter_btns["Divergentes"].setText(f"Divergentes ({div})")
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
                    # Status: fundo forte colorido + texto negrito
                    item.setBackground(bg_status)
                    item.setForeground(fg_status)
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                else:
                    # Demais células: alternância manual (não depende do stylesheet)
                    item.setBackground(row_bg)
                table.setItem(r, c, item)
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
        self._runtime_rules, self._log_entries = build_rules_from_confronto(
            self._grouped_rows, self._operation_type
        )
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
            "REGRA GERADA":  (_BG_OK,  _FG_OK),
            "NAO ENCONTRADO": (_BG_NAO, _FG_NAO),
            "SEM CORRECAO":  (_BG_DIV, _FG_DIV),
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
            f"relatorio_confronto_{self._operation_type.lower()}.xlsx",
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

    # ── Exportar ─────────────────────────────────────────────────────────────

    def _export(self) -> None:
        if not self._grouped_rows and not self._detail_rows:
            QMessageBox.warning(self, "Exportar", "Gere o confronto antes de exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Excel",
            f"confronto_{self._operation_type.lower()}.xlsx",
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
                    result.append([s] + [r] * (ncols - 1))
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

            grp_data = [[str(r.get(f) or "") for f in GROUPED_FIELDS] for r in self._grouped_rows]
            det_data = [[str(r.get(f) or "") for f in DETAIL_FIELDS]  for r in self._detail_rows]
            log_data = [[str(e.get(f) or "") for f in LOG_FIELDS]     for e in self._log_entries]
            write_simple_excel_workbook(
                Path(path),
                [
                    (
                        f"Agrupado {self._operation_type}",
                        GROUPED_HEADERS,
                        grp_data,
                        {"row_style_ids": _row_styles(self._grouped_rows, GROUPED_FIELDS), "include_total": False},
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
