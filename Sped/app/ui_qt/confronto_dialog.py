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

        # Abas
        self._tabs = QTabWidget()
        self._grouped_table = self._make_table(GROUPED_HEADERS)
        self._detail_table  = self._make_table(DETAIL_HEADERS)
        self._tabs.addTab(self._grouped_table, "Agrupado (por Produto)")
        self._tabs.addTab(self._detail_table,  "Detalhado (por Item)")
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
        self._fill_table(self._grouped_table, grouped, GROUPED_FIELDS)
        self._fill_table(self._detail_table,  detail,  DETAIL_FIELDS)
        self._update_metrics(grouped, detail)
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
                EXCEL_STYLE_HEADER,
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

            grp_data = [[str(r.get(f) or "") for f in GROUPED_FIELDS] for r in self._grouped_rows]
            det_data = [[str(r.get(f) or "") for f in DETAIL_FIELDS]  for r in self._detail_rows]
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
