from __future__ import annotations

import csv
import datetime as dt
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.config import MYSQL_DEFAULT_CONFIG
from app.exporters.workbook_exporter import write_simple_csv_file, write_simple_excel_workbook
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.services.app_paths import get_application_base_dir, get_application_environment, get_environment_config_path, get_project_root_dir
from app.services.analysis_reports import (
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_product_monthly_linear_dataset,
)
from app.services.operation_summary import build_filtered_apuracao_rows, build_reduction_launch_rows
from app.services.path_selection import append_unique_paths, collapse_xml_selection_paths, format_selected_paths, limit_selected_paths, parse_selected_paths
from app.services.period_comparisons import build_entry_period_comparison_rows
from app.services.tax_rules import compute_display_icms_rate


COLORS = {
    "bg": "#f3f6f9",
    "sidebar": "#17212b",
    "sidebar_active": "#25384a",
    "sidebar_text": "#eaf1f7",
    "panel": "#ffffff",
    "line": "#d6e0e8",
    "head": "#edf3f7",
    "text": "#1d2730",
    "muted": "#5d6e7d",
    "accent": "#1f6fd1",
    "ok": "#1f8f5f",
    "warn": "#b56b12",
    "bad": "#b42318",
}


class QtSpedApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.environment = get_application_environment()
        self.base_dir = get_application_base_dir(__file__)
        self.project_root_dir = get_project_root_dir(self.base_dir)
        self.mysql_config_path = get_environment_config_path(self.project_root_dir / "app" / "ui", "mysql_config", self.environment)
        self.mysql_repo = MysqlCadastroRepository(
            self.mysql_config_path,
            self.project_root_dir / "mysql_schema.sql",
            self.get_mysql_default_config(),
        )
        self.entry_rows: list[dict[str, object]] = []
        self.filtered_entry_rows: list[dict[str, object]] = []
        self.period_checks: dict[str, QCheckBox] = {}

        self.setWindowTitle(f"Revisor de SPED Qt [{self.environment.upper()}]")
        self.resize(1360, 820)
        self.apply_styles()
        self.build_shell()
        self.show_page(0, "Dashboard")

    def get_mysql_default_config(self) -> dict[str, str]:
        config = dict(MYSQL_DEFAULT_CONFIG)
        if self.environment == "dev":
            config["database"] = "sped_icms_dev"
        return config

    def apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QMainWindow, QWidget#contentHost, QDialog#popupDialog {{
                background: {COLORS["bg"]};
                color: {COLORS["text"]};
                font-family: Segoe UI;
                font-size: 14px;
            }}
            QWidget#sidebar {{
                background: {COLORS["sidebar"]};
            }}
            QLabel#brand {{
                color: {COLORS["sidebar_text"]};
                font-size: 21px;
                font-weight: 700;
            }}
            QPushButton#navButton {{
                background: transparent;
                color: #d9e5ef;
                border: 0;
                border-radius: 6px;
                padding: 11px 12px;
                text-align: left;
                font-weight: 600;
            }}
            QPushButton#navButton:hover, QPushButton#navButton:checked {{
                background: {COLORS["sidebar_active"]};
                color: #ffffff;
            }}
            QLabel#pageTitle {{
                color: {COLORS["text"]};
                font-size: 24px;
                font-weight: 750;
            }}
            QLabel#muted {{
                color: {COLORS["muted"]};
                font-size: 13px;
            }}
            QFrame#panel {{
                background: {COLORS["panel"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 8px;
            }}
            QLabel#sectionTitle {{
                color: {COLORS["text"]};
                font-size: 15px;
                font-weight: 750;
            }}
            QLabel#popupTitle {{
                color: {COLORS["text"]};
                font-size: 18px;
                font-weight: 800;
            }}
            QLabel#popupFooter {{
                color: {COLORS["text"]};
                font-size: 14px;
                font-weight: 750;
            }}
            QLabel#metricKey {{
                color: {COLORS["muted"]};
                font-size: 12px;
                font-weight: 700;
            }}
            QLabel#metricValue {{
                color: {COLORS["text"]};
                font-size: 18px;
                font-weight: 800;
            }}
            QPushButton {{
                background: #ffffff;
                border: 1px solid #bdcbd7;
                border-radius: 6px;
                color: {COLORS["text"]};
                padding: 8px 12px;
                font-weight: 650;
            }}
            QPushButton:hover {{
                background: #f8fbfd;
            }}
            QPushButton#primaryButton {{
                background: {COLORS["accent"]};
                border-color: {COLORS["accent"]};
                color: #ffffff;
            }}
            QLineEdit, QComboBox {{
                background: #ffffff;
                border: 1px solid #c8d4df;
                border-radius: 6px;
                color: {COLORS["text"]};
                padding: 7px 9px;
                min-height: 22px;
            }}
            QComboBox QAbstractItemView {{
                background: #ffffff;
                color: {COLORS["text"]};
                selection-background-color: #dbeafe;
                selection-color: {COLORS["text"]};
            }}
            QCheckBox {{
                color: {COLORS["text"]};
                spacing: 7px;
            }}
            QTableWidget {{
                background: #ffffff;
                alternate-background-color: #fbfdff;
                border: 1px solid {COLORS["line"]};
                border-radius: 8px;
                color: {COLORS["text"]};
                gridline-color: #e7edf2;
                selection-background-color: #dbeafe;
                selection-color: {COLORS["text"]};
            }}
            QTableWidget::item {{
                color: {COLORS["text"]};
                padding: 4px;
            }}
            QTableWidget::item:selected {{
                background: #dbeafe;
                color: {COLORS["text"]};
            }}
            QScrollBar:vertical {{
                background: #edf3f7;
                width: 16px;
                margin: 0;
                border-left: 1px solid {COLORS["line"]};
            }}
            QScrollBar::handle:vertical {{
                background: #8fa5b8;
                min-height: 32px;
                border-radius: 7px;
                margin: 2px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #6f879d;
            }}
            QScrollBar:horizontal {{
                background: #edf3f7;
                height: 16px;
                margin: 0;
                border-top: 1px solid {COLORS["line"]};
            }}
            QScrollBar::handle:horizontal {{
                background: #8fa5b8;
                min-width: 32px;
                border-radius: 7px;
                margin: 2px;
            }}
            QScrollBar::handle:horizontal:hover {{
                background: #6f879d;
            }}
            QScrollBar::add-line, QScrollBar::sub-line {{
                width: 0;
                height: 0;
            }}
            QScrollBar::add-page, QScrollBar::sub-page {{
                background: transparent;
            }}
            QHeaderView::section {{
                background: {COLORS["head"]};
                color: #273542;
                border: 0;
                border-right: 1px solid {COLORS["line"]};
                border-bottom: 1px solid {COLORS["line"]};
                padding: 8px 9px;
                font-size: 12px;
                font-weight: 750;
            }}
            QStatusBar {{
                background: {COLORS["bg"]};
                color: {COLORS["muted"]};
            }}
            """
        )

    def build_shell(self) -> None:
        root = QWidget()
        root.setObjectName("contentHost")
        shell = QHBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(230)
        nav_layout = QVBoxLayout(sidebar)
        nav_layout.setContentsMargins(10, 18, 10, 18)
        nav_layout.setSpacing(4)
        brand = QLabel("SPED Next")
        brand.setObjectName("brand")
        nav_layout.addWidget(brand)
        nav_layout.addSpacing(12)

        self.nav_buttons: list[QPushButton] = []
        for title, index in (("Dashboard", 0), ("Consulta Entradas", 1), ("SPEDs Arquivados", 2), ("Configuracoes", 3)):
            button = QPushButton(title)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.clicked.connect(lambda _checked=False, current=index, label=title: self.show_page(current, label))
            nav_layout.addWidget(button)
            self.nav_buttons.append(button)
        nav_layout.addStretch()
        shell.addWidget(sidebar)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 10)
        content_layout.setSpacing(12)

        header = QHBoxLayout()
        self.page_title = QLabel("")
        self.page_title.setObjectName("pageTitle")
        env_label = QLabel(f"Ambiente: {self.environment.upper()}")
        env_label.setObjectName("muted")
        header.addWidget(self.page_title)
        header.addStretch()
        header.addWidget(env_label)
        content_layout.addLayout(header)

        self.stack = QStackedWidget()
        self.stack.addWidget(self.build_dashboard_page())
        self.stack.addWidget(self.build_entry_page())
        self.stack.addWidget(self.build_placeholder_page("SPEDs Arquivados", "Esta tela sera migrada depois da consulta principal."))
        self.stack.addWidget(self.build_placeholder_page("Configuracoes", f"Config MySQL: {self.mysql_config_path}"))
        content_layout.addWidget(self.stack, 1)
        shell.addWidget(content, 1)
        self.setCentralWidget(root)
        self.statusBar().showMessage("Pronto.")

    def show_page(self, index: int, title: str) -> None:
        self.stack.setCurrentIndex(index)
        self.page_title.setText(title)
        for button_index, button in enumerate(self.nav_buttons):
            button.setChecked(button_index == index)

    def build_dashboard_page(self) -> QWidget:
        page = QWidget()
        layout = QGridLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        for column, (title, value) in enumerate((("Perfis Arquivados", "-"), ("Arquivos SPED", "-"), ("Banco de Dados", "Pronto"))):
            layout.addWidget(self.create_metric_card(title, value), 0, column)
        layout.setColumnStretch(0, 1)
        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(2, 1)
        layout.setRowStretch(1, 1)
        return page

    def build_entry_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Consulta Entradas ICMS")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_entry_page))
        toolbar_layout.addWidget(self.create_button("Exportar", self.export_entry_filter))
        toolbar_layout.addWidget(self.create_button("Processar", self.process_entry_query, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.sped_input = QLineEdit()
        self.xml_input = QLineEdit()
        self.add_path_row(setup_layout, 0, "SPEDs", self.sped_input, self.select_sped_files, lambda: self.sped_input.clear())
        self.add_path_row(setup_layout, 1, "XMLs 55/65", self.xml_input, self.select_xml_sources, lambda: self.xml_input.clear())
        layout.addWidget(setup)

        filters = self.create_panel()
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(12, 12, 12, 12)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(8)
        self.periods_widget = QWidget()
        self.periods_layout = QHBoxLayout(self.periods_widget)
        self.periods_layout.setContentsMargins(0, 0, 0, 0)
        self.periods_layout.setSpacing(8)
        self.cst_input = QLineEdit()
        self.cst_input.setPlaceholderText("000, 020, 060")
        self.cfop_input = QLineEdit()
        self.cfop_input.setPlaceholderText("1101, 1401")
        self.status_combo = QComboBox()
        self.status_combo.addItems(["Todos", "Ok", "Sem credito", "Sem entrada", "Com divergencia"])
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Codigo, descricao ou fornecedor")
        self.add_filter_box(filters_layout, 0, 0, "Periodos", self.periods_widget)
        self.add_filter_box(filters_layout, 0, 1, "CST", self.cst_input)
        self.add_filter_box(filters_layout, 0, 2, "CFOP", self.cfop_input)
        search_box = QWidget()
        search_layout = QVBoxLayout(search_box)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)
        search_layout.addWidget(self.status_combo)
        search_layout.addWidget(self.search_input)
        self.add_filter_box(filters_layout, 0, 3, "Status / Busca", search_box)
        filters_layout.setColumnStretch(0, 2)
        filters_layout.setColumnStretch(1, 1)
        filters_layout.setColumnStretch(2, 1)
        filters_layout.setColumnStretch(3, 2)
        layout.addWidget(filters)

        filter_actions = QHBoxLayout()
        filter_actions.addWidget(self.create_button("Aplicar filtros", self.refresh_entry_table, primary=True))
        filter_actions.addWidget(self.create_button("Limpar filtros", self.clear_entry_filters))
        filter_actions.addWidget(self.create_button("Exportar filtro atual", self.export_entry_filter))
        filter_actions.addWidget(self.create_button("Entradas", self.open_entry_operation_summary_popup))
        filter_actions.addWidget(self.create_button("Comp. Diag. Credito", self.open_entry_credit_comparison_popup))
        filter_actions.addWidget(self.create_button("Diag. Credito", self.open_entry_credit_diagnostic_popup))
        filter_actions.addWidget(self.create_button("Curva ABC", self.open_entry_abc_popup))
        filter_actions.addWidget(self.create_button("Reducao BC", self.open_entry_reduction_popup))
        filter_actions.addWidget(self.create_button("Apuracao", self.open_entry_apuracao_popup))
        filter_actions.addWidget(self.create_button("Espelho Docs", self.open_entry_docs_popup))
        filter_actions.addStretch()
        layout.addLayout(filter_actions)

        hint = QLabel("Dica: clique duas vezes em uma linha para abrir o Espelho Docs.")
        hint.setObjectName("muted")
        layout.addWidget(hint)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(8)
        self.metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate(
            (
                ("rows", "Linhas"),
                ("products", "Produtos"),
                ("operation", "Operacao"),
                ("base", "Base ICMS"),
                ("icms", "ICMS"),
            )
        ):
            card, value_label = self.create_metric_card_with_label(label, "0")
            self.metric_labels[key] = value_label
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        self.entry_table = QTableWidget()
        self.entry_table.setColumnCount(13)
        self.entry_table.setHorizontalHeaderLabels(
            ["Periodo", "Curva", "Codigo", "Descricao", "Fornecedor", "NCM", "CEST", "CFOP", "CST", "Aliq.", "Aliq. Efetiva", "Docs", "Lanc."]
        )
        self.entry_table.setAlternatingRowColors(True)
        self.entry_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.entry_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.entry_table.verticalHeader().setVisible(False)
        self.entry_table.horizontalHeader().setStretchLastSection(False)
        self.entry_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.entry_table.setColumnWidth(0, 88)
        self.entry_table.setColumnWidth(1, 72)
        self.entry_table.setColumnWidth(2, 105)
        self.entry_table.setColumnWidth(3, 330)
        self.entry_table.setColumnWidth(4, 260)
        self.entry_table.setColumnWidth(5, 90)
        self.entry_table.setColumnWidth(6, 90)
        self.entry_table.setColumnWidth(7, 85)
        self.entry_table.setColumnWidth(8, 80)
        self.entry_table.setColumnWidth(9, 85)
        self.entry_table.setColumnWidth(10, 105)
        self.entry_table.setColumnWidth(11, 65)
        self.entry_table.setColumnWidth(12, 75)
        self.entry_table.itemDoubleClicked.connect(lambda _item: self.open_entry_docs_popup())
        layout.addWidget(self.entry_table, 1)
        self.bind_entry_live_filters()
        return page

    def bind_entry_live_filters(self) -> None:
        self.cst_input.textChanged.connect(self.refresh_entry_table)
        self.cfop_input.textChanged.connect(self.refresh_entry_table)
        self.search_input.textChanged.connect(self.refresh_entry_table)
        self.status_combo.currentTextChanged.connect(self.refresh_entry_table)

    def create_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("panel")
        return panel

    def create_button(self, text: str, callback: Callable[[], None], primary: bool = False) -> QPushButton:
        button = QPushButton(text)
        if primary:
            button.setObjectName("primaryButton")
        button.clicked.connect(callback)
        return button

    def create_metric_card(self, label: str, value: str) -> QFrame:
        card, _value_label = self.create_metric_card_with_label(label, value)
        return card

    def create_metric_card_with_label(self, label: str, value: str) -> tuple[QFrame, QLabel]:
        card = self.create_panel()
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        key = QLabel(label)
        key.setObjectName("metricKey")
        value_label = QLabel(value)
        value_label.setObjectName("metricValue")
        layout.addWidget(key)
        layout.addWidget(value_label)
        return card, value_label

    def add_path_row(self, layout: QGridLayout, row: int, label: str, field: QLineEdit, select_callback: Callable[[], None], clear_callback: Callable[[], None]) -> None:
        label_widget = QLabel(label)
        label_widget.setObjectName("muted")
        layout.addWidget(label_widget, row, 0)
        layout.addWidget(field, row, 1)
        layout.addWidget(self.create_button("Selecionar", select_callback), row, 2)
        layout.addWidget(self.create_button("Limpar", clear_callback), row, 3)
        layout.setColumnStretch(1, 1)

    def add_filter_box(self, layout: QGridLayout, row: int, column: int, label: str, widget: QWidget) -> None:
        box = QWidget()
        box_layout = QVBoxLayout(box)
        box_layout.setContentsMargins(0, 0, 0, 0)
        box_layout.setSpacing(5)
        label_widget = QLabel(label)
        label_widget.setObjectName("metricKey")
        box_layout.addWidget(label_widget)
        box_layout.addWidget(widget)
        layout.addWidget(box, row, column)

    def build_placeholder_page(self, title: str, description: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        panel = self.create_panel()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        description_label = QLabel(description)
        description_label.setObjectName("muted")
        panel_layout.addWidget(title_label)
        panel_layout.addWidget(description_label)
        layout.addWidget(panel)
        layout.addStretch()
        return page

    def select_sped_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Fiscais", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        paths = append_unique_paths(parse_selected_paths(self.sped_input.text()), files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        self.sped_input.setText(format_selected_paths(paths))
        if limit_exceeded:
            QMessageBox.warning(self, "SPED Qt", "A consulta aceita no maximo 12 SPEDs.")

    def select_xml_sources(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar XMLs 55/65", "", "Arquivos XML (*.xml);;Todos os arquivos (*.*)")
        if files:
            current_paths = parse_selected_paths(self.xml_input.text())
            collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(files)
            if collapsed_to_folder:
                current_paths = [path for path in current_paths if not path.is_file()]
            self.xml_input.setText(format_selected_paths(append_unique_paths(current_paths, collapsed_paths)))
            return
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta de XMLs 55/65")
        if directory:
            self.xml_input.setText(format_selected_paths(append_unique_paths(parse_selected_paths(self.xml_input.text()), [Path(directory)])))

    def process_entry_query(self) -> None:
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            sped_paths = parse_selected_paths(self.sped_input.text())
            xml_sources = parse_selected_paths(self.xml_input.text())
            period_labels, rows = build_entry_period_comparison_rows(sped_paths, xml_sources)
            self.entry_rows = list(rows)
            self.rebuild_period_checks(list(period_labels))
            self.refresh_entry_table()
            self.statusBar().showMessage(f"Entradas processadas: {len(rows)} linha(s), {len(period_labels)} periodo(s). {dt.datetime.now().strftime('%H:%M:%S')}")
        except Exception as exc:
            QMessageBox.critical(self, "SPED Qt", str(exc))
            self.statusBar().showMessage(f"Falha no processamento: {exc}")
        finally:
            QApplication.restoreOverrideCursor()

    def rebuild_period_checks(self, labels: list[str]) -> None:
        while self.periods_layout.count():
            item = self.periods_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.period_checks = {}
        if not labels:
            empty = QLabel("Nenhum periodo")
            empty.setObjectName("muted")
            self.periods_layout.addWidget(empty)
            return
        for label in labels:
            check = QCheckBox(label)
            check.setChecked(True)
            check.stateChanged.connect(self.refresh_entry_table)
            self.period_checks[label] = check
            self.periods_layout.addWidget(check)
        self.periods_layout.addStretch()

    def parse_filter_tokens(self, value: str) -> set[str]:
        return {part.strip() for part in value.replace(";", ",").replace("|", ",").split(",") if part.strip()}

    def normalize_search_text(self, value: object) -> str:
        text = str(value or "").lower()
        replacements = {"á": "a", "à": "a", "ã": "a", "â": "a", "é": "e", "ê": "e", "í": "i", "ó": "o", "ô": "o", "õ": "o", "ú": "u", "ç": "c"}
        for source, target in replacements.items():
            text = text.replace(source, target)
        return " ".join(text.split())

    def first_row_value(self, row: dict[str, object], *keys: str, default: object = "") -> object:
        for key in keys:
            value = row.get(key)
            if value not in (None, ""):
                return value
        return default

    def decimal_value(self, value: object) -> Decimal:
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(str(value or "0").replace(",", "."))
        except (InvalidOperation, ValueError):
            return Decimal("0")

    def selected_periods(self) -> set[str]:
        selected = {label for label, check in self.period_checks.items() if check.isChecked()}
        return selected or set(self.period_checks)

    def filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.selected_periods()
        cst_filters = self.parse_filter_tokens(self.cst_input.text())
        cfop_filters = self.parse_filter_tokens(self.cfop_input.text())
        status_filter = self.status_combo.currentText().strip()
        search_text = self.normalize_search_text(self.search_input.text())
        filtered: list[dict[str, object]] = []
        for row in self.entry_rows:
            period = str(self.first_row_value(row, "period", "period_label"))
            if selected_periods and period not in selected_periods:
                continue
            row_csts = {part.strip() for part in str(self.first_row_value(row, "cst_icms", "cst")).split("|") if part.strip()}
            if cst_filters and not row_csts.intersection(cst_filters):
                continue
            row_cfops = {part.strip() for part in str(row.get("cfop", "")).split("|") if part.strip()}
            if cfop_filters and not row_cfops.intersection(cfop_filters):
                continue
            status_value = str(row.get("status", row.get("diagnostic_status", ""))).strip()
            if status_filter == "Ok" and status_value != "Ok":
                continue
            if status_filter == "Sem credito" and "Sem credito" not in status_value:
                continue
            if status_filter == "Sem entrada" and status_value != "Sem entrada":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem entrada"} and "Sem credito" not in status_value:
                continue
            if search_text:
                searchable = self.normalize_search_text(
                    f"{row.get('code', '')} {row.get('description', '')} {row.get('cest', '')} {row.get('suppliers', '')} "
                    f"{row.get('supplier', '')} {row.get('participant_name', '')}"
                )
                if search_text not in searchable:
                    continue
            filtered.append(row)
        return filtered

    def refresh_entry_table(self) -> None:
        self.filtered_entry_rows = self.filtered_rows()
        self.entry_table.setRowCount(len(self.filtered_entry_rows))
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        products: set[str] = set()
        for row_index, row in enumerate(self.filtered_entry_rows):
            operation = self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            base = self.decimal_value(row.get("base_icms"))
            icms = self.decimal_value(row.get("icms_value"))
            total_operation += operation
            total_base += base
            total_icms += icms
            code = str(row.get("code", "")).strip()
            if code:
                products.add(code)
            values = [
                self.first_row_value(row, "period", "period_label"),
                self.first_row_value(row, "curve_abc", "abc_curve", default="C"),
                code,
                row.get("description", ""),
                self.format_supplier_text(row),
                row.get("ncm", ""),
                row.get("cest", ""),
                row.get("cfop", ""),
                self.first_row_value(row, "cst_icms", "cst"),
                self.format_number(self.first_row_value(row, "icms_rate", "nominal_icms_rate")),
                self.format_number(self.first_row_value(row, "effective_icms_rate", "icms_effective_rate")),
                self.first_row_value(row, "document_count", "documents", default=0),
                self.first_row_value(row, "launch_count", "launches", default=0),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if column in {0, 1, 2, 5, 6, 7, 8, 9, 10, 11, 12}:
                    item.setTextAlignment(Qt.AlignCenter)
                if str(row.get("status", "")).strip() not in {"", "Ok"}:
                    item.setBackground(QColor("#fff8e8"))
                self.entry_table.setItem(row_index, column, item)
        self.metric_labels["rows"].setText(str(len(self.filtered_entry_rows)))
        self.metric_labels["products"].setText(str(len(products)))
        self.metric_labels["operation"].setText(self.format_number(total_operation))
        self.metric_labels["base"].setText(self.format_number(total_base))
        self.metric_labels["icms"].setText(self.format_number(total_icms))

    def format_supplier_text(self, row: dict[str, object]) -> str:
        supplier = str(self.first_row_value(row, "suppliers", "supplier", "participant_name")).strip()
        count = self.first_row_value(row, "supplier_count", default="")
        try:
            count_number = int(count or 0)
        except (TypeError, ValueError):
            count_number = 0
        if count_number > 1:
            return f"[MULTIPLOS: {count_number}] {supplier}" if supplier else f"[MULTIPLOS: {count_number}]"
        return supplier

    def format_number(self, value: object) -> str:
        if isinstance(value, Decimal):
            return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        text = str(value if value is not None else "").strip()
        return text.replace(".", ",") if text else "0,00"

    def clear_entry_filters(self) -> None:
        self.cst_input.clear()
        self.cfop_input.clear()
        self.status_combo.setCurrentText("Todos")
        self.search_input.clear()
        for check in self.period_checks.values():
            check.setChecked(True)
        self.refresh_entry_table()

    def clear_entry_page(self) -> None:
        self.sped_input.clear()
        self.xml_input.clear()
        self.entry_rows = []
        self.filtered_entry_rows = []
        self.rebuild_period_checks([])
        self.refresh_entry_table()
        self.statusBar().showMessage("Tela limpa.")

    def export_entry_filter(self) -> None:
        if not self.filtered_entry_rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar consulta filtrada",
            "consulta_entradas_filtrada.xlsx",
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        headers = ["Periodo", "Curva ABC", "Codigo", "Descricao", "Fornecedor", "NCM", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Docs", "Lancamentos", "Valor Operacao", "Base ICMS", "Valor ICMS", "Status"]
        rows = [
            [
                self.first_row_value(row, "period", "period_label"),
                self.first_row_value(row, "curve_abc", "abc_curve", default="C"),
                row.get("code", ""),
                row.get("description", ""),
                self.format_supplier_text(row),
                row.get("ncm", ""),
                row.get("cest", ""),
                row.get("cfop", ""),
                self.first_row_value(row, "cst_icms", "cst"),
                self.decimal_value(self.first_row_value(row, "icms_rate", "nominal_icms_rate")),
                self.decimal_value(self.first_row_value(row, "effective_icms_rate", "icms_effective_rate")),
                self.integer_from_display_value(self.first_row_value(row, "document_count", "documents", default=0)),
                self.integer_from_display_value(self.first_row_value(row, "launch_count", "launches", default=0)),
                self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value")),
                self.decimal_value(row.get("base_icms")),
                self.decimal_value(row.get("icms_value")),
                row.get("status", ""),
            ]
            for row in self.filtered_entry_rows
        ]
        total_row = self.build_popup_total_row(headers, rows)
        rows_to_write = rows + ([total_row] if total_row else [])
        output_path = Path(output)
        if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
            if output_path.suffix.lower() != ".csv":
                output_path = output_path.with_suffix(".csv")
            write_simple_csv_file(output_path, headers, rows_to_write)
        else:
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            write_simple_excel_workbook(output_path, [("Consulta Entradas", headers, rows_to_write, {"include_total": False})])
        self.statusBar().showMessage(f"Consulta exportada: {output_path}")

    def open_table_popup(
        self,
        title: str,
        headers: list[str],
        rows: list[list[object]] | list[tuple[object, ...]],
        width: int = 1180,
        height: int = 640,
        footer_text: str = "",
    ) -> None:
        if not rows:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        export_rows = [list(row) for row in rows]
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        dialog.resize(width, height)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)

        table = QTableWidget()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(export_rows))
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setWordWrap(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().setStretchLastSection(False)
        for row_index, row in enumerate(export_rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if not (column_index < len(headers) and headers[column_index].lower() in {"descricao", "fornecedor", "chave"}):
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        self.apply_popup_column_widths(table, headers)
        layout.addWidget(table, 1)
        if footer_text:
            footer = QLabel(footer_text)
            footer.setObjectName("popupFooter")
            layout.addWidget(footer)
        close_row = QHBoxLayout()
        close_row.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, headers, export_rows, "xlsx")))
        close_row.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, headers, export_rows, "csv")))
        close_row.addStretch()
        close_button = self.create_button("Fechar", dialog.close)
        close_row.addWidget(close_button)
        layout.addLayout(close_row)
        dialog.exec()

    def apply_popup_column_widths(self, table: QTableWidget, headers: list[str]) -> None:
        width_by_header = {
            "CST": 78,
            "CFOP": 86,
            "Aliq ICMS": 105,
            "Aliq Efetiva": 110,
            "Valor IPI": 105,
            "Valor ICMS": 112,
            "Base ICMS ST": 125,
            "Valor ICMS ST": 125,
            "Total Operacao": 130,
            "Base ICMS": 120,
            "Dif. Oper/Base": 125,
            "Reducao BC": 115,
            "Docs": 70,
            "Lanc.": 70,
            "Periodo": 95,
            "Codigo": 110,
            "Descricao": 330,
            "Fornecedor": 280,
            "Status": 130,
            "Chave": 310,
        }
        for index, header in enumerate(headers):
            table.setColumnWidth(index, width_by_header.get(header, 135))

    def export_popup_dataset(self, title: str, headers: list[str], rows: list[list[object]], output_type: str) -> None:
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "Exportar popup",
            f"{self.slugify_filename(title)}{suffix}",
            "Arquivo CSV (*.csv)" if output_type == "csv" else "Arquivo Excel (*.xlsx)",
        )
        if not selected:
            return
        output_path = Path(selected)
        export_rows = self.prepare_popup_export_rows(headers, rows)
        total_row = self.build_popup_total_row(headers, export_rows)
        rows_to_write = export_rows + ([total_row] if total_row else [])
        try:
            if output_type == "csv":
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                write_simple_excel_workbook(output_path, [(title[:31] or "Popup", headers, rows_to_write, {"include_total": False})])
            self.statusBar().showMessage(f"Popup exportado: {output_path}")
            QMessageBox.information(self, "Exportar popup", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Exportar popup", str(exc))

    def prepare_popup_export_rows(self, headers: list[str], rows: list[list[object]]) -> list[list[object]]:
        return [
            [self.normalize_popup_export_value(headers[column_index], value) for column_index, value in enumerate(row)]
            for row in rows
        ]

    def normalize_popup_export_value(self, header: str, value: object) -> object:
        if isinstance(value, Decimal) or isinstance(value, int):
            return value
        if self.popup_header_is_integer(header):
            return self.integer_from_display_value(value)
        if self.popup_header_is_decimal(header):
            return self.decimal_from_display_value(value)
        return value

    def popup_header_is_integer(self, header: str) -> bool:
        normalized = self.normalize_search_text(header)
        return normalized in {"docs", "lanc", "lancamentos", "produtos", "linhas", "periodos"}

    def popup_header_is_decimal(self, header: str) -> bool:
        normalized = self.normalize_search_text(header)
        if normalized in {"cst", "cfop", "codigo", "periodo", "item", "documento", "chave"}:
            return False
        decimal_tokens = (
            "aliq",
            "valor",
            "base",
            "total",
            "operacao",
            "icms",
            "ipi",
            "st",
            "reducao",
            "dif",
            "credito",
            "debito",
            "%",
        )
        return any(token in normalized for token in decimal_tokens)

    def decimal_from_display_value(self, value: object) -> Decimal:
        if isinstance(value, Decimal):
            return value
        text = str(value or "").strip()
        if not text:
            return Decimal("0")
        text = text.replace("R$", "").replace("%", "").strip()
        text = text.replace(".", "").replace(",", ".")
        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            return Decimal("0")

    def integer_from_display_value(self, value: object) -> int:
        if isinstance(value, int):
            return value
        try:
            return int(self.decimal_from_display_value(value))
        except (InvalidOperation, ValueError):
            return 0

    def build_popup_total_row(self, headers: list[str], rows: list[list[object]]) -> list[object]:
        if not rows:
            return []
        total_row: list[object] = [""] * len(headers)
        label_index = next((index for index, header in enumerate(headers) if not self.popup_header_is_decimal(header) and not self.popup_header_is_integer(header)), 0)
        total_row[label_index] = "Total"
        has_total = False
        for column_index, header in enumerate(headers):
            if not self.popup_header_is_decimal(header) and not self.popup_header_is_integer(header):
                continue
            values = [row[column_index] for row in rows if column_index < len(row)]
            if not values:
                continue
            if self.popup_header_is_integer(header):
                total_row[column_index] = sum(int(value or 0) for value in values)
            else:
                total_row[column_index] = sum((self.decimal_value(value) for value in values), Decimal("0"))
            has_total = True
        return total_row if has_total else []

    def slugify_filename(self, value: str) -> str:
        normalized = self.normalize_search_text(value).replace(" ", "_")
        cleaned = "".join(char for char in normalized if char.isalnum() or char in {"_", "-"}).strip("_")
        return cleaned or "popup"

    def get_filtered_launch_details(self) -> list[dict[str, object]]:
        details: list[dict[str, object]] = []
        for row in self.filtered_entry_rows:
            row_period = str(self.first_row_value(row, "period", "period_label"))
            launch_details = row.get("launch_details")
            if isinstance(launch_details, list) and launch_details:
                for detail in launch_details:
                    if not isinstance(detail, dict):
                        continue
                    enriched = dict(detail)
                    enriched.setdefault("period", row_period)
                    enriched.setdefault("operation_type", "Entrada")
                    enriched.setdefault("source_register", "C170")
                    details.append(enriched)
                continue
            fallback = dict(row)
            fallback.setdefault("period", row_period)
            fallback.setdefault("operation_type", "Entrada")
            fallback.setdefault("source_register", "C170")
            fallback.setdefault("total_operation_value", self.first_row_value(row, "sale_value", "total_operation_value"))
            details.append(fallback)
        return details

    def get_filtered_summary_rows_from_details(self) -> list[dict[str, object]]:
        summary_rows: list[dict[str, object]] = []
        for detail in self.get_filtered_launch_details():
            sale_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(detail.get("base_icms"))
            icms_value = self.decimal_value(detail.get("icms_value"))
            summary_rows.append(
                {
                    "operation_type": "Entrada",
                    "period": detail.get("period", ""),
                    "source_register": detail.get("source_register", "C170") or "C170",
                    "cst_icms": self.first_row_value(detail, "cst_icms", "cst"),
                    "cfop": detail.get("cfop", ""),
                    "icms_rate": self.decimal_value(detail.get("icms_rate")).quantize(Decimal("0.01")),
                    "total_operation_value": sale_value,
                    "base_icms": base_icms,
                    "icms_value": icms_value,
                    "base_icms_st": self.decimal_value(detail.get("base_icms_st")),
                    "icms_st_value": self.decimal_value(detail.get("icms_st_value")),
                    "reduction_value": max(Decimal("0"), sale_value - base_icms),
                    "document_key": detail.get("document_key", ""),
                    "document_number": detail.get("document_number", ""),
                }
            )
        return summary_rows

    def open_multi_table_popup(
        self,
        title: str,
        sections: list[tuple[str, list[str], list[list[object]]]],
        footer_text: str = "",
        width: int = 1280,
        height: int = 720,
    ) -> None:
        non_empty_sections = [(name, headers, rows) for name, headers, rows in sections if rows]
        if not non_empty_sections:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        dialog.resize(width, height)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)
        for section_name, headers, rows in non_empty_sections:
            section_label = QLabel(section_name)
            section_label.setObjectName("sectionTitle")
            layout.addWidget(section_label)
            table = QTableWidget()
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)
            table.setRowCount(len(rows))
            table.setAlternatingRowColors(True)
            table.setSelectionBehavior(QAbstractItemView.SelectRows)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
            table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
            table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
            table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
            table.setWordWrap(False)
            table.verticalHeader().setVisible(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            table.horizontalHeader().setStretchLastSection(False)
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    item.setForeground(QColor(COLORS["text"]))
                    if not (column_index < len(headers) and headers[column_index].lower() in {"descricao", "fornecedor", "chave", "motivo"}):
                        item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(row_index, column_index, item)
            self.apply_popup_column_widths(table, headers)
            layout.addWidget(table, 1)
        if footer_text:
            footer = QLabel(footer_text)
            footer.setObjectName("popupFooter")
            layout.addWidget(footer)
        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_multi_popup_dataset(title, non_empty_sections, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_multi_popup_dataset(title, non_empty_sections, "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def export_multi_popup_dataset(self, title: str, sections: list[tuple[str, list[str], list[list[object]]]], output_type: str) -> None:
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "Exportar popup",
            f"{self.slugify_filename(title)}{suffix}",
            "Arquivo CSV (*.csv)" if output_type == "csv" else "Arquivo Excel (*.xlsx)",
        )
        if not selected:
            return
        output_path = Path(selected)
        try:
            if output_type == "xlsx":
                sheets = []
                for section_name, headers, rows in sections:
                    export_rows = self.prepare_popup_export_rows(headers, rows)
                    total_row = self.build_popup_total_row(headers, export_rows)
                    sheets.append((section_name[:31] or "Dados", headers, export_rows + ([total_row] if total_row else []), {"include_total": False}))
                write_simple_excel_workbook(output_path, sheets)
            else:
                csv_rows: list[list[object]] = []
                csv_headers = ["Secao", "Campo"] + [f"Valor {index}" for index in range(1, max((len(headers) for _, headers, _ in sections), default=0))]
                for section_name, headers, rows in sections:
                    csv_rows.append([section_name, *headers])
                    export_rows = self.prepare_popup_export_rows(headers, rows)
                    total_row = self.build_popup_total_row(headers, export_rows)
                    for row in export_rows + ([total_row] if total_row else []):
                        csv_rows.append(["", *row])
                    csv_rows.append([])
                write_simple_csv_file(output_path, csv_headers, csv_rows)
            self.statusBar().showMessage(f"Popup exportado: {output_path}")
            QMessageBox.information(self, "Exportar popup", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Exportar popup", str(exc))

    def open_entry_products_popup(self) -> None:
        rows = [
            (
                self.first_row_value(row, "period", "period_label"),
                row.get("code", ""),
                row.get("description", ""),
                self.format_supplier_text(row),
                row.get("cfop", ""),
                self.first_row_value(row, "cst_icms", "cst"),
                self.format_number(self.first_row_value(row, "sale_value", "total_operation_value")),
                self.format_number(row.get("base_icms")),
                self.format_number(row.get("icms_value")),
                row.get("status", ""),
            )
            for row in self.filtered_entry_rows
        ]
        self.open_table_popup(
            "Entradas filtradas",
            ["Periodo", "Codigo", "Descricao", "Fornecedor", "CFOP", "CST", "Valor Operacao", "Base ICMS", "Valor ICMS", "Status"],
            rows,
            1280,
            680,
        )

    def open_entry_operation_summary_popup(self) -> None:
        grouped: dict[tuple[str, str, Decimal], dict[str, object]] = {}
        for row in self.filtered_entry_rows:
            details = row.get("launch_details")
            source_details = details if isinstance(details, list) and details else [row]
            for detail in source_details:
                if not isinstance(detail, dict):
                    continue
                cst = str(self.first_row_value(detail, "cst_icms", "cst")).strip()
                cfop = str(detail.get("cfop", "")).strip()
                rate = self.decimal_value(detail.get("icms_rate")).quantize(Decimal("0.01"))
                key = (cst, cfop, rate)
                bucket = grouped.setdefault(
                    key,
                    {
                        "sale_value": Decimal("0"),
                        "base_icms": Decimal("0"),
                        "icms_value": Decimal("0"),
                        "base_icms_st": Decimal("0"),
                        "icms_st_value": Decimal("0"),
                        "ipi_value": Decimal("0"),
                        "document_keys": set(),
                        "launch_count": 0,
                    },
                )
                sale_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
                base_icms = self.decimal_value(detail.get("base_icms"))
                icms_value = self.decimal_value(detail.get("icms_value"))
                bucket["sale_value"] = Decimal(bucket["sale_value"]) + sale_value
                bucket["base_icms"] = Decimal(bucket["base_icms"]) + base_icms
                bucket["icms_value"] = Decimal(bucket["icms_value"]) + icms_value
                bucket["base_icms_st"] = Decimal(bucket["base_icms_st"]) + self.decimal_value(detail.get("base_icms_st"))
                bucket["icms_st_value"] = Decimal(bucket["icms_st_value"]) + self.decimal_value(detail.get("icms_st_value"))
                bucket["ipi_value"] = Decimal(bucket["ipi_value"]) + self.decimal_value(detail.get("ipi_value"))
                document_key = str(detail.get("document_key", "") or detail.get("document_number", "")).strip()
                if document_key:
                    bucket["document_keys"].add(document_key)
                bucket["launch_count"] = int(bucket["launch_count"]) + int(self.first_row_value(detail, "launch_count", default=1) or 1)

        rows: list[tuple[object, ...]] = []
        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for (cst, cfop, rate), values in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1], item[0][2])):
            sale_value = Decimal(values["sale_value"])
            base_icms = Decimal(values["base_icms"])
            icms_value = Decimal(values["icms_value"])
            effective_rate = compute_display_icms_rate(rate, sale_value, base_icms, icms_value)
            rows.append(
                (
                    cst,
                    cfop,
                    self.format_number(rate),
                    self.format_number(effective_rate),
                    self.format_number(Decimal(values["ipi_value"])),
                    self.format_number(icms_value),
                    self.format_number(Decimal(values["base_icms_st"])),
                    self.format_number(Decimal(values["icms_st_value"])),
                    self.format_number(sale_value),
                    self.format_number(base_icms),
                    self.format_number((sale_value - base_icms).quantize(Decimal("0.01"))),
                    self.format_number((sale_value - base_icms).quantize(Decimal("0.01"))),
                    len(values["document_keys"]),
                    values["launch_count"],
                )
            )
            total_sale += sale_value
            total_base += base_icms
            total_icms += icms_value

        ratio = (total_base * Decimal("100") / total_sale).quantize(Decimal("0.01")) if total_sale else Decimal("0.00")
        footer = (
            f"Linhas: {len(rows)}    Total Operacao: {self.format_number(total_sale)}    "
            f"Base ICMS: {self.format_number(total_base)}    % Base/Oper: {self.format_number(ratio)}    "
            f"Valor ICMS: {self.format_number(total_icms)}"
        )
        self.open_table_popup(
            "Resumo Entradas",
            [
                "CST",
                "CFOP",
                "Aliq ICMS",
                "Aliq Efetiva",
                "Valor IPI",
                "Valor ICMS",
                "Base ICMS ST",
                "Valor ICMS ST",
                "Total Operacao",
                "Base ICMS",
                "Dif. Oper/Base",
                "Reducao BC",
                "Docs",
                "Lanc.",
            ],
            rows,
            1220,
            560,
            footer,
        )

    def open_entry_abc_popup(self) -> None:
        _periods, headers, _display_rows, export_rows = build_product_monthly_linear_dataset(self.filtered_entry_rows, "Entrada")
        self.open_table_popup("Levantamento por Produto - Entrada", headers, export_rows, 1360, 620)

    def open_entry_reduction_popup(self) -> None:
        reduction_rows = build_reduction_launch_rows(self.get_filtered_launch_details())
        rows = [
            (
                row.get("operation_type", ""),
                row.get("document_date", ""),
                row.get("document_number", ""),
                row.get("participant_name", ""),
                row.get("document_key", ""),
                row.get("item_number", ""),
                row.get("code", ""),
                row.get("description", ""),
                row.get("cfop", ""),
                row.get("cst_icms", ""),
                row.get("icms_rate", Decimal("0")),
                compute_display_icms_rate(row.get("icms_rate", Decimal("0")), row.get("sale_value", Decimal("0")), row.get("base_icms", Decimal("0")), row.get("icms_value", Decimal("0"))),
                row.get("sale_value", Decimal("0")),
                row.get("base_icms", Decimal("0")),
                (self.decimal_value(row.get("base_icms")) * Decimal("100") / self.decimal_value(row.get("sale_value"))).quantize(Decimal("0.01")) if self.decimal_value(row.get("sale_value")) > 0 else Decimal("0.00"),
                row.get("reduction_value", Decimal("0")),
                row.get("icms_value", Decimal("0")),
            )
            for row in reduction_rows
        ]
        self.open_table_popup(
            "Itens com Reducao de Base",
            ["Tipo", "Data", "Documento", "Fornecedor", "Chave", "Item", "Codigo", "Descricao", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Valor Operacao", "Base ICMS", "% Base/Oper", "Reducao BC", "Valor ICMS"],
            rows,
            1320,
            720,
        )

    def open_entry_apuracao_popup(self) -> None:
        entry_totals = {
            "icms_value": sum((self.decimal_value(row.get("icms_value")) for row in self.filtered_entry_rows), Decimal("0")),
        }
        exit_totals = {"icms_value": Decimal("0")}
        original_summary = {
            "debitos": Decimal("0"),
            "ajustes_debito_doc": Decimal("0"),
            "ajustes_debito": Decimal("0"),
            "estornos_credito": Decimal("0"),
            "creditos": Decimal("0"),
            "ajustes_credito_doc": Decimal("0"),
            "ajustes_credito": Decimal("0"),
            "estornos_debito": Decimal("0"),
            "saldo_credor_anterior": Decimal("0"),
            "saldo_devedor": Decimal("0"),
            "deducoes": Decimal("0"),
            "icms_recolher": Decimal("0"),
            "saldo_credor_transportar": Decimal("0"),
            "extra_apuracao": Decimal("0"),
        }
        rows = build_filtered_apuracao_rows(entry_totals, exit_totals, original_summary, False)
        self.open_table_popup("Apuracao do ICMS", ["Descricao", "Valor R$"], rows, 1020, 520)

    def open_entry_credit_diagnostic_popup(self) -> None:
        summary_rows, detail_rows, totals = build_credit_diagnostic_datasets(
            self.get_filtered_summary_rows_from_details(),
            "Entrada",
            self.parse_filter_tokens(self.cst_input.text()),
            self.parse_filter_tokens(self.cfop_input.text()),
            self.get_filtered_launch_details(),
        )
        summary_headers = ["Periodo", "Motivo", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        summary_export_rows = [
            [row["period"], row["reason"], row["row_count"], row["total_operation_value"], row["base_icms"], row["base_gap"], row["base_ratio"], row["icms_value"]]
            for row in summary_rows
        ]
        detail_headers = ["Periodo", "Motivo", "Registro", "CST", "CFOP", "Aliq", "Aliq Efetiva", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        detail_export_rows = [
            [
                row["period"],
                row["reason"],
                row["source_register"],
                row["cst_icms"],
                row["cfop"],
                row["icms_rate"],
                compute_display_icms_rate(row["icms_rate"], row["total_operation_value"], row["base_icms"], row["icms_value"]),
                row["row_count"],
                row["total_operation_value"],
                row["base_icms"],
                row["base_gap"],
                row["base_ratio"],
                row["icms_value"],
            ]
            for row in detail_rows
        ]
        self.open_multi_table_popup(
            "Diagnostico Credito ICMS - Entradas",
            [("Resumo dos Motivos", summary_headers, summary_export_rows), ("Detalhamento Fiscal", detail_headers, detail_export_rows)],
            footer_text=(
                f"Valor Operacao: {self.format_number(totals.get('total_operation_value', Decimal('0')))}    "
                f"Base ICMS: {self.format_number(totals.get('base_icms', Decimal('0')))}    "
                f"Perda de Base: {self.format_number(totals.get('base_gap', Decimal('0')))}    "
                f"Valor ICMS: {self.format_number(totals.get('icms_value', Decimal('0')))}"
            ),
            width=1380,
            height=760,
        )

    def open_entry_credit_comparison_popup(self) -> None:
        periods = sorted({str(self.first_row_value(row, "period", "period_label")) for row in self.filtered_entry_rows if str(self.first_row_value(row, "period", "period_label")).strip()})
        if len(periods) < 2:
            QMessageBox.warning(self, "Comparacao diagnostico de credito", "Selecione ao menos dois SPEDs/periodos para comparar.")
            return
        _periods, headers, display_rows, export_rows, _detail_rows = build_credit_diagnostic_period_comparison_dataset(
            self.get_filtered_summary_rows_from_details(),
            self.get_filtered_launch_details(),
            self.parse_filter_tokens(self.cst_input.text()),
            self.parse_filter_tokens(self.cfop_input.text()),
            "Entrada",
        )
        rows = export_rows or [list(row.get("values", [])) for row in display_rows]
        self.open_table_popup("Comparacao Diagnostico de Credito", headers, rows, 1560, 760)

    def open_entry_docs_popup(self) -> None:
        grouped: dict[tuple[str, str, str], dict[str, object]] = {}
        for detail in self.get_filtered_launch_details():
            key = (
                str(detail.get("period", "")),
                str(detail.get("document_number", "")),
                str(detail.get("document_key", "")),
            )
            bucket = grouped.setdefault(
                key,
                {
                    "period": detail.get("period", ""),
                    "document_date": detail.get("document_date", detail.get("date", "")),
                    "document_number": detail.get("document_number", ""),
                    "document_series": detail.get("document_series", ""),
                    "document_model": detail.get("document_model", ""),
                    "participant_name": detail.get("participant_name", ""),
                    "participant_tax_id": detail.get("participant_tax_id", ""),
                    "document_key": detail.get("document_key", ""),
                    "items": 0,
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            bucket["items"] = int(bucket["items"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            bucket["base_icms"] = Decimal(bucket["base_icms"]) + self.decimal_value(detail.get("base_icms"))
            bucket["icms_value"] = Decimal(bucket["icms_value"]) + self.decimal_value(detail.get("icms_value"))
        rows = [
            (
                row["period"],
                row["document_date"],
                row["document_number"],
                row["document_series"],
                row["document_model"],
                row["participant_name"],
                row["participant_tax_id"],
                row["document_key"],
                row["items"],
                row["sale_value"],
                row["base_icms"],
                row["icms_value"],
            )
            for row in sorted(grouped.values(), key=lambda item: (str(item["period"]), str(item["document_date"]), str(item["document_number"])))
        ]
        self.open_table_popup(
            "Espelho de Documentos Fiscais - Entradas",
            ["Periodo", "Data", "Documento", "Serie", "Modelo", "Fornecedor", "CPF/CNPJ", "Chave", "Itens", "Valor Operacao", "Base ICMS", "Valor ICMS"],
            rows,
            1380,
            720,
        )
