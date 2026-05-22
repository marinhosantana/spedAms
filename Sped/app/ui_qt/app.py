from __future__ import annotations

import csv
import datetime as dt
import os
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QThread, Qt, Signal
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
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.config import MYSQL_DEFAULT_CONFIG
from app.exporters.workbook_exporter import write_simple_csv_file, write_simple_excel_workbook
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.parsers.sped_fiscal_parser import read_sped_file
from app.services.app_paths import get_application_base_dir, get_application_environment, get_environment_config_path, get_project_root_dir
from app.services.compare_operations import filter_xml_summary_rows_by_scope, normalize_compare_operation_scope
from app.services.compare_workflows import build_xml_cfop_summary_rows, build_xml_entry_credit_rows, compare_sped_with_sheet, compare_sped_with_xml_folder
from app.services.analysis_reports import (
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_contrib_operation_launch_details_map,
    build_contrib_operation_summary_rows,
    build_contrib_product_monthly_linear_dataset,
    build_product_monthly_linear_dataset,
)
from app.services.operation_summary import build_filtered_apuracao_rows, build_reduction_launch_rows
from app.services.path_selection import append_unique_paths, collapse_xml_selection_paths, format_selected_paths, limit_selected_paths, parse_selected_paths
from app.services.period_comparisons import build_entry_period_comparison_rows, build_sale_period_comparison_rows
from app.services.sped_archive import archive_original_sped_file
from app.services.tax_rules import compute_display_icms_rate
from app.services.runtime_rules import parse_runtime_rule_lines, runtime_rule_summary
from app.services.xml_reconciliation import build_pis_cofins_period_comparison_rows


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


class CompareWorker(QObject):
    finished = Signal(list, list, list, dict, str, str, str)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, sped_path: Path, source_path: Path, compare_mode: str, source_kind: str, operation_scope: str) -> None:
        super().__init__()
        self.sped_path = sped_path
        self.source_path = source_path
        self.compare_mode = compare_mode
        self.source_kind = source_kind
        self.operation_scope = operation_scope

    def run(self) -> None:
        try:
            if self.source_kind == "sheet":
                present, missing, stats = compare_sped_with_sheet(
                    self.sped_path,
                    self.source_path,
                    self.operation_scope,
                    self.progress.emit,
                )
                sped_missing_xml: list[object] = []
            else:
                present, missing, sped_missing_xml, stats = compare_sped_with_xml_folder(
                    self.sped_path,
                    self.source_path,
                    self.operation_scope,
                    self.progress.emit,
                )
            self.finished.emit(present, missing, sped_missing_xml, stats, self.compare_mode, self.source_kind, self.operation_scope)
        except Exception as exc:
            self.failed.emit(str(exc))


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
        self.sale_rows: list[dict[str, object]] = []
        self.filtered_sale_rows: list[dict[str, object]] = []
        self.sale_period_checks: dict[str, QCheckBox] = {}
        self.contrib_entry_rows: list[dict[str, object]] = []
        self.filtered_contrib_entry_rows: list[dict[str, object]] = []
        self.contrib_entry_period_checks: dict[str, QCheckBox] = {}
        self.contrib_sale_rows: list[dict[str, object]] = []
        self.filtered_contrib_sale_rows: list[dict[str, object]] = []
        self.contrib_sale_period_checks: dict[str, QCheckBox] = {}
        self.xml_summary_all_rows: list[dict[str, object]] = []
        self.xml_summary_rows: list[dict[str, object]] = []
        self.xml_summary_stats: dict[str, object] = {}
        self.xml_credit_invoice_rows: list[dict[str, object]] = []
        self.xml_credit_cfop_rows: list[dict[str, object]] = []
        self.compare_rows: list[list[object]] = []
        self.compare_stats: dict[str, object] = {}
        self.archive_profile_rows: dict[int, dict[str, object]] = {}
        self.archive_file_rows: dict[int, dict[str, object]] = {}

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
        for title, index in (
            ("Dashboard", 0),
            ("Consulta Entradas", 1),
            ("Consulta Saidas", 2),
            ("PIS/COFINS Entradas", 3),
            ("PIS/COFINS Saidas", 4),
            ("XML", 5),
            ("SPED x XML", 6),
            ("Regras Dinamicas", 7),
            ("SPEDs Arquivados", 8),
            ("Configuracoes", 9),
        ):
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
        self.stack.addWidget(self.build_sale_page())
        self.stack.addWidget(self.build_contrib_page("Entrada"))
        self.stack.addWidget(self.build_contrib_page("Saida"))
        self.stack.addWidget(self.build_xml_page())
        self.stack.addWidget(self.build_compare_page())
        self.stack.addWidget(self.build_runtime_rules_page())
        self.stack.addWidget(self.build_archives_page())
        self.stack.addWidget(self.build_settings_page())
        content_layout.addWidget(self.stack, 1)
        shell.addWidget(content, 1)
        self.setCentralWidget(root)
        self.statusBar().showMessage("Pronto.")

    def show_page(self, index: int, title: str) -> None:
        self.stack.setCurrentIndex(index)
        self.page_title.setText(title)
        for button_index, button in enumerate(self.nav_buttons):
            button.setChecked(button_index == index)
        if title == "SPEDs Arquivados":
            self.refresh_archives()

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

    def build_archives_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("SPEDs Arquivados")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Importar SPED", self.import_speds, primary=True))
        toolbar_layout.addWidget(self.create_button("Atualizar", self.refresh_archives))
        toolbar_layout.addWidget(self.create_button("Abrir Pasta", self.open_selected_archive_folder))
        layout.addWidget(toolbar)

        split = QHBoxLayout()
        split.setSpacing(12)
        self.archive_profile_table = self.create_data_table(["ID", "Empresa", "CNPJ/CPF", "Periodo", "Arquivos"])
        self.archive_profile_table.setColumnWidth(0, 60)
        self.archive_profile_table.setColumnWidth(1, 330)
        self.archive_profile_table.setColumnWidth(2, 150)
        self.archive_profile_table.setColumnWidth(3, 170)
        self.archive_profile_table.setColumnWidth(4, 85)
        self.archive_profile_table.itemSelectionChanged.connect(self.refresh_archive_files)
        self.archive_file_table = self.create_data_table(["ID", "Tipo", "Arquivo", "Periodo", "0200", "Docs", "C170", "C190", "Hash"])
        self.archive_file_table.setColumnWidth(0, 60)
        self.archive_file_table.setColumnWidth(1, 100)
        self.archive_file_table.setColumnWidth(2, 300)
        self.archive_file_table.setColumnWidth(3, 170)
        self.archive_file_table.setColumnWidth(8, 130)
        split.addWidget(self.archive_profile_table, 1)
        split.addWidget(self.archive_file_table, 2)
        layout.addLayout(split, 1)
        return page

    def build_settings_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        panel = self.create_panel()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        title = QLabel("Configuracoes")
        title.setObjectName("sectionTitle")
        config_label = QLabel(f"Config MySQL: {self.mysql_config_path}")
        config_label.setObjectName("muted")
        panel_layout.addWidget(title)
        panel_layout.addWidget(config_label)
        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Criar/Atualizar Banco", self.create_schema, primary=True))
        actions.addWidget(self.create_button("Testar Conexao", self.test_mysql_connection))
        actions.addStretch()
        panel_layout.addLayout(actions)
        layout.addWidget(panel)
        layout.addStretch()
        return page

    def create_data_table(self, headers: list[str]) -> QTableWidget:
        table = QTableWidget()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
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
        return table

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

    def build_sale_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Consulta Saidas ICMS")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_sale_page))
        toolbar_layout.addWidget(self.create_button("Exportar", self.export_sale_filter))
        toolbar_layout.addWidget(self.create_button("Processar", self.process_sale_query, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.sale_sped_input = QLineEdit()
        self.sale_xml_input = QLineEdit()
        self.add_path_row(setup_layout, 0, "SPEDs", self.sale_sped_input, self.select_sale_sped_files, lambda: self.sale_sped_input.clear())
        self.add_path_row(setup_layout, 1, "XMLs 55/65", self.sale_xml_input, self.select_sale_xml_sources, lambda: self.sale_xml_input.clear())
        layout.addWidget(setup)

        filters = self.create_panel()
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(12, 12, 12, 12)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(8)
        self.sale_periods_widget = QWidget()
        self.sale_periods_layout = QHBoxLayout(self.sale_periods_widget)
        self.sale_periods_layout.setContentsMargins(0, 0, 0, 0)
        self.sale_periods_layout.setSpacing(8)
        self.sale_cst_input = QLineEdit()
        self.sale_cst_input.setPlaceholderText("000, 020, 060")
        self.sale_cfop_input = QLineEdit()
        self.sale_cfop_input.setPlaceholderText("5101, 5405, 6108")
        self.sale_status_combo = QComboBox()
        self.sale_status_combo.addItems(["Todos", "Ok", "Sem debito", "Sem saida", "Com divergencia"])
        self.sale_search_input = QLineEdit()
        self.sale_search_input.setPlaceholderText("Codigo, descricao ou cliente")
        self.add_filter_box(filters_layout, 0, 0, "Periodos", self.sale_periods_widget)
        self.add_filter_box(filters_layout, 0, 1, "CST", self.sale_cst_input)
        self.add_filter_box(filters_layout, 0, 2, "CFOP", self.sale_cfop_input)
        search_box = QWidget()
        search_layout = QVBoxLayout(search_box)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)
        search_layout.addWidget(self.sale_status_combo)
        search_layout.addWidget(self.sale_search_input)
        self.add_filter_box(filters_layout, 0, 3, "Status / Busca", search_box)
        filters_layout.setColumnStretch(0, 2)
        filters_layout.setColumnStretch(1, 1)
        filters_layout.setColumnStretch(2, 1)
        filters_layout.setColumnStretch(3, 2)
        layout.addWidget(filters)

        filter_actions = QHBoxLayout()
        filter_actions.addWidget(self.create_button("Aplicar filtros", self.refresh_sale_table, primary=True))
        filter_actions.addWidget(self.create_button("Limpar filtros", self.clear_sale_filters))
        filter_actions.addWidget(self.create_button("Exportar filtro atual", self.export_sale_filter))
        filter_actions.addWidget(self.create_button("Saidas", self.open_sale_operation_summary_popup))
        filter_actions.addWidget(self.create_button("Comp. Diag. Debito", self.open_sale_debit_comparison_popup))
        filter_actions.addWidget(self.create_button("Diag. Debito", self.open_sale_debit_diagnostic_popup))
        filter_actions.addWidget(self.create_button("Curva ABC", self.open_sale_abc_popup))
        filter_actions.addWidget(self.create_button("Reducao BC", self.open_sale_reduction_popup))
        filter_actions.addWidget(self.create_button("Apuracao", self.open_sale_apuracao_popup))
        filter_actions.addWidget(self.create_button("Espelho Docs", self.open_sale_docs_popup))
        filter_actions.addStretch()
        layout.addLayout(filter_actions)

        hint = QLabel("Dica: clique duas vezes em uma linha para abrir o Espelho Docs.")
        hint.setObjectName("muted")
        layout.addWidget(hint)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(8)
        self.sale_metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate((("rows", "Linhas"), ("products", "Produtos"), ("operation", "Operacao"), ("base", "Base ICMS"), ("icms", "ICMS"))):
            card, value_label = self.create_metric_card_with_label(label, "0")
            self.sale_metric_labels[key] = value_label
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        self.sale_table = QTableWidget()
        self.sale_table.setColumnCount(13)
        self.sale_table.setHorizontalHeaderLabels(["Periodo", "Curva", "Codigo", "Descricao", "Cliente", "NCM", "CEST", "CFOP", "CST", "Aliq.", "Aliq. Efetiva", "Docs", "Lanc."])
        self.sale_table.setAlternatingRowColors(True)
        self.sale_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.sale_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.sale_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.sale_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.sale_table.verticalHeader().setVisible(False)
        self.sale_table.horizontalHeader().setStretchLastSection(False)
        self.sale_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        for column, width in enumerate((88, 72, 105, 330, 260, 90, 90, 85, 80, 85, 105, 65, 75)):
            self.sale_table.setColumnWidth(column, width)
        self.sale_table.itemDoubleClicked.connect(lambda _item: self.open_sale_docs_popup())
        layout.addWidget(self.sale_table, 1)
        self.bind_sale_live_filters()
        return page

    def build_contrib_page(self, operation_type: str) -> QWidget:
        is_entry = operation_type == "Entrada"
        prefix = "contrib_entry" if is_entry else "contrib_sale"
        caption = "Entradas" if is_entry else "Saidas"

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel(f"Consulta {caption} PIS/COFINS")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", lambda: self.clear_contrib_page(operation_type)))
        toolbar_layout.addWidget(self.create_button("Exportar", lambda: self.export_contrib_filter(operation_type)))
        toolbar_layout.addWidget(self.create_button("Processar", lambda: self.process_contrib_query(operation_type), primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        sped_input = QLineEdit()
        xml_input = QLineEdit()
        setattr(self, f"{prefix}_sped_input", sped_input)
        setattr(self, f"{prefix}_xml_input", xml_input)
        self.add_path_row(setup_layout, 0, "SPEDs Contrib.", sped_input, lambda: self.select_contrib_sped_files(operation_type), sped_input.clear)
        self.add_path_row(setup_layout, 1, "XMLs 55/65", xml_input, lambda: self.select_contrib_xml_sources(operation_type), xml_input.clear)
        layout.addWidget(setup)

        filters = self.create_panel()
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(12, 12, 12, 12)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(8)
        periods_widget = QWidget()
        periods_layout = QHBoxLayout(periods_widget)
        periods_layout.setContentsMargins(0, 0, 0, 0)
        periods_layout.setSpacing(8)
        setattr(self, f"{prefix}_periods_layout", periods_layout)
        cst_input = QLineEdit()
        cst_input.setPlaceholderText("01, 50, 70")
        cfop_input = QLineEdit()
        cfop_input.setPlaceholderText("1101, 5102")
        status_combo = QComboBox()
        status_combo.addItems(["Todos", "Ok", "Sem credito" if is_entry else "Sem debito", "Sem entrada" if is_entry else "Sem saida", "Com divergencia"])
        search_input = QLineEdit()
        search_input.setPlaceholderText("Codigo, descricao ou participante")
        setattr(self, f"{prefix}_cst_input", cst_input)
        setattr(self, f"{prefix}_cfop_input", cfop_input)
        setattr(self, f"{prefix}_status_combo", status_combo)
        setattr(self, f"{prefix}_search_input", search_input)
        self.add_filter_box(filters_layout, 0, 0, "Periodos", periods_widget)
        self.add_filter_box(filters_layout, 0, 1, "CST PIS/COFINS", cst_input)
        self.add_filter_box(filters_layout, 0, 2, "CFOP", cfop_input)
        search_box = QWidget()
        search_layout = QVBoxLayout(search_box)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)
        search_layout.addWidget(status_combo)
        search_layout.addWidget(search_input)
        self.add_filter_box(filters_layout, 0, 3, "Status / Busca", search_box)
        filters_layout.setColumnStretch(0, 2)
        filters_layout.setColumnStretch(1, 1)
        filters_layout.setColumnStretch(2, 1)
        filters_layout.setColumnStretch(3, 2)
        layout.addWidget(filters)

        filter_actions = QHBoxLayout()
        filter_actions.addWidget(self.create_button("Aplicar filtros", lambda: self.refresh_contrib_table(operation_type), primary=True))
        filter_actions.addWidget(self.create_button("Limpar filtros", lambda: self.clear_contrib_filters(operation_type)))
        filter_actions.addWidget(self.create_button("Exportar filtro atual", lambda: self.export_contrib_filter(operation_type)))
        filter_actions.addWidget(self.create_button(caption, lambda: self.open_contrib_operation_summary_popup(operation_type)))
        filter_actions.addWidget(self.create_button("Comp. Diagnostico", lambda: self.open_contrib_diagnostic_comparison_popup(operation_type)))
        filter_actions.addWidget(self.create_button("Diagnostico", lambda: self.open_contrib_diagnostic_popup(operation_type)))
        filter_actions.addWidget(self.create_button("Curva ABC", lambda: self.open_contrib_abc_popup(operation_type)))
        filter_actions.addWidget(self.create_button("Apuracao", lambda: self.open_contrib_apuracao_popup(operation_type)))
        filter_actions.addWidget(self.create_button("Espelho Docs", lambda: self.open_contrib_docs_popup(operation_type)))
        filter_actions.addStretch()
        layout.addLayout(filter_actions)

        hint = QLabel("Dica: clique duas vezes em uma linha para abrir o Espelho Docs.")
        hint.setObjectName("muted")
        layout.addWidget(hint)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(8)
        metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate(
            (
                ("rows", "Linhas"),
                ("products", "Produtos"),
                ("operation", "Operacao"),
                ("base_pis", "Base PIS"),
                ("pis", "PIS"),
                ("base_cofins", "Base COFINS"),
                ("cofins", "COFINS"),
            )
        ):
            card, value_label = self.create_metric_card_with_label(label, "0")
            metric_labels[key] = value_label
            metrics.addWidget(card, 0, column)
        setattr(self, f"{prefix}_metric_labels", metric_labels)
        layout.addLayout(metrics)

        table = QTableWidget()
        headers = [
            "Periodo",
            "Curva",
            "Codigo",
            "Descricao",
            "Participante",
            "NCM",
            "CEST",
            "CFOP",
            "CST PIS",
            "CST COFINS",
            "Aliq PIS",
            "Aliq COFINS",
            "Docs",
            "Lanc.",
        ]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setWordWrap(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        for column, width in enumerate((88, 72, 105, 330, 260, 90, 90, 85, 85, 105, 90, 105, 65, 75)):
            table.setColumnWidth(column, width)
        table.itemDoubleClicked.connect(lambda _item, op=operation_type: self.open_contrib_docs_popup(op))
        setattr(self, f"{prefix}_table", table)
        layout.addWidget(table, 1)
        self.bind_contrib_live_filters(operation_type)
        return page

    def build_xml_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("XML - Resumo CFOP e Creditos")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_xml_page))
        toolbar_layout.addWidget(self.create_button("Exportar Resumo", lambda: self.export_xml_summary("xlsx")))
        toolbar_layout.addWidget(self.create_button("Exportar Creditos", lambda: self.export_xml_credits("xlsx")))
        toolbar_layout.addWidget(self.create_button("Processar XMLs", self.process_xml_page, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.xml_source_input = QLineEdit()
        self.xml_company_input = QLineEdit()
        self.xml_company_input.setPlaceholderText("CNPJ/CPF da empresa para filtrar")
        self.xml_scope_combo = QComboBox()
        self.xml_scope_combo.addItems(["Todos", "Entrada", "Saida"])
        self.xml_scope_combo.currentTextChanged.connect(lambda _text: self.refresh_xml_summary_table())
        self.add_path_row(setup_layout, 0, "XMLs", self.xml_source_input, self.select_xml_summary_source, self.xml_source_input.clear)
        setup_layout.addWidget(QLabel("Empresa"), 1, 0)
        setup_layout.addWidget(self.xml_company_input, 1, 1)
        setup_layout.addWidget(QLabel("Operacao"), 1, 2)
        setup_layout.addWidget(self.xml_scope_combo, 1, 3)
        layout.addWidget(setup)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(8)
        self.xml_metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate((("xmls", "XMLs"), ("cfops", "CFOPs"), ("operation", "Operacao"), ("base", "Base ICMS"), ("icms", "ICMS"))):
            card, value_label = self.create_metric_card_with_label(label, "0")
            self.xml_metric_labels[key] = value_label
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        self.xml_summary_table = self.create_data_table(["Operacao", "CFOP", "Docs", "Itens", "Valor Operacao", "Desconto", "Base ICMS", "Valor ICMS", "Base ST", "Valor ST", "IPI", "Base PIS", "PIS", "Base COFINS", "COFINS"])
        for column, width in enumerate((90, 85, 70, 70, 120, 105, 110, 110, 105, 105, 95, 110, 95, 120, 105)):
            self.xml_summary_table.setColumnWidth(column, width)
        self.xml_summary_table.itemDoubleClicked.connect(lambda _item: self.open_xml_summary_details_popup())
        layout.addWidget(self.xml_summary_table, 1)

        credit_title = QLabel("Creditos XML de Entradas")
        credit_title.setObjectName("sectionTitle")
        layout.addWidget(credit_title)
        self.xml_credit_table = self.create_data_table(["Documento", "Serie", "Data", "Emitente", "CFOPs", "Itens", "Valor Operacao", "Base ICMS", "Credito ICMS", "Base ST", "Valor ST", "IPI", "Chave"])
        for column, width in enumerate((105, 70, 90, 260, 110, 70, 120, 110, 110, 105, 105, 95, 300)):
            self.xml_credit_table.setColumnWidth(column, width)
        self.xml_credit_table.itemDoubleClicked.connect(lambda _item: self.open_xml_credit_items_popup())
        layout.addWidget(self.xml_credit_table, 1)
        return page

    def build_compare_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("SPED x XML/Planilha")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_compare_page))
        toolbar_layout.addWidget(self.create_button("Exportar Excel", lambda: self.export_compare_rows("xlsx")))
        toolbar_layout.addWidget(self.create_button("Exportar CSV", lambda: self.export_compare_rows("csv")))
        self.compare_run_button = self.create_button("Comparar notas", self.process_compare_page, primary=True)
        toolbar_layout.addWidget(self.compare_run_button)
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.compare_sped_icms_input = QLineEdit()
        self.compare_sped_contrib_input = QLineEdit()
        self.compare_xml_input = QLineEdit()
        self.compare_sheet_input = QLineEdit()
        self.compare_mode_combo = QComboBox()
        self.compare_mode_combo.addItems(["SPED ICMS", "SPED Contribuicoes"])
        self.compare_scope_combo = QComboBox()
        self.compare_scope_combo.addItems(["Ambos", "Entrada", "Saida"])
        self.add_path_row(setup_layout, 0, "SPED ICMS", self.compare_sped_icms_input, lambda: self.select_single_file(self.compare_sped_icms_input, "Selecionar SPED ICMS", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)"), self.compare_sped_icms_input.clear)
        self.add_path_row(setup_layout, 1, "SPED Contrib.", self.compare_sped_contrib_input, lambda: self.select_single_file(self.compare_sped_contrib_input, "Selecionar SPED Contribuicoes", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)"), self.compare_sped_contrib_input.clear)
        self.add_path_row(setup_layout, 2, "Pasta XML", self.compare_xml_input, lambda: self.select_directory(self.compare_xml_input, "Selecionar pasta XML"), self.compare_xml_input.clear)
        self.add_path_row(setup_layout, 3, "Planilha", self.compare_sheet_input, lambda: self.select_single_file(self.compare_sheet_input, "Selecionar planilha", "Planilhas (*.xlsx *.xls *.xlsm);;Todos os arquivos (*.*)"), self.compare_sheet_input.clear)
        setup_layout.addWidget(QLabel("Comparar com"), 4, 0)
        setup_layout.addWidget(self.compare_mode_combo, 4, 1)
        setup_layout.addWidget(QLabel("Operacao"), 4, 2)
        setup_layout.addWidget(self.compare_scope_combo, 4, 3)
        layout.addWidget(setup)

        self.compare_status_label = QLabel("Selecione SPED + pasta XML ou planilha e clique em Comparar notas.")
        self.compare_status_label.setObjectName("muted")
        layout.addWidget(self.compare_status_label)

        self.compare_table = self.create_data_table(["Status", "Operacao", "Modelo", "Chave", "Numero", "Serie", "Data", "Emitente CNPJ", "Valor", "Arquivo"])
        for column, width in enumerate((130, 90, 80, 310, 95, 70, 90, 145, 110, 340)):
            self.compare_table.setColumnWidth(column, width)
        layout.addWidget(self.compare_table, 1)
        return page

    def build_runtime_rules_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Regras Dinamicas")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_runtime_rules))
        toolbar_layout.addWidget(self.create_button("Validar regras", self.validate_runtime_rules, primary=True))
        layout.addWidget(toolbar)

        editor_panel = self.create_panel()
        editor_layout = QVBoxLayout(editor_panel)
        editor_layout.setContentsMargins(12, 12, 12, 12)
        self.runtime_rules_text = QTextEdit()
        self.runtime_rules_text.setPlaceholderText("Uma regra por linha, no mesmo formato da tela antiga.")
        editor_layout.addWidget(self.runtime_rules_text, 1)
        layout.addWidget(editor_panel, 1)

        self.runtime_rules_status = QLabel("Informe as regras e clique em Validar regras.")
        self.runtime_rules_status.setObjectName("muted")
        layout.addWidget(self.runtime_rules_status)

        self.runtime_rules_table = self.create_data_table(["Linha", "Resumo"])
        self.runtime_rules_table.setColumnWidth(0, 70)
        self.runtime_rules_table.setColumnWidth(1, 760)
        layout.addWidget(self.runtime_rules_table, 1)
        return page

    def bind_entry_live_filters(self) -> None:
        self.cst_input.textChanged.connect(self.refresh_entry_table)
        self.cfop_input.textChanged.connect(self.refresh_entry_table)
        self.search_input.textChanged.connect(self.refresh_entry_table)
        self.status_combo.currentTextChanged.connect(self.refresh_entry_table)

    def bind_sale_live_filters(self) -> None:
        self.sale_cst_input.textChanged.connect(self.refresh_sale_table)
        self.sale_cfop_input.textChanged.connect(self.refresh_sale_table)
        self.sale_search_input.textChanged.connect(self.refresh_sale_table)
        self.sale_status_combo.currentTextChanged.connect(self.refresh_sale_table)

    def bind_contrib_live_filters(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        getattr(self, f"{prefix}_cst_input").textChanged.connect(lambda _text: self.refresh_contrib_table(operation_type))
        getattr(self, f"{prefix}_cfop_input").textChanged.connect(lambda _text: self.refresh_contrib_table(operation_type))
        getattr(self, f"{prefix}_search_input").textChanged.connect(lambda _text: self.refresh_contrib_table(operation_type))
        getattr(self, f"{prefix}_status_combo").currentTextChanged.connect(lambda _text: self.refresh_contrib_table(operation_type))

    def contrib_prefix(self, operation_type: str) -> str:
        return "contrib_entry" if operation_type == "Entrada" else "contrib_sale"

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

    def set_table_rows(self, table: QTableWidget, rows: list[list[object]]) -> None:
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)

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

    def select_sale_sped_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Fiscais", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        paths = append_unique_paths(parse_selected_paths(self.sale_sped_input.text()), files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        self.sale_sped_input.setText(format_selected_paths(paths))
        if limit_exceeded:
            QMessageBox.warning(self, "SPED Qt", "A consulta aceita no maximo 12 SPEDs.")

    def select_sale_xml_sources(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar XMLs 55/65", "", "Arquivos XML (*.xml);;Todos os arquivos (*.*)")
        if files:
            current_paths = parse_selected_paths(self.sale_xml_input.text())
            collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(files)
            if collapsed_to_folder:
                current_paths = [path for path in current_paths if not path.is_file()]
            self.sale_xml_input.setText(format_selected_paths(append_unique_paths(current_paths, collapsed_paths)))
            return
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta de XMLs 55/65")
        if directory:
            self.sale_xml_input.setText(format_selected_paths(append_unique_paths(parse_selected_paths(self.sale_xml_input.text()), [Path(directory)])))

    def select_contrib_sped_files(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        field: QLineEdit = getattr(self, f"{prefix}_sped_input")
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Contribuicoes", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        paths = append_unique_paths(parse_selected_paths(field.text()), files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        field.setText(format_selected_paths(paths))
        if limit_exceeded:
            QMessageBox.warning(self, "SPED Qt", "A consulta aceita no maximo 12 SPEDs.")

    def select_contrib_xml_sources(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        field: QLineEdit = getattr(self, f"{prefix}_xml_input")
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar XMLs 55/65", "", "Arquivos XML (*.xml);;Todos os arquivos (*.*)")
        if files:
            current_paths = parse_selected_paths(field.text())
            collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(files)
            if collapsed_to_folder:
                current_paths = [path for path in current_paths if not path.is_file()]
            field.setText(format_selected_paths(append_unique_paths(current_paths, collapsed_paths)))
            return
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta de XMLs 55/65")
        if directory:
            field.setText(format_selected_paths(append_unique_paths(parse_selected_paths(field.text()), [Path(directory)])))

    def select_xml_summary_source(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta de XMLs")
        if directory:
            self.xml_source_input.setText(str(Path(directory)))
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar XML", "", "Arquivos XML (*.xml);;Todos os arquivos (*.*)")
        if file_path:
            self.xml_source_input.setText(file_path)

    def select_single_file(self, field: QLineEdit, title: str, file_filter: str) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        if file_path:
            field.setText(file_path)

    def select_directory(self, field: QLineEdit, title: str) -> None:
        directory = QFileDialog.getExistingDirectory(self, title)
        if directory:
            field.setText(directory)

    def get_current_runtime_rules(self) -> list[dict[str, object]]:
        if not hasattr(self, "runtime_rules_text"):
            return []
        return parse_runtime_rule_lines(self.runtime_rules_text.toPlainText())

    def validate_runtime_rules(self) -> None:
        try:
            rules = self.get_current_runtime_rules()
            lines = [line.strip() for line in self.runtime_rules_text.toPlainText().splitlines() if line.strip()]
            rows = [[index + 1, runtime_rule_summary(line)] for index, line in enumerate(lines)]
            self.set_table_rows(self.runtime_rules_table, rows)
            self.runtime_rules_status.setText(f"Regras validas: {len(rules)}.")
            self.statusBar().showMessage(f"{len(rules)} regra(s) dinamica(s) validada(s).")
        except Exception as exc:
            self.runtime_rules_status.setText(f"Erro: {exc}")
            QMessageBox.critical(self, "Regras Dinamicas", str(exc))

    def clear_runtime_rules(self) -> None:
        self.runtime_rules_text.clear()
        self.runtime_rules_table.setRowCount(0)
        self.runtime_rules_status.setText("Informe as regras e clique em Validar regras.")

    def refresh_archives(self) -> None:
        if not hasattr(self, "archive_profile_table"):
            return
        try:
            self.mysql_repo.ensure_schema()
            profiles = self.mysql_repo.list_sped_profiles(self.environment)
            self.archive_profile_rows = {int(row["id"]): row for row in profiles}
            rows = [
                [
                    int(row["id"]),
                    row.get("empresa_nome_sped", "") or row.get("empresa_nome", "") or row.get("nome", ""),
                    row.get("empresa_cnpj_sped", "") or row.get("empresa_cnpj", ""),
                    self.format_period(row.get("periodo_inicio"), row.get("periodo_fim")),
                    row.get("total_arquivos", 0),
                ]
                for row in profiles
            ]
            self.archive_profile_table.blockSignals(True)
            self.set_table_rows(self.archive_profile_table, rows)
            self.archive_profile_table.blockSignals(False)
            if rows:
                self.archive_profile_table.selectRow(0)
                self.refresh_archive_files()
            else:
                self.archive_file_rows = {}
                self.archive_file_table.setRowCount(0)
            self.statusBar().showMessage(f"{len(profiles)} perfil(is) carregado(s).")
        except Exception as exc:
            self.statusBar().showMessage(f"Falha ao carregar perfis: {exc}")
            QMessageBox.critical(self, "SPEDs Arquivados", str(exc))

    def refresh_archive_files(self) -> None:
        if not hasattr(self, "archive_profile_table"):
            return
        selected = self.archive_profile_table.selectionModel().selectedRows()
        if not selected:
            self.archive_file_table.setRowCount(0)
            return
        profile_id_item = self.archive_profile_table.item(selected[0].row(), 0)
        if profile_id_item is None:
            return
        profile_id = int(profile_id_item.text())
        archives = self.mysql_repo.list_sped_archives(self.environment, profile_id)
        self.archive_file_rows = {int(row["id"]): row for row in archives}
        rows = [
            [
                int(row["id"]),
                row.get("tipo_sped", ""),
                row.get("arquivo_nome_original", ""),
                self.format_period(row.get("periodo_inicio"), row.get("periodo_fim")),
                row.get("total_produtos", 0),
                row.get("total_documentos", 0),
                row.get("total_itens", 0),
                row.get("total_c190", 0),
                str(row.get("arquivo_hash_sha256", ""))[:12],
            ]
            for row in archives
        ]
        self.set_table_rows(self.archive_file_table, rows)

    def import_speds(self) -> None:
        selected_files, _ = QFileDialog.getOpenFileNames(self, "Importar SPEDs", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not selected_files:
            return
        imported = 0
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            for selected in selected_files:
                if self.register_sped(Path(selected)):
                    imported += 1
            self.refresh_archives()
            self.statusBar().showMessage(f"{imported} de {len(selected_files)} arquivo(s) importado(s).")
        except Exception as exc:
            QMessageBox.critical(self, "Importar SPED", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def register_sped(self, sped_path: Path) -> int | None:
        metadata = archive_original_sped_file(sped_path, self.project_root_dir / "storage" / "original_speds", self.environment)
        self.mysql_repo.ensure_schema()
        existing = self.mysql_repo.get_sped_archive_by_hash(self.environment, metadata.file_hash_sha256)
        if existing:
            return int(existing["id"])
        profile_id = self.mysql_repo.ensure_sped_profile(
            self.environment,
            metadata.default_profile_name,
            "Perfil criado pela interface Qt.",
            metadata.company_name,
            metadata.company_tax_id,
        )
        archive_id = self.mysql_repo.save_sped_archive(
            {
                "perfil_id": profile_id,
                "ambiente": self.environment,
                "tipo_sped": metadata.sped_type,
                "periodo_inicio": metadata.period_start,
                "periodo_fim": metadata.period_end,
                "empresa_nome_sped": metadata.company_name,
                "empresa_cnpj": metadata.company_tax_id,
                "arquivo_nome_original": metadata.file_name,
                "arquivo_hash_sha256": metadata.file_hash_sha256,
                "arquivo_tamanho": metadata.file_size,
                "caminho_arquivo_original": str(metadata.source_path),
                "caminho_arquivo_arquivado": str(metadata.archived_path),
                "observacao": f"Importado pela interface Qt. Empresa no SPED: {metadata.company_name}",
            }
        )
        if metadata.sped_type == "fiscal":
            products, _sales_rows, detailed_items, c190_rows, _c190_product_rows = read_sped_file(metadata.source_path)
            self.mysql_repo.replace_sped_extracted_data(archive_id, products, detailed_items, c190_rows)
        return archive_id

    def open_selected_archive_folder(self) -> None:
        selected = self.archive_file_table.selectionModel().selectedRows() if hasattr(self, "archive_file_table") else []
        if not selected:
            QMessageBox.warning(self, "SPEDs Arquivados", "Selecione um arquivo arquivado.")
            return
        archive_id_item = self.archive_file_table.item(selected[0].row(), 0)
        if archive_id_item is None:
            return
        row = self.archive_file_rows.get(int(archive_id_item.text()))
        if not row:
            return
        path = Path(str(row.get("caminho_arquivo_arquivado", "")))
        if not path.exists():
            path = Path(str(row.get("caminho_arquivo_original", "")))
        if path.exists():
            os.startfile(path.parent)

    def create_schema(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            self.statusBar().showMessage("Banco e tabelas atualizados.")
            QMessageBox.information(self, "Configuracoes", "Banco e tabelas atualizados.")
        except Exception as exc:
            QMessageBox.critical(self, "Configuracoes", str(exc))

    def test_mysql_connection(self) -> None:
        try:
            self.mysql_repo.test_connection()
            self.statusBar().showMessage("Conexao MySQL OK.")
            QMessageBox.information(self, "Configuracoes", "Conexao MySQL OK.")
        except Exception as exc:
            QMessageBox.critical(self, "Configuracoes", str(exc))

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

    def process_sale_query(self) -> None:
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            sped_paths = parse_selected_paths(self.sale_sped_input.text())
            xml_sources = parse_selected_paths(self.sale_xml_input.text())
            period_labels, rows = build_sale_period_comparison_rows(sped_paths, xml_sources)
            self.sale_rows = list(rows)
            self.rebuild_sale_period_checks(list(period_labels))
            self.refresh_sale_table()
            self.statusBar().showMessage(f"Saidas processadas: {len(rows)} linha(s), {len(period_labels)} periodo(s). {dt.datetime.now().strftime('%H:%M:%S')}")
        except Exception as exc:
            QMessageBox.critical(self, "SPED Qt", str(exc))
            self.statusBar().showMessage(f"Falha no processamento: {exc}")
        finally:
            QApplication.restoreOverrideCursor()

    def process_contrib_query(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            sped_paths = parse_selected_paths(getattr(self, f"{prefix}_sped_input").text())
            xml_sources = parse_selected_paths(getattr(self, f"{prefix}_xml_input").text())
            period_labels, rows = build_pis_cofins_period_comparison_rows(sped_paths, operation_type, xml_sources=xml_sources)
            if operation_type == "Entrada":
                self.contrib_entry_rows = list(rows)
            else:
                self.contrib_sale_rows = list(rows)
            self.rebuild_contrib_period_checks(operation_type, list(period_labels))
            self.refresh_contrib_table(operation_type)
            self.statusBar().showMessage(f"{operation_type}s PIS/COFINS processadas: {len(rows)} linha(s), {len(period_labels)} periodo(s). {dt.datetime.now().strftime('%H:%M:%S')}")
        except Exception as exc:
            QMessageBox.critical(self, "SPED Qt", str(exc))
            self.statusBar().showMessage(f"Falha no processamento PIS/COFINS: {exc}")
        finally:
            QApplication.restoreOverrideCursor()

    def process_xml_page(self) -> None:
        source_text = self.xml_source_input.text().strip()
        if not source_text:
            QMessageBox.warning(self, "XML", "Selecione uma pasta ou arquivo XML.")
            return
        source_path = Path(source_text)
        if not source_path.exists():
            QMessageBox.warning(self, "XML", f"Caminho nao encontrado: {source_path}")
            return
        company_tax_id = self.xml_company_input.text().strip()
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            rows, stats = build_xml_cfop_summary_rows(source_path, company_tax_id, self.xml_scope_combo.currentText())
            invoice_rows, cfop_rows, credit_stats = build_xml_entry_credit_rows(source_path, company_tax_id)
            self.xml_summary_all_rows = list(rows)
            self.xml_summary_stats = dict(stats)
            self.xml_credit_invoice_rows = list(invoice_rows)
            self.xml_credit_cfop_rows = list(cfop_rows)
            self.refresh_xml_summary_table()
            self.refresh_xml_credit_table()
            self.statusBar().showMessage(
                f"XMLs processados: {stats.get('xml_total', 0)} total, {stats.get('xml_ignored', 0)} invalido(s). Creditos: {credit_stats.get('xml_total', stats.get('xml_total', 0))} XML(s)."
            )
        except Exception as exc:
            QMessageBox.critical(self, "XML", str(exc))
            self.statusBar().showMessage(f"Falha no processamento XML: {exc}")
        finally:
            QApplication.restoreOverrideCursor()

    def process_compare_page(self) -> None:
        compare_mode = "icms" if self.compare_mode_combo.currentText() == "SPED ICMS" else "contrib"
        sped_text = self.compare_sped_icms_input.text().strip() if compare_mode == "icms" else self.compare_sped_contrib_input.text().strip()
        xml_text = self.compare_xml_input.text().strip()
        sheet_text = self.compare_sheet_input.text().strip()
        if not sped_text or not Path(sped_text).exists():
            QMessageBox.warning(self, "SPED x XML", "Selecione um arquivo SPED valido para o modo escolhido.")
            return
        has_xml = bool(xml_text and Path(xml_text).exists())
        has_sheet = bool(sheet_text and Path(sheet_text).exists())
        if has_xml == has_sheet:
            QMessageBox.warning(self, "SPED x XML", "Selecione apenas uma origem: pasta XML ou planilha.")
            return
        operation_scope = normalize_compare_operation_scope(self.compare_scope_combo.currentText())
        source_kind = "sheet" if has_sheet else "xml"
        source_path = Path(sheet_text) if has_sheet else Path(xml_text)
        self.compare_rows = []
        self.compare_stats = {}
        self.compare_table.setRowCount(0)
        self.compare_status_label.setText("Comparando arquivos... aguarde.")
        self.statusBar().showMessage("Comparando SPED x XML/Planilha...")
        self.compare_run_button.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)

        self.compare_thread = QThread(self)
        self.compare_worker = CompareWorker(Path(sped_text), source_path, compare_mode, source_kind, operation_scope)
        self.compare_worker.moveToThread(self.compare_thread)
        self.compare_thread.started.connect(self.compare_worker.run)
        self.compare_worker.progress.connect(self.update_compare_status)
        self.compare_worker.finished.connect(self.render_compare_results)
        self.compare_worker.failed.connect(self.handle_compare_error)
        self.compare_worker.finished.connect(self.compare_thread.quit)
        self.compare_worker.failed.connect(self.compare_thread.quit)
        self.compare_thread.finished.connect(self.compare_worker.deleteLater)
        self.compare_thread.finished.connect(self.compare_thread.deleteLater)
        self.compare_thread.finished.connect(lambda: setattr(self, "compare_worker", None))
        self.compare_thread.finished.connect(lambda: setattr(self, "compare_thread", None))
        self.compare_thread.start()

    def update_compare_status(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.compare_status_label.setText(f"{message} ({percent}%)")

    def handle_compare_error(self, message: str) -> None:
        QApplication.restoreOverrideCursor()
        self.compare_run_button.setEnabled(True)
        QMessageBox.critical(self, "SPED x XML", message)
        self.compare_status_label.setText(f"Falha na comparacao: {message}")
        self.statusBar().showMessage(f"Falha na comparacao: {message}")

    def render_compare_results(
        self,
        present: list[object],
        missing: list[object],
        sped_missing_xml: list[object],
        stats: dict[str, object],
        compare_mode: str,
        source_kind: str,
        operation_scope: str,
    ) -> None:
        groups = [("Lancada", present), ("Nao lancada", missing)]
        if source_kind == "xml":
            groups.append(("Lancado SPED / Sem XML", sped_missing_xml))
            cancelled = stats.get("cancelled_xml_invoices", [])
            if isinstance(cancelled, list):
                groups.append(("XML cancelado", cancelled))
        operation_by_key = stats.get("operation_by_key", {})
        source_operation_by_key = stats.get("source_operation_by_key", {})
        if not isinstance(operation_by_key, dict):
            operation_by_key = {}
        if not isinstance(source_operation_by_key, dict):
            source_operation_by_key = {}
        selected_operation = normalize_compare_operation_scope(operation_scope)
        rows: list[list[object]] = []
        for status, invoices in groups:
            for invoice in invoices:
                key = getattr(invoice, "key", "")
                operation = str(operation_by_key.get(key, "") or source_operation_by_key.get(key, "") or "")
                if not operation and selected_operation != "Ambos":
                    operation = selected_operation
                rows.append(
                    [
                        status,
                        operation,
                        getattr(invoice, "model", ""),
                        key,
                        getattr(invoice, "number", ""),
                        getattr(invoice, "series", ""),
                        getattr(invoice, "issue_date", ""),
                        getattr(invoice, "issuer_cnpj", ""),
                        getattr(invoice, "total_value", ""),
                        getattr(invoice, "file_path", ""),
                    ]
                )
        self.compare_rows = rows
        self.compare_stats = dict(stats)
        self.set_table_rows(self.compare_table, rows)
        self.compare_status_label.setText(
            f"Docs SPED: {stats.get('sped_keys', 0)} | Encontrados: {stats.get('present_total', 0)} | Nao lancados/nao encontrados: {stats.get('missing_total', 0)} | SPED sem XML: {stats.get('sped_missing_xml_total', 0)}"
        )
        self.statusBar().showMessage("Comparacao SPED x XML/Planilha concluida.")
        self.compare_run_button.setEnabled(True)
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

    def rebuild_sale_period_checks(self, labels: list[str]) -> None:
        while self.sale_periods_layout.count():
            item = self.sale_periods_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.sale_period_checks = {}
        if not labels:
            empty = QLabel("Nenhum periodo")
            empty.setObjectName("muted")
            self.sale_periods_layout.addWidget(empty)
            return
        for label in labels:
            check = QCheckBox(label)
            check.setChecked(True)
            check.stateChanged.connect(self.refresh_sale_table)
            self.sale_period_checks[label] = check
            self.sale_periods_layout.addWidget(check)
        self.sale_periods_layout.addStretch()

    def rebuild_contrib_period_checks(self, operation_type: str, labels: list[str]) -> None:
        prefix = self.contrib_prefix(operation_type)
        periods_layout: QHBoxLayout = getattr(self, f"{prefix}_periods_layout")
        while periods_layout.count():
            item = periods_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        checks: dict[str, QCheckBox] = {}
        if not labels:
            empty = QLabel("Nenhum periodo")
            empty.setObjectName("muted")
            periods_layout.addWidget(empty)
        else:
            for label in labels:
                check = QCheckBox(label)
                check.setChecked(True)
                check.stateChanged.connect(lambda _state, op=operation_type: self.refresh_contrib_table(op))
                checks[label] = check
                periods_layout.addWidget(check)
            periods_layout.addStretch()
        if operation_type == "Entrada":
            self.contrib_entry_period_checks = checks
        else:
            self.contrib_sale_period_checks = checks

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

    def selected_sale_periods(self) -> set[str]:
        selected = {label for label, check in self.sale_period_checks.items() if check.isChecked()}
        return selected or set(self.sale_period_checks)

    def selected_contrib_periods(self, operation_type: str) -> set[str]:
        checks = self.contrib_entry_period_checks if operation_type == "Entrada" else self.contrib_sale_period_checks
        selected = {label for label, check in checks.items() if check.isChecked()}
        return selected or set(checks)

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

    def filtered_sale_table_rows(self) -> list[dict[str, object]]:
        selected_periods = self.selected_sale_periods()
        cst_filters = self.parse_filter_tokens(self.sale_cst_input.text())
        cfop_filters = self.parse_filter_tokens(self.sale_cfop_input.text())
        status_filter = self.sale_status_combo.currentText().strip()
        search_text = self.normalize_search_text(self.sale_search_input.text())
        filtered: list[dict[str, object]] = []
        for row in self.sale_rows:
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
            if status_filter == "Sem debito" and "Sem debito" not in status_value:
                continue
            if status_filter == "Sem saida" and status_value != "Sem saida":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem saida"} and "Sem debito" not in status_value:
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

    def filtered_contrib_table_rows(self, operation_type: str) -> list[dict[str, object]]:
        prefix = self.contrib_prefix(operation_type)
        source_rows = self.contrib_entry_rows if operation_type == "Entrada" else self.contrib_sale_rows
        selected_periods = self.selected_contrib_periods(operation_type)
        cst_filters = self.parse_filter_tokens(getattr(self, f"{prefix}_cst_input").text())
        cfop_filters = self.parse_filter_tokens(getattr(self, f"{prefix}_cfop_input").text())
        status_filter = getattr(self, f"{prefix}_status_combo").currentText().strip()
        search_text = self.normalize_search_text(getattr(self, f"{prefix}_search_input").text())
        filtered: list[dict[str, object]] = []
        for row in source_rows:
            period = str(row.get("period", ""))
            if selected_periods and period not in selected_periods:
                continue
            if cst_filters:
                row_csts = {part.strip() for part in f"{row.get('cst_pis', '')}|{row.get('cst_cofins', '')}".split("|") if part.strip()}
                if not row_csts.intersection(cst_filters):
                    continue
            if cfop_filters:
                row_cfops = {part.strip() for part in str(row.get("cfop", "")).split("|") if part.strip()}
                if not row_cfops.intersection(cfop_filters):
                    continue
            status_value = str(row.get("status", "")).strip()
            if status_filter == "Ok" and status_value != "Ok":
                continue
            if status_filter == "Sem credito" and "Sem credito" not in status_value:
                continue
            if status_filter == "Sem entrada" and status_value != "Sem entrada":
                continue
            if status_filter == "Sem debito" and "Sem debito" not in status_value:
                continue
            if status_filter == "Sem saida" and status_value != "Sem saida":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem entrada", "Sem saida"} and "Sem credito" not in status_value and "Sem debito" not in status_value:
                continue
            if search_text:
                searchable = self.normalize_search_text(f"{row.get('code', '')} {row.get('description', '')} {row.get('cest', '')} {row.get('suppliers', '')}")
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

    def refresh_sale_table(self) -> None:
        self.filtered_sale_rows = self.filtered_sale_table_rows()
        self.sale_table.setRowCount(len(self.filtered_sale_rows))
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        products: set[str] = set()
        for row_index, row in enumerate(self.filtered_sale_rows):
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
                self.sale_table.setItem(row_index, column, item)
        self.sale_metric_labels["rows"].setText(str(len(self.filtered_sale_rows)))
        self.sale_metric_labels["products"].setText(str(len(products)))
        self.sale_metric_labels["operation"].setText(self.format_number(total_operation))
        self.sale_metric_labels["base"].setText(self.format_number(total_base))
        self.sale_metric_labels["icms"].setText(self.format_number(total_icms))

    def refresh_contrib_table(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        filtered_rows = self.filtered_contrib_table_rows(operation_type)
        if operation_type == "Entrada":
            self.filtered_contrib_entry_rows = filtered_rows
        else:
            self.filtered_contrib_sale_rows = filtered_rows
        table: QTableWidget = getattr(self, f"{prefix}_table")
        table.setRowCount(len(filtered_rows))
        total_operation = Decimal("0")
        total_base_pis = Decimal("0")
        total_pis = Decimal("0")
        total_base_cofins = Decimal("0")
        total_cofins = Decimal("0")
        products: set[str] = set()
        for row_index, row in enumerate(filtered_rows):
            total_operation += self.decimal_value(row.get("sale_value"))
            total_base_pis += self.decimal_value(row.get("base_pis"))
            total_pis += self.decimal_value(row.get("pis_value"))
            total_base_cofins += self.decimal_value(row.get("base_cofins"))
            total_cofins += self.decimal_value(row.get("cofins_value"))
            code = str(row.get("code", "")).strip()
            if code:
                products.add(code)
            values = [
                row.get("period", ""),
                row.get("curve_abc", "C"),
                code,
                row.get("description", ""),
                self.format_supplier_text(row),
                row.get("ncm", ""),
                row.get("cest", ""),
                row.get("cfop", ""),
                row.get("cst_pis", ""),
                row.get("cst_cofins", ""),
                self.format_number(row.get("aliquota_pis")),
                self.format_number(row.get("aliquota_cofins")),
                row.get("document_count", 0),
                row.get("launch_count", 0),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if column in {0, 1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 13}:
                    item.setTextAlignment(Qt.AlignCenter)
                if str(row.get("status", "")).strip() not in {"", "Ok"}:
                    item.setBackground(QColor("#fff8e8"))
                table.setItem(row_index, column, item)
        metric_labels: dict[str, QLabel] = getattr(self, f"{prefix}_metric_labels")
        metric_labels["rows"].setText(str(len(filtered_rows)))
        metric_labels["products"].setText(str(len(products)))
        metric_labels["operation"].setText(self.format_number(total_operation))
        metric_labels["base_pis"].setText(self.format_number(total_base_pis))
        metric_labels["pis"].setText(self.format_number(total_pis))
        metric_labels["base_cofins"].setText(self.format_number(total_base_cofins))
        metric_labels["cofins"].setText(self.format_number(total_cofins))

    def refresh_xml_summary_table(self) -> None:
        if not hasattr(self, "xml_summary_table"):
            return
        rows = filter_xml_summary_rows_by_scope(self.xml_summary_all_rows, self.xml_scope_combo.currentText())
        self.xml_summary_rows = list(rows)
        table_rows: list[list[object]] = []
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        document_keys: set[str] = set()
        for row in rows:
            total_operation += self.decimal_value(row.get("operation_value"))
            total_base += self.decimal_value(row.get("base_icms"))
            total_icms += self.decimal_value(row.get("icms_value"))
            for detail in row.get("details", []):
                if isinstance(detail, dict) and str(detail.get("key", "")).strip():
                    document_keys.add(str(detail.get("key", "")).strip())
            table_rows.append(
                [
                    row.get("operation_type", ""),
                    row.get("cfop", ""),
                    row.get("document_count", 0),
                    row.get("items", 0),
                    self.format_number(row.get("operation_value")),
                    self.format_number(row.get("discount_value")),
                    self.format_number(row.get("base_icms")),
                    self.format_number(row.get("icms_value")),
                    self.format_number(row.get("base_icms_st")),
                    self.format_number(row.get("icms_st_value")),
                    self.format_number(row.get("ipi_value")),
                    self.format_number(row.get("base_pis")),
                    self.format_number(row.get("pis_value")),
                    self.format_number(row.get("base_cofins")),
                    self.format_number(row.get("cofins_value")),
                ]
            )
        self.set_table_rows(self.xml_summary_table, table_rows)
        self.xml_metric_labels["xmls"].setText(str(self.xml_summary_stats.get("xml_total", len(document_keys))))
        self.xml_metric_labels["cfops"].setText(str(len(rows)))
        self.xml_metric_labels["operation"].setText(self.format_number(total_operation))
        self.xml_metric_labels["base"].setText(self.format_number(total_base))
        self.xml_metric_labels["icms"].setText(self.format_number(total_icms))

    def refresh_xml_credit_table(self) -> None:
        if not hasattr(self, "xml_credit_table"):
            return
        rows = [
            [
                row.get("number", ""),
                row.get("series", ""),
                row.get("issue_date", ""),
                row.get("issuer_name", ""),
                row.get("cfops", ""),
                row.get("items", 0),
                self.format_number(row.get("operation_value")),
                self.format_number(row.get("base_icms")),
                self.format_number(row.get("icms_value")),
                self.format_number(row.get("base_icms_st")),
                self.format_number(row.get("icms_st_value")),
                self.format_number(row.get("ipi_value")),
                row.get("key", ""),
            ]
            for row in self.xml_credit_invoice_rows
        ]
        self.set_table_rows(self.xml_credit_table, rows)

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

    def clear_sale_filters(self) -> None:
        self.sale_cst_input.clear()
        self.sale_cfop_input.clear()
        self.sale_status_combo.setCurrentText("Todos")
        self.sale_search_input.clear()
        for check in self.sale_period_checks.values():
            check.setChecked(True)
        self.refresh_sale_table()

    def clear_contrib_filters(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        getattr(self, f"{prefix}_cst_input").clear()
        getattr(self, f"{prefix}_cfop_input").clear()
        getattr(self, f"{prefix}_status_combo").setCurrentText("Todos")
        getattr(self, f"{prefix}_search_input").clear()
        checks = self.contrib_entry_period_checks if operation_type == "Entrada" else self.contrib_sale_period_checks
        for check in checks.values():
            check.setChecked(True)
        self.refresh_contrib_table(operation_type)

    def clear_entry_page(self) -> None:
        self.sped_input.clear()
        self.xml_input.clear()
        self.entry_rows = []
        self.filtered_entry_rows = []
        self.rebuild_period_checks([])
        self.refresh_entry_table()
        self.statusBar().showMessage("Tela limpa.")

    def clear_sale_page(self) -> None:
        self.sale_sped_input.clear()
        self.sale_xml_input.clear()
        self.sale_rows = []
        self.filtered_sale_rows = []
        self.rebuild_sale_period_checks([])
        self.refresh_sale_table()
        self.statusBar().showMessage("Tela limpa.")

    def clear_contrib_page(self, operation_type: str) -> None:
        prefix = self.contrib_prefix(operation_type)
        getattr(self, f"{prefix}_sped_input").clear()
        getattr(self, f"{prefix}_xml_input").clear()
        if operation_type == "Entrada":
            self.contrib_entry_rows = []
            self.filtered_contrib_entry_rows = []
        else:
            self.contrib_sale_rows = []
            self.filtered_contrib_sale_rows = []
        self.rebuild_contrib_period_checks(operation_type, [])
        self.refresh_contrib_table(operation_type)
        self.statusBar().showMessage("Tela limpa.")

    def clear_xml_page(self) -> None:
        self.xml_source_input.clear()
        self.xml_company_input.clear()
        self.xml_scope_combo.setCurrentText("Todos")
        self.xml_summary_all_rows = []
        self.xml_summary_rows = []
        self.xml_summary_stats = {}
        self.xml_credit_invoice_rows = []
        self.xml_credit_cfop_rows = []
        self.refresh_xml_summary_table()
        self.refresh_xml_credit_table()
        self.statusBar().showMessage("Tela XML limpa.")

    def clear_compare_page(self) -> None:
        self.compare_sped_icms_input.clear()
        self.compare_sped_contrib_input.clear()
        self.compare_xml_input.clear()
        self.compare_sheet_input.clear()
        self.compare_rows = []
        self.compare_stats = {}
        self.compare_table.setRowCount(0)
        self.compare_status_label.setText("Selecione SPED + pasta XML ou planilha e clique em Comparar notas.")
        self.statusBar().showMessage("Tela SPED x XML limpa.")

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

    def export_sale_filter(self) -> None:
        self.export_consultation_rows(self.filtered_sale_rows, "consulta_saidas_filtrada.xlsx", "Consulta Saidas")

    def export_contrib_filter(self, operation_type: str) -> None:
        rows = self.filtered_contrib_entry_rows if operation_type == "Entrada" else self.filtered_contrib_sale_rows
        default_name = "consulta_entradas_pis_cofins.xlsx" if operation_type == "Entrada" else "consulta_saidas_pis_cofins.xlsx"
        if not rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar consulta filtrada",
            default_name,
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        headers = [
            "Periodo",
            "Arquivo SPED",
            "Curva ABC",
            "Codigo",
            "Descricao",
            "Participante",
            "NCM",
            "CEST",
            "CFOP",
            "CST PIS",
            "CST COFINS",
            "Aliq PIS",
            "Aliq COFINS",
            "Documentos",
            "Lancamentos",
            "Quantidade",
            "Valor Operacao",
            "Base PIS",
            "Base COFINS",
            "Valor PIS",
            "Valor COFINS",
            "Status",
        ]
        export_rows = [
            [
                row.get("period", ""),
                row.get("file_name", ""),
                row.get("curve_abc", "C"),
                row.get("code", ""),
                row.get("description", ""),
                row.get("suppliers", ""),
                row.get("ncm", ""),
                row.get("cest", ""),
                row.get("cfop", ""),
                row.get("cst_pis", ""),
                row.get("cst_cofins", ""),
                self.decimal_value(row.get("aliquota_pis")),
                self.decimal_value(row.get("aliquota_cofins")),
                self.integer_from_display_value(row.get("document_count", 0)),
                self.integer_from_display_value(row.get("launch_count", 0)),
                self.decimal_value(row.get("quantity")),
                self.decimal_value(row.get("sale_value")),
                self.decimal_value(row.get("base_pis")),
                self.decimal_value(row.get("base_cofins")),
                self.decimal_value(row.get("pis_value")),
                self.decimal_value(row.get("cofins_value")),
                row.get("status", ""),
            ]
            for row in rows
        ]
        total_row = self.build_popup_total_row(headers, export_rows)
        rows_to_write = export_rows + ([total_row] if total_row else [])
        output_path = Path(output)
        if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
            if output_path.suffix.lower() != ".csv":
                output_path = output_path.with_suffix(".csv")
            write_simple_csv_file(output_path, headers, rows_to_write)
        else:
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            write_simple_excel_workbook(output_path, [(f"Consulta {operation_type}"[:31], headers, rows_to_write, {"include_total": False})])
        self.statusBar().showMessage(f"Consulta PIS/COFINS exportada: {output_path}")

    def export_xml_summary(self, output_type: str) -> None:
        if not self.xml_summary_rows:
            QMessageBox.warning(self, "Exportar XML", "Nao ha resumo XML para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        output, _ = QFileDialog.getSaveFileName(self, "Exportar resumo XML", f"resumo_xml_cfop{suffix}", "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)")
        if not output:
            return
        headers = ["Operacao", "CFOP", "Docs", "Itens", "Valor Operacao", "Desconto", "Base ICMS", "Valor ICMS", "Base ST", "Valor ST", "IPI", "Base PIS", "PIS", "Base COFINS", "COFINS"]
        rows = [
            [
                row.get("operation_type", ""),
                row.get("cfop", ""),
                row.get("document_count", 0),
                row.get("items", 0),
                self.decimal_value(row.get("operation_value")),
                self.decimal_value(row.get("discount_value")),
                self.decimal_value(row.get("base_icms")),
                self.decimal_value(row.get("icms_value")),
                self.decimal_value(row.get("base_icms_st")),
                self.decimal_value(row.get("icms_st_value")),
                self.decimal_value(row.get("ipi_value")),
                self.decimal_value(row.get("base_pis")),
                self.decimal_value(row.get("pis_value")),
                self.decimal_value(row.get("base_cofins")),
                self.decimal_value(row.get("cofins_value")),
            ]
            for row in self.xml_summary_rows
        ]
        total_row = self.build_popup_total_row(headers, rows)
        output_path = Path(output)
        rows_to_write = rows + ([total_row] if total_row else [])
        if output_path.suffix.lower() == ".csv":
            write_simple_csv_file(output_path, headers, rows_to_write)
        else:
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            write_simple_excel_workbook(output_path, [("Resumo XML", headers, rows_to_write, {"include_total": False})])
        self.statusBar().showMessage(f"Resumo XML exportado: {output_path}")

    def export_xml_credits(self, output_type: str) -> None:
        if not self.xml_credit_invoice_rows:
            QMessageBox.warning(self, "Exportar XML", "Nao ha creditos XML para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        output, _ = QFileDialog.getSaveFileName(self, "Exportar creditos XML", f"creditos_xml_entradas{suffix}", "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)")
        if not output:
            return
        headers = ["Documento", "Serie", "Data", "Emitente", "CFOPs", "Itens", "Valor Operacao", "Base ICMS", "Credito ICMS", "Base ST", "Valor ST", "IPI", "Chave"]
        rows = [
            [
                row.get("number", ""),
                row.get("series", ""),
                row.get("issue_date", ""),
                row.get("issuer_name", ""),
                row.get("cfops", ""),
                row.get("items", 0),
                self.decimal_value(row.get("operation_value")),
                self.decimal_value(row.get("base_icms")),
                self.decimal_value(row.get("icms_value")),
                self.decimal_value(row.get("base_icms_st")),
                self.decimal_value(row.get("icms_st_value")),
                self.decimal_value(row.get("ipi_value")),
                row.get("key", ""),
            ]
            for row in self.xml_credit_invoice_rows
        ]
        total_row = self.build_popup_total_row(headers, rows)
        output_path = Path(output)
        rows_to_write = rows + ([total_row] if total_row else [])
        if output_path.suffix.lower() == ".csv":
            write_simple_csv_file(output_path, headers, rows_to_write)
        else:
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            write_simple_excel_workbook(output_path, [("Creditos XML", headers, rows_to_write, {"include_total": False})])
        self.statusBar().showMessage(f"Creditos XML exportados: {output_path}")

    def export_compare_rows(self, output_type: str) -> None:
        if not self.compare_rows:
            QMessageBox.warning(self, "Exportar comparacao", "Nao ha dados da comparacao para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        output, _ = QFileDialog.getSaveFileName(self, "Exportar comparacao", f"comparacao_sped_xml{suffix}", "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)")
        if not output:
            return
        headers = ["Status", "Operacao", "Modelo", "Chave", "Numero", "Serie", "Data", "Emitente CNPJ", "Valor", "Arquivo"]
        rows = [
            [*row[:8], self.decimal_from_display_value(row[8]), row[9]]
            for row in self.compare_rows
        ]
        output_path = Path(output)
        if output_path.suffix.lower() == ".csv":
            write_simple_csv_file(output_path, headers, rows)
        else:
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            write_simple_excel_workbook(output_path, [("Comparacao", headers, rows, {"include_total": False})])
        self.statusBar().showMessage(f"Comparacao exportada: {output_path}")

    def export_consultation_rows(self, source_rows: list[dict[str, object]], default_name: str, sheet_name: str) -> None:
        if not source_rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar consulta filtrada",
            default_name,
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
            for row in source_rows
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
            write_simple_excel_workbook(output_path, [(sheet_name, headers, rows_to_write, {"include_total": False})])
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
        return self.get_launch_details_from_rows(self.filtered_entry_rows, "Entrada")

    def get_filtered_sale_launch_details(self) -> list[dict[str, object]]:
        return self.get_launch_details_from_rows(self.filtered_sale_rows, "Saida")

    def get_launch_details_from_rows(self, rows: list[dict[str, object]], operation_type: str) -> list[dict[str, object]]:
        details: list[dict[str, object]] = []
        for row in rows:
            row_period = str(self.first_row_value(row, "period", "period_label"))
            launch_details = row.get("launch_details")
            if isinstance(launch_details, list) and launch_details:
                for detail in launch_details:
                    if not isinstance(detail, dict):
                        continue
                    enriched = dict(detail)
                    enriched.setdefault("period", row_period)
                    enriched.setdefault("operation_type", operation_type)
                    enriched.setdefault("source_register", "C170")
                    details.append(enriched)
                continue
            fallback = dict(row)
            fallback.setdefault("period", row_period)
            fallback.setdefault("operation_type", operation_type)
            fallback.setdefault("source_register", "C170")
            fallback.setdefault("total_operation_value", self.first_row_value(row, "sale_value", "total_operation_value"))
            details.append(fallback)
        return details

    def get_filtered_summary_rows_from_details(self) -> list[dict[str, object]]:
        return self.get_summary_rows_from_details(self.get_filtered_launch_details(), "Entrada")

    def get_filtered_sale_summary_rows_from_details(self) -> list[dict[str, object]]:
        return self.get_summary_rows_from_details(self.get_filtered_sale_launch_details(), "Saida")

    def get_summary_rows_from_details(self, details: list[dict[str, object]], operation_type: str) -> list[dict[str, object]]:
        summary_rows: list[dict[str, object]] = []
        for detail in details:
            sale_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(detail.get("base_icms"))
            icms_value = self.decimal_value(detail.get("icms_value"))
            summary_rows.append(
                {
                    "operation_type": operation_type,
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

    def open_sale_operation_summary_popup(self) -> None:
        self.open_operation_summary_for_rows(self.filtered_sale_rows, "Saida", "Resumo Saidas")

    def open_sale_abc_popup(self) -> None:
        _periods, headers, _display_rows, export_rows = build_product_monthly_linear_dataset(self.filtered_sale_rows, "Saida")
        self.open_table_popup("Levantamento por Produto - Saida", headers, export_rows, 1360, 620)

    def open_sale_reduction_popup(self) -> None:
        reduction_rows = build_reduction_launch_rows(self.get_filtered_sale_launch_details())
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
            "Itens com Reducao de Base - Saidas",
            ["Tipo", "Data", "Documento", "Participante", "Chave", "Item", "Codigo", "Descricao", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Valor Operacao", "Base ICMS", "% Base/Oper", "Reducao BC", "Valor ICMS"],
            rows,
            1320,
            720,
        )

    def open_sale_apuracao_popup(self) -> None:
        entry_totals = {"icms_value": Decimal("0")}
        exit_totals = {
            "icms_value": sum((self.decimal_value(row.get("icms_value")) for row in self.filtered_sale_rows), Decimal("0")),
        }
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
        self.open_table_popup("Apuracao do ICMS - Saidas", ["Descricao", "Valor R$"], rows, 1020, 520)

    def open_sale_debit_diagnostic_popup(self) -> None:
        summary_rows, detail_rows, totals = build_credit_diagnostic_datasets(
            self.get_filtered_sale_summary_rows_from_details(),
            "Saida",
            self.parse_filter_tokens(self.sale_cst_input.text()),
            self.parse_filter_tokens(self.sale_cfop_input.text()),
            self.get_filtered_sale_launch_details(),
        )
        summary_headers = ["Periodo", "Motivo", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        summary_export_rows = [[row["period"], row["reason"], row["row_count"], row["total_operation_value"], row["base_icms"], row["base_gap"], row["base_ratio"], row["icms_value"]] for row in summary_rows]
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
            "Diagnostico Debito ICMS - Saidas",
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

    def open_sale_debit_comparison_popup(self) -> None:
        periods = sorted({str(self.first_row_value(row, "period", "period_label")) for row in self.filtered_sale_rows if str(self.first_row_value(row, "period", "period_label")).strip()})
        if len(periods) < 2:
            QMessageBox.warning(self, "Comparacao diagnostico de debito", "Selecione ao menos dois SPEDs/periodos para comparar.")
            return
        _periods, headers, display_rows, export_rows, _detail_rows = build_credit_diagnostic_period_comparison_dataset(
            self.get_filtered_sale_summary_rows_from_details(),
            self.get_filtered_sale_launch_details(),
            self.parse_filter_tokens(self.sale_cst_input.text()),
            self.parse_filter_tokens(self.sale_cfop_input.text()),
            "Saida",
        )
        rows = export_rows or [list(row.get("values", [])) for row in display_rows]
        self.open_table_popup("Comparacao Diagnostico de Debito", headers, rows, 1560, 760)

    def open_sale_docs_popup(self) -> None:
        self.open_documents_popup_for_details(self.get_filtered_sale_launch_details(), "Saidas")

    def open_xml_summary_details_popup(self) -> None:
        selected = self.xml_summary_table.selectionModel().selectedRows() if hasattr(self, "xml_summary_table") else []
        if not selected:
            QMessageBox.warning(self, "XML", "Selecione uma linha do resumo XML.")
            return
        row_index = selected[0].row()
        if row_index < 0 or row_index >= len(self.xml_summary_rows):
            return
        summary_row = self.xml_summary_rows[row_index]
        details = [detail for detail in summary_row.get("details", []) if isinstance(detail, dict)]
        rows = [
            [
                detail.get("key", ""),
                detail.get("number", ""),
                detail.get("series", ""),
                detail.get("issue_date", ""),
                detail.get("participant_name", detail.get("issuer_name", "")),
                detail.get("item_no", ""),
                detail.get("code", ""),
                detail.get("description", ""),
                detail.get("ncm", ""),
                detail.get("cfop", ""),
                detail.get("cst_icms", ""),
                self.decimal_value(detail.get("quantity")),
                self.decimal_value(detail.get("operation_value")),
                self.decimal_value(detail.get("base_icms")),
                self.decimal_value(detail.get("icms_value")),
                self.decimal_value(detail.get("base_icms_st")),
                self.decimal_value(detail.get("icms_st_value")),
                self.decimal_value(detail.get("ipi_value")),
            ]
            for detail in details
        ]
        self.open_table_popup(
            f"Detalhes XML - CFOP {summary_row.get('cfop', '')}",
            ["Chave", "Documento", "Serie", "Data", "Participante", "Item", "Codigo", "Descricao", "NCM", "CFOP", "CST", "Quantidade", "Valor Operacao", "Base ICMS", "Valor ICMS", "Base ST", "Valor ST", "IPI"],
            rows,
            1480,
            720,
        )

    def open_xml_credit_items_popup(self) -> None:
        selected = self.xml_credit_table.selectionModel().selectedRows() if hasattr(self, "xml_credit_table") else []
        if not selected:
            QMessageBox.warning(self, "XML", "Selecione uma nota de credito XML.")
            return
        row_index = selected[0].row()
        if row_index < 0 or row_index >= len(self.xml_credit_invoice_rows):
            return
        invoice_row = self.xml_credit_invoice_rows[row_index]
        details = [detail for detail in invoice_row.get("details", []) if isinstance(detail, dict)]
        rows = [
            [
                detail.get("item_no", ""),
                detail.get("code", ""),
                detail.get("description", ""),
                detail.get("ncm", ""),
                detail.get("cfop", ""),
                detail.get("cst_icms", ""),
                self.decimal_value(detail.get("quantity")),
                self.decimal_value(detail.get("operation_value")),
                self.decimal_value(detail.get("base_icms")),
                self.decimal_value(detail.get("icms_rate")),
                self.decimal_value(detail.get("icms_value")),
                self.decimal_value(detail.get("base_icms_st")),
                self.decimal_value(detail.get("icms_st_value")),
                self.decimal_value(detail.get("ipi_value")),
            ]
            for detail in details
        ]
        self.open_table_popup(
            f"Itens XML Credito - Documento {invoice_row.get('number', '')}",
            ["Item", "Codigo", "Descricao", "NCM", "CFOP", "CST", "Quantidade", "Valor Operacao", "Base ICMS", "Aliq ICMS", "Valor ICMS", "Base ST", "Valor ST", "IPI"],
            rows,
            1380,
            700,
        )

    def get_filtered_contrib_rows(self, operation_type: str) -> list[dict[str, object]]:
        return self.filtered_contrib_entry_rows if operation_type == "Entrada" else self.filtered_contrib_sale_rows

    def get_filtered_contrib_launch_details(self, operation_type: str) -> list[dict[str, object]]:
        details: list[dict[str, object]] = []
        for row in self.get_filtered_contrib_rows(operation_type):
            row_period = str(row.get("period", ""))
            launch_details = row.get("launch_details")
            source_details = launch_details if isinstance(launch_details, list) and launch_details else [row]
            for detail in source_details:
                if not isinstance(detail, dict):
                    continue
                enriched = dict(detail)
                enriched.setdefault("period", row_period)
                enriched.setdefault("operation_type", operation_type)
                details.append(enriched)
        return details

    def open_contrib_operation_summary_popup(self, operation_type: str) -> None:
        details = self.get_filtered_contrib_launch_details(operation_type)
        rows, totals = build_contrib_operation_summary_rows(details, operation_type)
        grouped_details = build_contrib_operation_launch_details_map(details, operation_type)
        export_rows: list[list[object]] = []
        for row in rows:
            key = (
                str(row.get("cst_pis", "")).strip(),
                str(row.get("cst_cofins", "")).strip(),
                str(row.get("cfop", "")).strip(),
                self.decimal_value(row.get("aliquota_pis")).quantize(Decimal("0.01")),
                self.decimal_value(row.get("aliquota_cofins")).quantize(Decimal("0.01")),
            )
            launch_details = grouped_details.get(key, [])
            document_keys = {
                str(detail.get("document_key", "") or detail.get("document_number", "")).strip()
                for detail in launch_details
                if str(detail.get("document_key", "") or detail.get("document_number", "")).strip()
            }
            export_rows.append(
                [
                    row.get("cst_pis", ""),
                    row.get("cst_cofins", ""),
                    row.get("cfop", ""),
                    self.decimal_value(row.get("aliquota_pis")),
                    self.decimal_value(row.get("aliquota_cofins")),
                    self.decimal_value(row.get("sale_value")),
                    self.decimal_value(row.get("base_pis")),
                    self.decimal_value(row.get("pis_value")),
                    self.decimal_value(row.get("base_cofins")),
                    self.decimal_value(row.get("cofins_value")),
                    len(document_keys),
                    row.get("launch_count", 0),
                ]
            )
        footer = (
            f"Linhas: {len(export_rows)}    Valor Operacao: {self.format_number(totals.get('sale_value', Decimal('0')))}    "
            f"Base PIS: {self.format_number(totals.get('base_pis', Decimal('0')))}    Valor PIS: {self.format_number(totals.get('pis_value', Decimal('0')))}    "
            f"Base COFINS: {self.format_number(totals.get('base_cofins', Decimal('0')))}    Valor COFINS: {self.format_number(totals.get('cofins_value', Decimal('0')))}"
        )
        self.open_table_popup(
            f"Resumo {operation_type}s PIS/COFINS",
            ["CST PIS", "CST COFINS", "CFOP", "Aliq PIS", "Aliq COFINS", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS", "Docs", "Lanc."],
            export_rows,
            1320,
            620,
            footer,
        )

    def open_contrib_diagnostic_popup(self, operation_type: str) -> None:
        grouped: dict[str, dict[str, object]] = {}
        for row in self.get_filtered_contrib_rows(operation_type):
            status = str(row.get("status", "")).strip() or "Sem status"
            bucket = grouped.setdefault(status, {"status": status, "lines": 0, "products": set(), "sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
            bucket["lines"] = int(bucket["lines"]) + 1
            bucket["products"].add(str(row.get("code", "")).strip())
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(row.get("sale_value"))
            bucket["base_pis"] = Decimal(bucket["base_pis"]) + self.decimal_value(row.get("base_pis"))
            bucket["pis_value"] = Decimal(bucket["pis_value"]) + self.decimal_value(row.get("pis_value"))
            bucket["base_cofins"] = Decimal(bucket["base_cofins"]) + self.decimal_value(row.get("base_cofins"))
            bucket["cofins_value"] = Decimal(bucket["cofins_value"]) + self.decimal_value(row.get("cofins_value"))
        rows = [
            [row["status"], row["lines"], len(row["products"]), row["sale_value"], row["base_pis"], row["pis_value"], row["base_cofins"], row["cofins_value"]]
            for row in sorted(grouped.values(), key=lambda item: self.normalize_search_text(item.get("status", "")))
        ]
        self.open_table_popup(
            f"Diagnostico PIS/COFINS - {operation_type}",
            ["Status", "Linhas", "Produtos", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS"],
            rows,
            1180,
            560,
        )

    def open_contrib_diagnostic_comparison_popup(self, operation_type: str) -> None:
        grouped: dict[tuple[str, str], dict[str, object]] = {}
        for row in self.get_filtered_contrib_rows(operation_type):
            key = (str(row.get("period", "")).strip(), str(row.get("status", "")).strip() or "Sem status")
            bucket = grouped.setdefault(key, {"period": key[0], "status": key[1], "lines": 0, "sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
            bucket["lines"] = int(bucket["lines"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(row.get("sale_value"))
            bucket["base_pis"] = Decimal(bucket["base_pis"]) + self.decimal_value(row.get("base_pis"))
            bucket["pis_value"] = Decimal(bucket["pis_value"]) + self.decimal_value(row.get("pis_value"))
            bucket["base_cofins"] = Decimal(bucket["base_cofins"]) + self.decimal_value(row.get("base_cofins"))
            bucket["cofins_value"] = Decimal(bucket["cofins_value"]) + self.decimal_value(row.get("cofins_value"))
        rows = [
            [row["period"], row["status"], row["lines"], row["sale_value"], row["base_pis"], row["pis_value"], row["base_cofins"], row["cofins_value"]]
            for row in sorted(grouped.values(), key=lambda item: (str(item.get("period", "")), self.normalize_search_text(item.get("status", ""))))
        ]
        self.open_table_popup(
            f"Comparativo Diagnostico PIS/COFINS - {operation_type}",
            ["Periodo", "Status", "Linhas", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS"],
            rows,
            1240,
            580,
        )

    def open_contrib_abc_popup(self, operation_type: str) -> None:
        _periods, headers, _display_rows, export_rows = build_contrib_product_monthly_linear_dataset(self.get_filtered_contrib_rows(operation_type), operation_type)
        self.open_table_popup(f"Levantamento por Produto PIS/COFINS - {operation_type}", headers, export_rows, 1460, 700)

    def open_contrib_apuracao_popup(self, operation_type: str) -> None:
        grouped: dict[str, dict[str, Decimal]] = {}
        for detail in self.get_filtered_contrib_launch_details(operation_type):
            period = str(detail.get("period", "")).strip() or "(sem periodo)"
            bucket = grouped.setdefault(period, {"sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
            bucket["sale_value"] += self.decimal_value(detail.get("sale_value"))
            bucket["base_pis"] += self.decimal_value(detail.get("base_pis"))
            bucket["pis_value"] += self.decimal_value(detail.get("pis_value"))
            bucket["base_cofins"] += self.decimal_value(detail.get("base_cofins"))
            bucket["cofins_value"] += self.decimal_value(detail.get("cofins_value"))
        rows = []
        for period, totals in sorted(grouped.items(), key=lambda item: item[0]):
            pis_rate = (totals["pis_value"] * Decimal("100") / totals["base_pis"]).quantize(Decimal("0.01")) if totals["base_pis"] > 0 else Decimal("0.00")
            cofins_rate = (totals["cofins_value"] * Decimal("100") / totals["base_cofins"]).quantize(Decimal("0.01")) if totals["base_cofins"] > 0 else Decimal("0.00")
            rows.append([period, totals["sale_value"], totals["base_pis"], totals["pis_value"], pis_rate, totals["base_cofins"], totals["cofins_value"], cofins_rate])
        self.open_table_popup(
            f"Apuracao PIS/COFINS - {operation_type}",
            ["Periodo", "Valor Operacao", "Base PIS", "Valor PIS", "Aliq PIS Efetiva", "Base COFINS", "Valor COFINS", "Aliq COFINS Efetiva"],
            rows,
            1120,
            560,
        )

    def open_contrib_docs_popup(self, operation_type: str) -> None:
        grouped: dict[tuple[str, str, str], dict[str, object]] = {}
        for detail in self.get_filtered_contrib_launch_details(operation_type):
            key = (str(detail.get("period", "")), str(detail.get("document_number", "")), str(detail.get("document_key", "")))
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
                    "base_pis": Decimal("0"),
                    "pis_value": Decimal("0"),
                    "base_cofins": Decimal("0"),
                    "cofins_value": Decimal("0"),
                },
            )
            bucket["items"] = int(bucket["items"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(detail.get("sale_value"))
            bucket["base_pis"] = Decimal(bucket["base_pis"]) + self.decimal_value(detail.get("base_pis"))
            bucket["pis_value"] = Decimal(bucket["pis_value"]) + self.decimal_value(detail.get("pis_value"))
            bucket["base_cofins"] = Decimal(bucket["base_cofins"]) + self.decimal_value(detail.get("base_cofins"))
            bucket["cofins_value"] = Decimal(bucket["cofins_value"]) + self.decimal_value(detail.get("cofins_value"))
        rows = [
            [
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
                row["base_pis"],
                row["pis_value"],
                row["base_cofins"],
                row["cofins_value"],
            ]
            for row in sorted(grouped.values(), key=lambda item: (str(item["period"]), str(item["document_date"]), str(item["document_number"])))
        ]
        self.open_table_popup(
            f"Espelho de Documentos PIS/COFINS - {operation_type}s",
            ["Periodo", "Data", "Documento", "Serie", "Modelo", "Participante", "CPF/CNPJ", "Chave", "Itens", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS"],
            rows,
            1460,
            720,
        )

    def open_operation_summary_for_rows(self, source_rows: list[dict[str, object]], operation_type: str, title: str) -> None:
        grouped: dict[tuple[str, str, Decimal], dict[str, object]] = {}
        for row in source_rows:
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
                    {"sale_value": Decimal("0"), "base_icms": Decimal("0"), "icms_value": Decimal("0"), "base_icms_st": Decimal("0"), "icms_st_value": Decimal("0"), "ipi_value": Decimal("0"), "document_keys": set(), "launch_count": 0},
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
        rows = []
        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for (cst, cfop, rate), values in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1], item[0][2])):
            sale_value = Decimal(values["sale_value"])
            base_icms = Decimal(values["base_icms"])
            icms_value = Decimal(values["icms_value"])
            rows.append((cst, cfop, rate, compute_display_icms_rate(rate, sale_value, base_icms, icms_value), Decimal(values["ipi_value"]), icms_value, Decimal(values["base_icms_st"]), Decimal(values["icms_st_value"]), sale_value, base_icms, (sale_value - base_icms).quantize(Decimal("0.01")), (sale_value - base_icms).quantize(Decimal("0.01")), len(values["document_keys"]), values["launch_count"]))
            total_sale += sale_value
            total_base += base_icms
            total_icms += icms_value
        ratio = (total_base * Decimal("100") / total_sale).quantize(Decimal("0.01")) if total_sale else Decimal("0.00")
        footer = f"Linhas: {len(rows)}    Total Operacao: {self.format_number(total_sale)}    Base ICMS: {self.format_number(total_base)}    % Base/Oper: {self.format_number(ratio)}    Valor ICMS: {self.format_number(total_icms)}"
        self.open_table_popup(title, ["CST", "CFOP", "Aliq ICMS", "Aliq Efetiva", "Valor IPI", "Valor ICMS", "Base ICMS ST", "Valor ICMS ST", "Total Operacao", "Base ICMS", "Dif. Oper/Base", "Reducao BC", "Docs", "Lanc."], rows, 1220, 560, footer)

    def open_documents_popup_for_details(self, details: list[dict[str, object]], caption: str) -> None:
        grouped: dict[tuple[str, str, str], dict[str, object]] = {}
        for detail in details:
            key = (str(detail.get("period", "")), str(detail.get("document_number", "")), str(detail.get("document_key", "")))
            bucket = grouped.setdefault(
                key,
                {"period": detail.get("period", ""), "document_date": detail.get("document_date", detail.get("date", "")), "document_number": detail.get("document_number", ""), "document_series": detail.get("document_series", ""), "document_model": detail.get("document_model", ""), "participant_name": detail.get("participant_name", ""), "participant_tax_id": detail.get("participant_tax_id", ""), "document_key": detail.get("document_key", ""), "items": 0, "sale_value": Decimal("0"), "base_icms": Decimal("0"), "icms_value": Decimal("0")},
            )
            bucket["items"] = int(bucket["items"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            bucket["base_icms"] = Decimal(bucket["base_icms"]) + self.decimal_value(detail.get("base_icms"))
            bucket["icms_value"] = Decimal(bucket["icms_value"]) + self.decimal_value(detail.get("icms_value"))
        rows = [(row["period"], row["document_date"], row["document_number"], row["document_series"], row["document_model"], row["participant_name"], row["participant_tax_id"], row["document_key"], row["items"], row["sale_value"], row["base_icms"], row["icms_value"]) for row in sorted(grouped.values(), key=lambda item: (str(item["period"]), str(item["document_date"]), str(item["document_number"])))]
        self.open_table_popup(f"Espelho de Documentos Fiscais - {caption}", ["Periodo", "Data", "Documento", "Serie", "Modelo", "Participante", "CPF/CNPJ", "Chave", "Itens", "Valor Operacao", "Base ICMS", "Valor ICMS"], rows, 1380, 720)
