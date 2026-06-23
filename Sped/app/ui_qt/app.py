from __future__ import annotations

import csv
import datetime as dt
import os
import subprocess
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
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
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QMenu,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.config import APP_DEFAULT_CONFIG, MYSQL_DEFAULT_CONFIG
from app.exporters.workbook_exporter import write_simple_csv_file, write_simple_excel_workbook
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.parsers.sped_fiscal_parser import read_sped_file
from app.parsers.sped_parser import get_field, normalize_document_key, normalize_sped_line
from app.services.app_paths import get_application_base_dir, get_application_environment, get_audit_log_path, get_environment_config_path, get_project_root_dir, get_runtime_rule_history_path
from app.services.app_config_service import load_app_config
from app.services.catalog_xml_import import (
    CatalogImportPreviewRow,
    UPDATABLE_PRODUCT_FIELDS,
    build_catalog_import_existing_review,
    build_catalog_import_preview,
    import_catalogs_from_preview,
)
from app.services.compare_operations import filter_xml_summary_rows_by_scope, normalize_compare_operation_scope
from app.services.compare_workflows import build_xml_cfop_summary_rows, build_xml_entry_credit_rows, compare_sped_with_sheet, compare_sped_with_xml_folder
from app.services.analysis_reports import (
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_contrib_operation_launch_details_map,
    build_entry_exit_analysis_rows,
    build_entry_exit_footer_rows,
    build_contrib_operation_summary_rows,
    build_contrib_product_monthly_linear_dataset,
    build_product_monthly_linear_dataset,
    write_entry_exit_analysis_excel,
)
from app.services.operation_summary import build_filtered_apuracao_rows, build_reduction_launch_rows
from app.services.path_selection import append_unique_paths, collapse_xml_selection_paths, format_selected_paths, limit_selected_paths, parse_selected_paths
from app.services.period_comparisons import build_entry_period_comparison_rows, build_sale_period_comparison_rows
from app.services.sped_archive import archive_original_sped_file
from app.services.adjusted_sped import write_adjusted_sped
from app.services.audit_utils import format_audit_paths
from app.services.tax_rules import compute_display_icms_rate, normalize_text
from app.services.runtime_rules import extract_tax_id_from_document_key, get_first_matching_icms_rule, parse_runtime_rule_lines, runtime_rule_summary
from app.services.runtime_rule_history import (
    load_runtime_rule_history as load_runtime_rule_history_file,
    remember_runtime_rules as remember_runtime_rules_in_history,
    remove_runtime_rule,
    save_runtime_rule_history as save_runtime_rule_history_file,
)
from app.services.xml_reconciliation import build_pis_cofins_period_comparison_rows
from app.ui_qt.import_planilha import ImportPlanilhaDialog


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

CATALOG_IMPORT_POPUP_ROW_LIMIT = 2000

USER_PERMISSION_GROUPS: tuple[tuple[str, tuple[tuple[str, str], ...]], ...] = (
    (
        "Consultas",
        (
            ("dashboard", "Dashboard"),
            ("icms_ipi_entradas", "Icms/Ipi Entradas"),
            ("icms_ipi_saidas", "Icms/Ipi Saidas"),
            ("pis_cofins_entradas", "Pis/Cofins Entradas"),
            ("pis_cofins_saidas", "Pis/Cofins Saidas"),
            ("xml", "Xml"),
            ("sped_xml", "Sped X Xml"),
            ("analise_entrada_saida", "Analise Entrada E Saida"),
        ),
    ),
    (
        "Cadastros",
        (
            ("empresas", "Empresas"),
            ("fornecedores", "Fornecedores"),
            ("classificacao_produto", "Classificacao Do Produto"),
            ("produtos", "Produtos"),
            ("ncm", "Ncm"),
            ("limpeza_duplicados", "Limpeza Duplicados"),
        ),
    ),
    (
        "Sistema",
        (
            ("regras_dinamicas", "Regras Dinamicas"),
            ("speds_arquivados", "Speds Arquivados"),
            ("importacao_xml_cadastros", "Importacao Xml Cadastros"),
            ("usuarios_permissoes", "Usuarios E Permissoes"),
            ("configuracoes", "Configuracoes"),
            ("extrair_chave_nfe", "Extrair Chave Nfe"),
        ),
    ),
)


class LoginDialog(QDialog):
    def __init__(self, app_title: str, environment: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Acesso Ao Sistema")
        self.setModal(True)
        self.resize(420, 240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        title = QLabel(app_title)
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        env_label = QLabel(f"Ambiente: {environment}")
        env_label.setObjectName("muted")
        layout.addWidget(env_label)

        self.login_input = QLineEdit()
        self.login_input.setPlaceholderText("Login")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.message_label = QLabel("")
        self.message_label.setObjectName("muted")

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.addWidget(QLabel("Usuario"), 0, 0)
        form.addWidget(self.login_input, 0, 1)
        form.addWidget(QLabel("Senha"), 1, 0)
        form.addWidget(self.password_input, 1, 1)
        layout.addLayout(form)
        layout.addWidget(self.message_label)

        actions = QHBoxLayout()
        actions.addStretch()
        cancel_button = QPushButton("Cancelar")
        cancel_button.setAutoDefault(False)
        cancel_button.setDefault(False)
        cancel_button.clicked.connect(self.reject)
        enter_button = QPushButton("Entrar")
        enter_button.setObjectName("primaryButton")
        enter_button.setAutoDefault(True)
        enter_button.setDefault(True)
        enter_button.clicked.connect(self.accept)
        actions.addWidget(cancel_button)
        actions.addWidget(enter_button)
        layout.addLayout(actions)

        self.login_input.setText("admin")
        self.password_input.setFocus()
        self.password_input.returnPressed.connect(self.accept)

    def credentials(self) -> tuple[str, str]:
        return self.login_input.text().strip(), self.password_input.text()


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


class NFeKeyExtractWorker(QObject):
    finished = Signal(list)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, sped_paths: list[Path]) -> None:
        super().__init__()
        self.sped_paths = sped_paths

    def run(self) -> None:
        try:
            rows: list[dict[str, object]] = []
            total_files = len(self.sped_paths)
            for index, sped_path in enumerate(self.sped_paths, start=1):
                self.progress.emit(index - 1, total_files, f"Lendo {sped_path.name}...")
                rows.extend(extract_nfe_keys_from_sped_file(sped_path))
            rows.sort(
                key=lambda item: (
                    str(item.get("file_name", "")).lower(),
                    str(item.get("operation_type", "")),
                    str(item.get("document_date", "")),
                    str(item.get("document_number", "")),
                    str(item.get("document_key", "")),
                )
            )
            self.progress.emit(total_files, total_files, "Finalizando carga...")
            self.finished.emit(rows)
        except Exception as exc:
            self.failed.emit(str(exc))


class SpedReprocessWorker(QObject):
    finished = Signal(list)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(
        self,
        sped_paths: list[Path],
        output_paths: list[Path],
        operation_type: str,
        runtime_rules: list[dict[str, object]],
        rule_lines: list[str],
        audit_log_path: Path,
        audit_session_id: str,
    ) -> None:
        super().__init__()
        self.sped_paths = sped_paths
        self.output_paths = output_paths
        self.operation_type = operation_type
        self.runtime_rules = runtime_rules
        self.rule_lines = rule_lines
        self.audit_log_path = audit_log_path
        self.audit_session_id = audit_session_id

    def write_audit_log(self, event_type: str, message: str) -> None:
        try:
            timestamp = dt.datetime.now().isoformat(timespec="seconds")
            event = str(event_type or "").strip() or "EVENTO"
            text = str(message or "").replace("\n", " ")
            self.audit_log_path.parent.mkdir(parents=True, exist_ok=True)
            with self.audit_log_path.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{timestamp}\t{self.audit_session_id}\t{event}\t{text}\n")
        except Exception:
            pass

    def run(self) -> None:
        try:
            generated_files: list[str] = []
            total_files = len(self.sped_paths)
            for index, (sped_path, output_path) in enumerate(zip(self.sped_paths, self.output_paths), start=1):
                self.progress.emit(index - 1, total_files, f"Lendo {sped_path.name}...")
                _products, _sales_rows, detailed_sales, _c190_rows, _c190_product_rows = read_sped_file(sped_path)
                self.write_audit_log(
                    "REPROCESSAMENTO_INICIO",
                    f"sped_origem={sped_path}; sped_saida={output_path}; escopo={self.operation_type}; regras_dinamicas={len(self.rule_lines)}",
                )
                self.progress.emit(index - 1, total_files, f"Gravando {output_path.name}...")
                write_adjusted_sped(
                    output_path,
                    sped_path,
                    detailed_sales,
                    [],
                    set(),
                    set(),
                    None,
                    "",
                    "",
                    set(),
                    self.operation_type,
                    "",
                    [],
                    set(),
                    self.runtime_rules,
                )
                log_path = output_path.with_name(f"{output_path.stem}_log_ajustes.xlsx")
                generated_files.append(str(output_path))
                if log_path.exists():
                    generated_files.append(str(log_path))
                self.write_audit_log(
                    "REPROCESSAMENTO_FIM",
                    f"sped_origem={sped_path}; arquivo_reprocessado={output_path}; extras={format_audit_paths([log_path])}",
                )
                self.progress.emit(index, total_files, f"Concluido {output_path.name}.")
            self.finished.emit(generated_files)
        except Exception as exc:
            self.write_audit_log("REPROCESSAMENTO_ERRO", str(exc))
            self.failed.emit(str(exc))


class CatalogImportPreviewWorker(QObject):
    finished = Signal(list, dict)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, repository: MysqlCadastroRepository, environment: str, source_value: str) -> None:
        super().__init__()
        self.repository = repository
        self.environment = environment
        self.source_value = source_value

    def run(self) -> None:
        try:
            preview_rows, stats = build_catalog_import_preview(
                self.repository,
                self.environment,
                self.source_value,
                progress_callback=self.progress.emit,
            )
            self.finished.emit(preview_rows, stats)
        except Exception as exc:
            self.failed.emit(str(exc))


class CatalogImportWorker(QObject):
    finished = Signal(dict, list, list, str)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(
        self,
        repository: MysqlCadastroRepository,
        environment: str,
        preview_rows: list[CatalogImportPreviewRow],
        allow_update_existing: bool,
        allow_insert_new: bool,
        update_fields: set[str],
    ) -> None:
        super().__init__()
        self.repository = repository
        self.environment = environment
        self.preview_rows = preview_rows
        self.allow_update_existing = allow_update_existing
        self.allow_insert_new = allow_insert_new
        self.update_fields = update_fields

    def run(self) -> None:
        try:
            duplicate_rows: list[list[object]] = []
            change_rows: list[list[object]] = []
            backup_path = ""
            if self.allow_update_existing:
                self.progress.emit(0, 3, "Conferindo produtos existentes em lote...")
                change_rows, duplicate_rows, product_ids_to_backup = build_catalog_import_existing_review(
                    self.repository,
                    self.environment,
                    self.preview_rows,
                    self.update_fields,
                    self.progress.emit,
                )
                if product_ids_to_backup:
                    self.progress.emit(2, 3, "Criando backup de seguranca dos produtos existentes...")
                    backup_path = str(self.repository.backup_supplier_products_by_ids(product_ids_to_backup))
            stats = import_catalogs_from_preview(
                self.repository,
                self.environment,
                self.preview_rows,
                allow_update_existing=self.allow_update_existing,
                allow_insert_new=self.allow_insert_new,
                progress_callback=self.progress.emit,
                update_fields=self.update_fields,
            )
            self.finished.emit(stats, duplicate_rows, change_rows, backup_path)
        except Exception as exc:
            self.failed.emit(str(exc))


class ProductCatalogWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, repository: MysqlCadastroRepository, environment: str) -> None:
        super().__init__()
        self.repository = repository
        self.environment = environment

    def run(self) -> None:
        try:
            self.repository.ensure_schema()
            payload = {
                "products": self.repository.list_products_catalog(self.environment),
                "suppliers": self.repository.list_suppliers_catalog(self.environment),
                "types": self.repository.list_product_types(self.environment),
            }
            self.finished.emit(payload)
        except Exception as exc:
            self.failed.emit(str(exc))


class ProductTableModel(QAbstractTableModel):
    def __init__(self, headers: list[str]) -> None:
        super().__init__()
        self.headers = headers
        self.rows: list[list[object]] = []

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.headers)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> object:
        if not index.isValid():
            return None
        if role == Qt.DisplayRole:
            return str(self.rows[index.row()][index.column()])
        if role == Qt.ForegroundRole:
            return QColor("#1d2730")
        if role == Qt.TextAlignmentRole:
            header = self.headers[index.column()].strip().lower()
            return Qt.AlignVCenter | Qt.AlignLeft if header in {"descrição", "descricao", "produto"} else Qt.AlignCenter
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> object:
        if role == Qt.DisplayRole and orientation == Qt.Horizontal and 0 <= section < len(self.headers):
            return self.headers[section]
        return None

    def set_rows(self, rows: list[list[object]]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def row_id(self, row_index: int) -> int:
        if row_index < 0 or row_index >= len(self.rows):
            return 0
        value = self.rows[row_index][1] if len(self.rows[row_index]) > 1 else 0
        return int(value) if str(value).isdigit() else 0


class ProductReviewImportWorker(QObject):
    finished = Signal(dict)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, repository: MysqlCadastroRepository, environment: str, company_id: int, excel_path: Path) -> None:
        super().__init__()
        self.repository = repository
        self.environment = environment
        self.company_id = company_id
        self.excel_path = excel_path

    def run(self) -> None:
        try:
            stats = self.repository.import_reviewed_products_from_excel(
                self.environment,
                self.company_id,
                self.excel_path,
                progress_callback=self.progress.emit,
            )
            self.finished.emit(stats)
        except Exception as exc:
            self.failed.emit(str(exc))


class BackupDatabaseWorker(QObject):
    finished = Signal(Path)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, connection_config: dict[str, object], dest_path: Path) -> None:
        super().__init__()
        self.connection_config = connection_config
        self.dest_path = dest_path

    def run(self) -> None:
        from app.services.database_backup import backup_database
        try:
            backup_database(self.connection_config, self.dest_path, self.progress.emit)
            self.finished.emit(self.dest_path)
        except Exception as exc:
            self.failed.emit(str(exc))


def extract_nfe_keys_from_sped_file(sped_path: Path) -> list[dict[str, object]]:
    def classify_document_type(document_model: str, document_key: str) -> str:
        model = str(document_model or "").strip()
        key_digits = normalize_document_key(document_key)
        if model in {"55", "65"} or len(key_digits) == 44:
            return "NFe"
        return "NFSe"

    rows: list[dict[str, object]] = []
    participants: dict[str, dict[str, str]] = {}
    with sped_path.open("r", encoding="latin-1") as sped_file:
        for raw_line in sped_file:
            if not raw_line.startswith("|"):
                continue
            fields = normalize_sped_line(raw_line)
            register = get_field(fields, 1)
            if register == "0150":
                participant_code = get_field(fields, 2)
                participants[participant_code] = {
                    "name": get_field(fields, 3),
                    "tax_id": get_field(fields, 5) or get_field(fields, 4),
                }
                continue
            if register != "C100":
                continue
            ind_oper = get_field(fields, 2)
            operation_type = "Entrada" if ind_oper == "0" else "Saida" if ind_oper == "1" else ""
            document_model = str(get_field(fields, 6) or "").strip()
            raw_document_key = str(get_field(fields, 9) or "").strip()
            normalized_document_key = normalize_document_key(raw_document_key)
            document_key = normalized_document_key if normalized_document_key else raw_document_key
            if not document_key:
                continue
            participant_code = get_field(fields, 4)
            participant_data = participants.get(participant_code, {})
            rows.append(
                {
                    "file_name": sped_path.name,
                    "operation_type": operation_type,
                    "document_type": classify_document_type(document_model, document_key),
                    "document_key": document_key,
                    "document_model": document_model,
                    "document_number": get_field(fields, 8),
                    "document_series": get_field(fields, 7),
                    "document_date": get_field(fields, 10),
                    "participant_name": participant_data.get("name", ""),
                    "participant_tax_id": participant_data.get("tax_id", ""),
                }
            )
    return rows


class QtSpedApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.scale = self._compute_scale()
        self.environment = get_application_environment()
        self.base_dir = get_application_base_dir(__file__)
        self.project_root_dir = get_project_root_dir(self.base_dir)
        self.app_config_path = get_environment_config_path(self.project_root_dir / "app" / "ui", "app_config", self.environment)
        self.app_config = load_app_config(self.app_config_path, self.get_app_default_config())
        self.app_window_title = str(self.app_config.get("window_title", self.get_app_default_config()["window_title"])).strip() or self.get_app_default_config()["window_title"]
        self.app_home_title = str(self.app_config.get("home_title", self.get_app_default_config()["home_title"])).strip() or self.get_app_default_config()["home_title"]
        self.mysql_config_path = get_environment_config_path(self.project_root_dir / "app" / "ui", "mysql_config", self.environment)
        self.runtime_rule_history_path = get_runtime_rule_history_path(self.project_root_dir / "app" / "ui")
        self.runtime_rule_history: list[dict[str, object]] = load_runtime_rule_history_file(self.runtime_rule_history_path)
        self.current_runtime_rules: list[dict[str, object]] = []
        self.audit_log_path = get_audit_log_path(self.project_root_dir / "app" / "ui")
        self.audit_session_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
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
        self.nfe_extract_rows: list[dict[str, object]] = []
        self.nfe_extract_all_rows: list[dict[str, object]] = []
        self.entry_exit_detail_rows: list[dict[str, object]] = []
        self.entry_exit_excel_rows: list[list[object]] = []
        self.entry_exit_totals: dict[str, Decimal] = {}

        self.setWindowTitle(self.app_window_title)
        self.resize(*self.initial_window_size(1360, 820))
        self.apply_styles()
        self.build_shell()
        self.current_user: dict[str, object] | None = None
        self.login_cancelled = False
        self.ensure_default_admin_user()
        self.show_page(0, "Dashboard")

    def authenticate_on_startup(self) -> bool:
        if self.show_login_dialog():
            self.show_page(0, "Dashboard")
            self.showMaximized()
            self.raise_()
            self.activateWindow()
            return True
        self.login_cancelled = True
        return False

    def get_app_default_config(self) -> dict[str, str]:
        return dict(APP_DEFAULT_CONFIG)

    def get_mysql_default_config(self) -> dict[str, str]:
        config = dict(MYSQL_DEFAULT_CONFIG)
        if self.environment == "dev":
            config["database"] = "sped_icms_dev"
        elif self.environment == "prod":
            config["database"] = "sped_icms"
        return config

    def _compute_scale(self) -> float:
        screen = QApplication.primaryScreen()
        if screen is None:
            return 1.0
        geom = screen.availableGeometry()
        scale = geom.height() / 1080.0
        return max(0.75, min(scale, 2.0))

    def _s(self, value: int) -> int:
        return max(1, round(value * self.scale))

    def apply_styles(self) -> None:
        s = self._s
        self.setStyleSheet(
            f"""
            QMainWindow, QWidget#contentHost, QDialog#popupDialog {{
                background: {COLORS["bg"]};
                color: {COLORS["text"]};
                font-family: Segoe UI;
                font-size: {s(14)}px;
            }}
            QStackedWidget, QScrollArea#pageScroll, QWidget#pageScrollViewport, QWidget#scrollPage {{
                background: {COLORS["bg"]};
                border: 0;
            }}
            QMessageBox {{
                background: {COLORS["panel"]};
                color: {COLORS["text"]};
            }}
            QLabel {{
                color: {COLORS["text"]};
                background: transparent;
            }}
            QMessageBox QLabel {{
                color: {COLORS["text"]};
                background: transparent;
            }}
            QWidget#sidebar {{
                background: {COLORS["sidebar"]};
            }}
            QLabel#brand {{
                color: {COLORS["sidebar_text"]};
                font-size: {s(21)}px;
                font-weight: 700;
            }}
            QPushButton#navButton {{
                background: transparent;
                color: #d9e5ef;
                border: 0;
                border-radius: {s(6)}px;
                padding: {s(11)}px {s(12)}px;
                text-align: left;
                font-weight: 600;
            }}
            QPushButton#navButton:hover, QPushButton#navButton:checked {{
                background: {COLORS["sidebar_active"]};
                color: #ffffff;
            }}
            QPushButton#navGroup {{
                background: transparent;
                color: #9fb3c8;
                border: 0;
                border-radius: {s(6)}px;
                font-size: {s(16)}px;
                font-weight: 900;
                padding: {s(13)}px {s(8)}px {s(9)}px {s(8)}px;
                text-align: left;
            }}
            QPushButton#navGroup:hover, QPushButton#navGroup:checked {{
                background: #203041;
                color: #ffffff;
            }}
            QScrollArea#navScroll {{
                background: transparent;
                border: 0;
            }}
            QScrollArea#navScroll QWidget {{
                background: transparent;
            }}
            QPushButton#sidebarToggle {{
                background: #203041;
                color: #ffffff;
                border: 0;
                border-radius: {s(6)}px;
                padding: {s(9)}px {s(10)}px;
                font-size: {s(16)}px;
                font-weight: 900;
            }}
            QPushButton#sidebarToggle:hover {{
                background: {COLORS["sidebar_active"]};
            }}
            QLabel#pageTitle {{
                color: {COLORS["text"]};
                font-size: {s(24)}px;
                font-weight: 750;
            }}
            QLabel#envBadgeDev {{
                background: #274766;
                color: #ffffff;
                border-radius: {s(6)}px;
                padding: {s(7)}px {s(8)}px;
                font-weight: 800;
            }}
            QLabel#envBadgeProd {{
                background: #8f1d2c;
                color: #ffffff;
                border-radius: {s(6)}px;
                padding: {s(7)}px {s(8)}px;
                font-weight: 900;
            }}
            QLabel#muted {{
                color: {COLORS["muted"]};
                font-size: {s(13)}px;
            }}
            QFrame#panel {{
                background: {COLORS["panel"]};
                border: 1px solid {COLORS["line"]};
                border-radius: {s(8)}px;
            }}
            QLabel#sectionTitle {{
                color: {COLORS["text"]};
                font-size: {s(15)}px;
                font-weight: 750;
            }}
            QLabel#popupTitle {{
                color: {COLORS["text"]};
                font-size: {s(18)}px;
                font-weight: 800;
            }}
            QLabel#popupFooter {{
                color: {COLORS["text"]};
                font-size: {s(14)}px;
                font-weight: 750;
            }}
            QLabel#metricKey {{
                color: {COLORS["muted"]};
                font-size: {s(12)}px;
                font-weight: 700;
            }}
            QLabel#metricValue {{
                color: {COLORS["text"]};
                font-size: {s(18)}px;
                font-weight: 800;
            }}
            QPushButton {{
                background: #ffffff;
                border: 1px solid #bdcbd7;
                border-radius: {s(6)}px;
                color: {COLORS["text"]};
                padding: {s(8)}px {s(12)}px;
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
            QLineEdit, QComboBox, QTextEdit {{
                background: #ffffff;
                border: 1px solid #c8d4df;
                border-radius: {s(6)}px;
                color: {COLORS["text"]};
                padding: {s(7)}px {s(9)}px;
                min-height: {s(22)}px;
            }}
            QTextEdit {{
                selection-background-color: #dbeafe;
                selection-color: {COLORS["text"]};
            }}
            QComboBox QAbstractItemView {{
                background: #ffffff;
                color: {COLORS["text"]};
                selection-background-color: #dbeafe;
                selection-color: {COLORS["text"]};
            }}
            QCheckBox {{
                color: {COLORS["text"]};
                spacing: {s(7)}px;
            }}
            QGroupBox {{
                background: #ffffff;
                border: 1px solid {COLORS["line"]};
                border-radius: {s(8)}px;
                color: {COLORS["text"]};
                font-weight: 750;
                margin-top: {s(18)}px;
                padding: {s(12)}px {s(10)}px {s(10)}px {s(10)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: {s(10)}px;
                padding: 0 {s(6)}px;
                color: {COLORS["text"]};
                background: {COLORS["bg"]};
            }}
            QTabWidget::pane {{
                border: 1px solid {COLORS["line"]};
                border-radius: {s(8)}px;
                background: #ffffff;
                top: -1px;
            }}
            QTabWidget, QTabBar {{
                background: {COLORS["bg"]};
            }}
            QTabBar::tab {{
                background: #edf3f7;
                color: {COLORS["text"]};
                border: 1px solid {COLORS["line"]};
                border-bottom-color: {COLORS["line"]};
                border-top-left-radius: {s(6)}px;
                border-top-right-radius: {s(6)}px;
                padding: {s(8)}px {s(14)}px;
                margin-right: {s(4)}px;
                min-width: {s(92)}px;
                font-weight: 700;
            }}
            QTabBar::tab:selected {{
                background: {COLORS["accent"]};
                color: #ffffff;
                border-color: {COLORS["accent"]};
            }}
            QTabBar::tab:hover:!selected {{
                background: #dbeafe;
                color: {COLORS["text"]};
            }}
            QWidget#productFormTab QLineEdit, QWidget#productFormTab QComboBox {{
                padding: {s(3)}px {s(8)}px;
                min-height: {s(18)}px;
                max-height: {s(28)}px;
            }}
            QWidget#productFormTab QLabel {{
                font-size: {s(13)}px;
            }}
            QTabWidget#productTribTabs::pane {{
                top: -1px;
            }}
            QTabWidget#productTribTabs QTabBar::tab {{
                padding: {s(6)}px {s(14)}px;
                min-width: {s(82)}px;
            }}
            QTableWidget, QTableView {{
                background: #ffffff;
                alternate-background-color: #fbfdff;
                border: 1px solid {COLORS["line"]};
                border-radius: {s(8)}px;
                color: {COLORS["text"]};
                gridline-color: #e7edf2;
                selection-background-color: #dbeafe;
                selection-color: {COLORS["text"]};
            }}
            QTableWidget::item, QTableView::item {{
                color: {COLORS["text"]};
                background: transparent;
                padding: {s(4)}px;
            }}
            QTableView::item:alternate {{
                background: #fbfdff;
            }}
            QTableWidget::item:selected, QTableView::item:selected {{
                background: #dbeafe;
                color: {COLORS["text"]};
            }}
            QScrollBar:vertical {{
                background: #edf3f7;
                width: {s(16)}px;
                margin: 0;
                border-left: 1px solid {COLORS["line"]};
            }}
            QScrollBar::handle:vertical {{
                background: #8fa5b8;
                min-height: {s(32)}px;
                border-radius: {s(7)}px;
                margin: {s(2)}px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #6f879d;
            }}
            QScrollBar:horizontal {{
                background: #edf3f7;
                height: {s(16)}px;
                margin: 0;
                border-top: 1px solid {COLORS["line"]};
            }}
            QScrollBar::handle:horizontal {{
                background: #8fa5b8;
                min-width: {s(32)}px;
                border-radius: {s(7)}px;
                margin: {s(2)}px;
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
                padding: {s(8)}px {s(9)}px;
                font-size: {s(12)}px;
                font-weight: 750;
            }}
            QStatusBar {{
                background: {COLORS["bg"]};
                color: {COLORS["muted"]};
            }}
            """
        )

    def available_screen_geometry(self):
        screen = self.screen() or QApplication.primaryScreen()
        if screen is None:
            return None
        return screen.availableGeometry()

    def initial_window_size(self, preferred_width: int, preferred_height: int) -> tuple[int, int]:
        geometry = self.available_screen_geometry()
        if geometry is None:
            return preferred_width, preferred_height
        return (
            min(preferred_width, max(720, int(geometry.width() * 0.92))),
            min(preferred_height, max(560, int(geometry.height() * 0.88))),
        )

    def resize_dialog_to_screen(self, dialog: QDialog, preferred_width: int, preferred_height: int) -> None:
        geometry = self.available_screen_geometry()
        if geometry is None:
            dialog.resize(preferred_width, preferred_height)
            return
        width = min(preferred_width, max(420, int(geometry.width() * 0.94)))
        height = min(preferred_height, max(360, int(geometry.height() * 0.90)))
        dialog.resize(width, height)

    def make_scrollable_page(self, page: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setObjectName("pageScroll")
        scroll.viewport().setObjectName("pageScrollViewport")
        if not page.objectName():
            page.setObjectName("scrollPage")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        page.setAutoFillBackground(True)
        scroll.setWidget(page)
        return scroll

    def build_shell(self) -> None:
        root = QWidget()
        root.setObjectName("contentHost")
        shell = QHBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(self._s(230))
        self.sidebar = sidebar
        self.sidebar_expanded_width = self._s(260)
        self.sidebar_collapsed_width = self._s(76)
        self.sidebar_collapsed = False
        nav_layout = QVBoxLayout(sidebar)
        nav_layout.setContentsMargins(self._s(6), self._s(10), self._s(6), self._s(10))
        nav_layout.setSpacing(self._s(4))

        self.sidebar_toggle_button = QPushButton("<")
        self.sidebar_toggle_button.setObjectName("sidebarToggle")
        self.sidebar_toggle_button.clicked.connect(self.toggle_sidebar)
        nav_layout.addWidget(self.sidebar_toggle_button)

        self.sidebar_content = QWidget()
        sidebar_content_layout = QVBoxLayout(self.sidebar_content)
        sidebar_content_layout.setContentsMargins(self._s(4), self._s(8), self._s(4), self._s(8))
        sidebar_content_layout.setSpacing(self._s(4))
        brand = QLabel(self.app_home_title)
        brand.setObjectName("brand")
        brand.setWordWrap(True)
        self.sidebar_brand_label = brand
        sidebar_content_layout.addWidget(brand)
        env_badge = QLabel()
        env_badge.setAlignment(Qt.AlignCenter)
        env_badge.setObjectName("envBadgeDev" if self.environment == "dev" else "envBadgeProd")
        env_badge.setToolTip(f"Ambiente atual: {self.environment}")
        self.sidebar_env_label = env_badge
        sidebar_content_layout.addWidget(env_badge)
        sidebar_content_layout.addSpacing(self._s(10))

        self.nav_buttons: list[QPushButton] = []
        self.nav_page_indices: list[int] = []
        self.nav_button_entries: list[tuple[QPushButton, str, int, str]] = []
        self.nav_group_buttons: dict[str, QPushButton] = {}
        self.nav_group_containers: dict[str, QWidget] = {}
        self.nav_page_groups: dict[int, str] = {}
        nav_groups = (
            (
                "Operacao Fiscal",
                (
                    ("Dashboard", 0),
                    ("ICMS/IPI Entradas", 1),
                    ("ICMS/IPI Saidas", 2),
                    ("PIS/COFINS Entradas", 3),
                    ("PIS/COFINS Saidas", 4),
                    ("XML", 5),
                    ("SPED x XML", 6),
                    ("Analise Entrada e Saida", 17),
                    ("Extrair chave NF-e", 16),
                ),
            ),
            (
                "Cadastros e Regras",
                (
                    ("Empresas", 9),
                    ("Fornecedores", 10),
                    ("Classificacao do Produto", 11),
                    ("Produtos", 12),
                    ("NCM", 18),
                    ("Importacao XML Cadastros", 13),
                    ("Limpeza Duplicados", 15),
                    ("Regras Dinamicas", 7),
                ),
            ),
            (
                "Arquivo e Administracao",
                (
                    ("SPEDs Arquivados", 8),
                    ("Usuarios E Permissoes", 19),
                    ("Configuracoes", 14),
                ),
            ),
        )
        for group_title, items in nav_groups:
            group_button = QPushButton(f"> {group_title}")
            group_button.setObjectName("navGroup")
            group_button.setCheckable(True)
            group_button.setToolTip(group_title)
            group_button.setChecked(group_title == "Operacao Fiscal")
            group_button.clicked.connect(lambda checked=False, current=group_title: self.toggle_nav_group(current))
            sidebar_content_layout.addWidget(group_button)
            self.nav_group_buttons[group_title] = group_button

            group_container = QWidget()
            group_container.setVisible(group_title == "Operacao Fiscal")
            group_layout = QVBoxLayout(group_container)
            group_layout.setContentsMargins(self._s(8), 0, 0, 0)
            group_layout.setSpacing(self._s(4))
            for title, index in items:
                button = QPushButton(title)
                button.setObjectName("navButton")
                button.setCheckable(True)
                button.setToolTip(title)
                button.clicked.connect(lambda _checked=False, current=index, label=title: self.show_page(current, label))
                group_layout.addWidget(button)
                self.nav_buttons.append(button)
                self.nav_page_indices.append(index)
                self.nav_button_entries.append((button, title, index, group_title))
                self.nav_page_groups[index] = group_title
            sidebar_content_layout.addWidget(group_container)
            self.nav_group_containers[group_title] = group_container
        sidebar_content_layout.addStretch()
        sidebar_scroll = QScrollArea()
        sidebar_scroll.setObjectName("navScroll")
        sidebar_scroll.setWidgetResizable(True)
        sidebar_scroll.setFrameShape(QFrame.NoFrame)
        sidebar_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sidebar_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        sidebar_scroll.setWidget(self.sidebar_content)
        nav_layout.addWidget(sidebar_scroll, 1)
        self.sidebar_scroll = sidebar_scroll
        shell.addWidget(sidebar)
        for group_title, container in self.nav_group_containers.items():
            self.set_nav_group_visible(group_title, container.isVisible())
        self.update_sidebar_mode_labels()

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(self._s(18), self._s(16), self._s(18), self._s(10))
        content_layout.setSpacing(self._s(12))

        header = QHBoxLayout()
        self.page_title = QLabel("")
        self.page_title.setObjectName("pageTitle")
        header.addWidget(self.page_title)
        header.addStretch()
        content_layout.addLayout(header)

        self.stack = QStackedWidget()
        pages = (
            self.build_dashboard_page(),
            self.build_entry_page(),
            self.build_sale_page(),
            self.build_contrib_page("Entrada"),
            self.build_contrib_page("Saida"),
            self.build_xml_page(),
            self.build_compare_page(),
            self.build_runtime_rules_page(),
            self.build_archives_page(),
            self.build_catalog_company_page(),
            self.build_catalog_supplier_page(),
            self.build_catalog_type_page(),
            self.build_catalog_product_page(),
            self.build_catalog_import_page(),
            self.build_settings_page(),
            self.build_duplicates_cleanup_page(),
            self.build_nfe_key_extract_page(),
            self.build_entry_exit_page(),
            self.build_catalog_ncm_page(),
            self.build_user_permissions_page(),
        )
        for page in pages:
            self.stack.addWidget(self.make_scrollable_page(page))
        content_layout.addWidget(self.stack, 1)
        shell.addWidget(content, 1)
        self.setCentralWidget(root)
        self.statusBar().showMessage("Pronto.")
        self.set_sidebar_collapsed(False)

    def show_page(self, index: int, title: str) -> None:
        self.statusBar().showMessage(f"Aguarde... carregando {title}.")
        self.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            QApplication.processEvents()
            self.stack.setCurrentIndex(index)
            self.page_title.setText(title)
            for button_index, button in enumerate(self.nav_buttons):
                button.setChecked(self.nav_page_indices[button_index] == index)
            active_group = self.nav_page_groups.get(index)
            if active_group:
                self.set_nav_group_visible(active_group, True)
            if title == "SPEDs Arquivados":
                self.refresh_archives()
            if title in {"Empresas", "Fornecedores", "Classificacao do Produto"}:
                self.refresh_current_catalog_page()
            if title == "Produtos":
                QTimer.singleShot(100, self.start_refresh_product_page_fast)
            if title == "Limpeza Duplicados":
                self.refresh_duplicates_cleanup_page()
            if title == "ICMS/IPI Saidas":
                self.sync_sale_page_with_entry_cache()
            if title != "Produtos":
                self.statusBar().showMessage(f"{title} carregada.")
        finally:
            QApplication.restoreOverrideCursor()
            self.setEnabled(True)

    def sync_sale_page_with_entry_cache(self) -> None:
        entry_sped = self.sped_input.text().strip() if hasattr(self, "sped_input") else ""
        entry_xml = self.xml_input.text().strip() if hasattr(self, "xml_input") else ""
        sale_sped = self.sale_sped_input.text().strip() if hasattr(self, "sale_sped_input") else ""
        sale_xml = self.sale_xml_input.text().strip() if hasattr(self, "sale_xml_input") else ""

        if entry_sped and not sale_sped:
            self.sale_sped_input.setText(entry_sped)
        if entry_xml and not sale_xml:
            self.sale_xml_input.setText(entry_xml)

        if self.sale_rows:
            self.refresh_sale_table()

    def toggle_sidebar(self) -> None:
        self.set_sidebar_collapsed(not self.sidebar_collapsed)

    def set_sidebar_collapsed(self, collapsed: bool) -> None:
        self.sidebar_collapsed = collapsed
        self.sidebar.setFixedWidth(self.sidebar_collapsed_width if collapsed else self.sidebar_expanded_width)
        self.sidebar_toggle_button.setText(">" if collapsed else "<")
        self.sidebar_toggle_button.setToolTip("Expandir menu" if collapsed else "Recolher menu")
        self.update_sidebar_mode_labels()

    def toggle_nav_group(self, group_title: str) -> None:
        container = self.nav_group_containers.get(group_title)
        if container is None:
            return
        self.set_nav_group_visible(group_title, not container.isVisible())

    def set_nav_group_visible(self, group_title: str, visible: bool) -> None:
        container = self.nav_group_containers.get(group_title)
        button = self.nav_group_buttons.get(group_title)
        if container is None or button is None:
            return
        container.setVisible(visible)
        button.setChecked(visible)
        prefix = "v" if visible else ">"
        label = self.compact_nav_label(group_title) if getattr(self, "sidebar_collapsed", False) else group_title
        button.setText(f"{prefix} {label}")

    def compact_nav_label(self, title: str) -> str:
        words = [word for word in title.replace("/", " ").replace("-", " ").split() if word.lower() not in {"e", "de", "do", "da"}]
        if not words:
            return title[:3].upper()
        if len(words) == 1:
            return words[0][:3].upper()
        return "".join(word[0].upper() for word in words)[:4]

    def update_sidebar_mode_labels(self) -> None:
        collapsed = getattr(self, "sidebar_collapsed", False)
        if hasattr(self, "sidebar_brand_label"):
            self.sidebar_brand_label.setVisible(not collapsed)
        if hasattr(self, "sidebar_env_label"):
            self.sidebar_env_label.setText(self.environment.upper() if collapsed else f"AMBIENTE {self.environment.upper()}")
        for group_title, container in getattr(self, "nav_group_containers", {}).items():
            self.set_nav_group_visible(group_title, container.isVisible())
        for button, title, _index, _group_title in getattr(self, "nav_button_entries", []):
            button.setText(self.compact_nav_label(title) if collapsed else title)
            button.setToolTip(title)

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

    def build_catalog_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Cadastro da Empresa")
        title.setObjectName("sectionTitle")
        self.catalog_status = QLabel("Atualize o cadastro para carregar empresas, fornecedores e produtos.")
        self.catalog_status.setObjectName("muted")
        toolbar_layout.addWidget(title)
        toolbar_layout.addWidget(self.catalog_status, 1)
        toolbar_layout.addWidget(self.create_button("Atualizar", lambda: self.run_guarded_ui_operation("refresh_catalog", "Atualizando cadastro...", self.refresh_catalog), primary=True))
        layout.addWidget(toolbar)

        self.catalog_company_fields = self.create_line_fields(("id", "nome", "cnpj", "inscricao_estadual", "observacao"))
        self.catalog_supplier_fields = self.create_line_fields(("id", "nome", "cnpj", "inscricao_estadual", "uf", "codigo", "regime_tributario", "observacao"))
        self.catalog_type_fields = self.create_line_fields(("id", "nome", "descricao"))
        self.catalog_product_fields = self.create_line_fields(
            (
                "id",
                "chave_nfe_origem",
                "codigo_fornecedor",
                "codigo_empresa",
                "descricao",
                "ean",
                "ncm",
                "cest",
                "origem_entrada",
                "cfop_saida_fornecedor",
                "cfop_entrada",
                "cfop_saida",
                "natureza_receita_entrada",
                "c_classtrib",
                "c_benef",
                "origem_saida",
                "cst_icms_saida",
                "cfop_saida_empresa",
                "aliquota_icms_saida",
                "cst_icms",
                "reducao_bc_icms",
                "aliquota_icms",
                "cst_ipi",
                "aliquota_ipi",
                "cst_pis_cofins",
                "cst_pis",
                "cst_pis_saida",
                "cst_cofins",
                "cst_cofins_saida",
                "natureza_receita_saida",
                "aliquota_pis",
                "aliquota_cofins",
                "bc_st",
                "mva",
                "valor_icms_st",
                "aliquota_icms_st",
            )
        )
        self.catalog_companies_by_id: dict[int, dict[str, object]] = {}
        self.catalog_suppliers_by_id: dict[int, dict[str, object]] = {}
        self.catalog_types_by_id: dict[int, dict[str, object]] = {}
        self.catalog_products_by_id: dict[int, dict[str, object]] = {}

        grid = QGridLayout()
        grid.setSpacing(12)
        layout.addLayout(grid, 1)

        self.catalog_company_table = self.create_catalog_section(
            grid,
            0,
            0,
            "Empresas",
            ["ID", "Empresa", "CNPJ", "IE"],
            self.catalog_company_fields,
            [("nome", "Empresa"), ("cnpj", "CNPJ"), ("inscricao_estadual", "IE"), ("observacao", "Observacao")],
            self.save_catalog_company,
            self.clear_catalog_company_form,
            self.delete_catalog_company,
        )
        self.catalog_company_table.itemSelectionChanged.connect(self.handle_catalog_company_select)

        self.catalog_supplier_table = self.create_catalog_section(
            grid,
            0,
            1,
            "Fornecedores da Empresa",
            ["ID", "Fornecedor", "CNPJ", "IE", "UF", "Codigo", "Regime"],
            self.catalog_supplier_fields,
            [("nome", "Fornecedor"), ("cnpj", "CNPJ"), ("inscricao_estadual", "IE"), ("uf", "UF"), ("codigo", "Codigo"), ("regime_tributario", "Regime Tributario"), ("observacao", "Observacao")],
            self.save_catalog_supplier,
            self.clear_catalog_supplier_form,
            self.delete_catalog_supplier,
        )
        self.catalog_supplier_table.itemSelectionChanged.connect(self.handle_catalog_supplier_select)

        self.catalog_type_table = self.create_catalog_section(
            grid,
            1,
            0,
            "Classificacao do Produto",
            ["ID", "Classificacao", "Descricao"],
            self.catalog_type_fields,
            [("nome", "Classificacao"), ("descricao", "Descricao")],
            self.save_catalog_type,
            self.clear_catalog_type_form,
            self.delete_catalog_type,
        )
        self.catalog_type_table.itemSelectionChanged.connect(self.handle_catalog_type_select)

        product_panel = self.create_panel()
        product_layout = QVBoxLayout(product_panel)
        product_layout.setContentsMargins(12, 12, 12, 12)
        product_layout.setSpacing(8)
        product_title = QLabel("Produtos do Fornecedor")
        product_title.setObjectName("sectionTitle")
        product_layout.addWidget(product_title)
        self.catalog_product_table = self.create_data_table(["ID", "Cod. Forn.", "Cod. Empresa", "Descricao", "Classificacao", "NCM", "Origem (entrada)", "CST ICMS (entrada)", "% Red BC ICMS", "CFOP saida fornecedor", "% ICMS (entrada)", "CFOP entrada empresa", "CST IPI", "% IPI", "CST PIS (entrada)", "CST PIS_COFINS (ENTRADA EMPRESA)", "% PIS", "CST COFINS (entrada)", "% COFINS", "Natureza da receita", "MVA", "Valor ICMS-ST", "cClassTrib", "cBenef", "Origem (saida)", "CST ICMS (saida)", "CFOP saida empresa", "% ICMS (saida)", "CST PIS (saida)", "CST COFINS (saida)", "Natureza da receita"])
        self.catalog_product_table.setColumnWidth(0, 55)
        self.catalog_product_table.setColumnWidth(1, 90)
        self.catalog_product_table.setColumnWidth(2, 100)
        self.catalog_product_table.setColumnWidth(3, 260)
        self.catalog_product_table.itemSelectionChanged.connect(self.handle_catalog_product_select)
        product_layout.addWidget(self.catalog_product_table, 1)

        product_form = QGridLayout()
        product_form.setHorizontalSpacing(8)
        product_form.setVerticalSpacing(6)
        self.catalog_product_type_combo = QComboBox()
        product_form.addWidget(QLabel("Classificacao do Produto"), 0, 0)
        product_form.addWidget(self.catalog_product_type_combo, 0, 1, 1, 3)
        grouped_product_labels = [
            ("Dados Gerais", [("codigo_fornecedor", "Cod. Produto Fornecedor"), ("codigo_empresa", "Cod. Produto Empresa"), ("descricao", "Descricao"), ("ean", "EAN"), ("ncm", "NCM"), ("cest", "CEST")]),
            ("Tributacao Entrada", [("origem_entrada", "Origem (entrada)"), ("cst_icms", "CST ICMS (entrada)"), ("reducao_bc_icms", "% Red BC ICMS"), ("cfop_saida_fornecedor", "CFOP saida fornecedor"), ("aliquota_icms", "% ICMS (entrada)"), ("cfop_entrada", "CFOP entrada empresa"), ("cst_ipi", "CST IPI"), ("aliquota_ipi", "% IPI"), ("cst_pis", "CST PIS (entrada)"), ("cst_pis_cofins", "CST PIS_COFINS (ENTRADA EMPRESA)"), ("aliquota_pis", "% PIS"), ("cst_cofins", "CST COFINS (entrada)"), ("aliquota_cofins", "% COFINS"), ("natureza_receita_entrada", "Natureza da receita"), ("mva", "MVA"), ("valor_icms_st", "Valor ICMS-ST"), ("c_classtrib", "cClassTrib"), ("c_benef", "cBenef")]),
            ("Dados de Saida", [("origem_saida", "Origem (saida)"), ("cst_icms_saida", "CST ICMS (saida)"), ("cfop_saida_empresa", "CFOP saida empresa"), ("aliquota_icms_saida", "% ICMS (saida)"), ("cst_pis_saida", "CST PIS (saida)"), ("cst_cofins_saida", "CST COFINS (saida)"), ("natureza_receita_saida", "Natureza da receita")]),
        ]
        current_row = 1
        for section_title, section_fields in grouped_product_labels:
            if section_title == "Dados de Saida":
                divider = QFrame()
                divider.setFrameShape(QFrame.HLine)
                divider.setFrameShadow(QFrame.Sunken)
                product_form.addWidget(divider, current_row, 0, 1, 4)
                current_row += 1
            section_label = QLabel(section_title)
            section_label.setObjectName("sectionTitle")
            product_form.addWidget(section_label, current_row, 0, 1, 4)
            current_row += 1
            for index, (key, label) in enumerate(section_fields):
                row = current_row + (index // 2)
                column = 0 if index % 2 == 0 else 2
                product_form.addWidget(QLabel(label), row, column)
                product_form.addWidget(self.catalog_product_fields[key], row, column + 1)
            current_row += (len(section_fields) + 1) // 2 + 1
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(self.create_button("Salvar Produto", self.save_catalog_product, primary=True))
        actions.addWidget(self.create_button("Novo", self.clear_catalog_product_form))
        actions.addWidget(self.create_button("Excluir", self.delete_catalog_product))
        product_form.addLayout(actions, current_row, 0, 1, 4)
        product_layout.addLayout(product_form)
        grid.addWidget(product_panel, 1, 1)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setRowStretch(0, 1)
        grid.setRowStretch(1, 1)
        return page

    def build_catalog_company_page(self) -> QWidget:
        page = self.create_single_catalog_page("Empresas")
        self.company_page_fields = self.create_line_fields(("id", "nome", "cnpj", "inscricao_estadual", "observacao"))
        self.company_page_table = self.create_data_table(["ID", "Empresa", "CNPJ", "IE"])
        self.company_page_table.setObjectName("empresas")
        self.add_single_catalog_form(
            page,
            self.company_page_table,
            self.company_page_fields,
            [("nome", "Empresa"), ("cnpj", "CNPJ"), ("inscricao_estadual", "IE"), ("observacao", "Observacao")],
            self.save_company_page,
            self.clear_company_page_form,
            self.delete_company_page,
        )
        self.company_page_table.itemSelectionChanged.connect(self.handle_company_page_select)
        return page

    def build_catalog_supplier_page(self) -> QWidget:
        page = self.create_single_catalog_page("Fornecedores")
        self.supplier_page_fields = self.create_line_fields(("id", "nome", "cnpj", "inscricao_estadual", "uf", "codigo", "regime_tributario", "observacao"))
        self.supplier_page_company_combo = QComboBox()
        self.supplier_page_regime_combo = QComboBox()
        self.populate_regime_tributario_combo(self.supplier_page_regime_combo)
        self.supplier_page_table = self.create_data_table(["ID", "Empresa", "Fornecedor", "CNPJ", "IE", "UF", "Codigo", "Regime"])
        self.supplier_page_table.setObjectName("fornecedores")
        self.add_single_catalog_form(
            page,
            self.supplier_page_table,
            self.supplier_page_fields,
            [("nome", "Fornecedor"), ("cnpj", "CNPJ"), ("inscricao_estadual", "IE"), ("uf", "UF"), ("codigo", "Codigo"), ("observacao", "Observacao")],
            self.save_supplier_page,
            self.clear_supplier_page_form,
            self.delete_supplier_page,
            [("Empresa", self.supplier_page_company_combo), ("Regime Tributario", self.supplier_page_regime_combo)],
        )
        self.supplier_page_table.itemSelectionChanged.connect(self.handle_supplier_page_select)
        return page

    def build_catalog_type_page(self) -> QWidget:
        page = self.create_single_catalog_page("Classificacao do Produto")
        self.type_page_fields = self.create_line_fields(("id", "nome", "descricao"))
        self.type_page_table = self.create_data_table(["ID", "Classificacao", "Descricao"])
        self.type_page_table.setObjectName("classificacao_produto")
        self.add_single_catalog_form(
            page,
            self.type_page_table,
            self.type_page_fields,
            [("nome", "Classificacao"), ("descricao", "Descricao")],
            self.save_type_page,
            self.clear_type_page_form,
            self.delete_type_page,
        )
        self.type_page_table.itemSelectionChanged.connect(self.handle_type_page_select)
        return page

    def build_catalog_ncm_page(self) -> QWidget:
        page = self.create_single_catalog_page("NCM")
        import_panel = self.create_panel()
        import_layout = QHBoxLayout(import_panel)
        import_layout.setContentsMargins(12, 10, 12, 10)
        self.ncm_import_path_input = QLineEdit()
        self.ncm_import_path_input.setPlaceholderText("Selecione a planilha de regras fiscais por NCM")
        import_layout.addWidget(QLabel("Planilha"))
        import_layout.addWidget(self.ncm_import_path_input, 1)
        import_layout.addWidget(self.create_button("Selecionar", self.select_ncm_import_file))
        import_layout.addWidget(self.create_button("Importar Planilha", self.import_ncm_rules_from_excel, primary=True))
        page.layout().addWidget(import_panel)

        ncm_keys = (
            "id", "atividade", "regime_tributario", "uf", "data_vigencia", "ncm", "descricao", "cest",
            "aliquota_ipi", "cst_ipi", "ex_tipi", "cst_pis_cofins_entrada", "cst_pis_cofins_saida",
            "codigo_sped", "aliquota_pis", "aliquota_cofins", "base_legal_pis_cofins", "cfop_entrada", "cfop_saida",
            "cst_csosn", "ad_rem_icms", "aliquota_icms", "reducao_bc_icms", "reducao_bc_icms_st",
            "aliquota_icms_st", "aliquota_red_bc_icms", "mva", "fcp", "codigo_beneficio_fiscal",
            "antecipado", "percentual_diferimento", "percentual_isencao", "codigo_anp", "base_legal_icms",
        )
        self.ncm_fields = self.create_line_fields(ncm_keys)
        self.ncm_table = self.create_data_table([
            "ID", "Atividade", "Regime", "UF", "Vigencia", "NCM", "Descricao", "CEST", "% IPI",
            "CST IPI", "EX", "CST PIS/COFINS Entrada", "CST PIS/COFINS Saida", "Codigo SPED",
            "% PIS", "% COFINS", "Base Legal PIS/COFINS", "CFOP Entrada", "CFOP Saida", "CST/CSOSN", "AD REM ICMS",
            "% ICMS", "% Red BC ICMS", "% Red BC ICMS ST", "% ICMS ST", "% Aliq Red BC ICMS",
            "MVA", "FCP", "Cod. Beneficio Fiscal", "Antecipado", "% Diferimento", "% Isencao",
            "Codigo ANP", "Base Legal ICMS",
        ])
        self.ncm_table.setObjectName("ncm")
        ncm_form_fields = [
            ("atividade", "Atividade"), ("regime_tributario", "Regime tributario"), ("uf", "UF"),
            ("data_vigencia", "Data vigencia"), ("ncm", "NCM"), ("descricao", "Descricao"),
            ("cest", "CEST"), ("aliquota_ipi", "% IPI"), ("cst_ipi", "CST IPI"), ("ex_tipi", "EX"),
            ("cst_pis_cofins_entrada", "CST PIS/COFINS entrada"), ("cst_pis_cofins_saida", "CST PIS/COFINS saida"),
            ("codigo_sped", "Codigo SPED"), ("aliquota_pis", "% PIS"), ("aliquota_cofins", "% COFINS"),
            ("base_legal_pis_cofins", "Base legal PIS/COFINS"), ("cfop_entrada", "CFOP entrada"), ("cfop_saida", "CFOP saida"), ("cst_csosn", "CST/CSOSN"),
            ("ad_rem_icms", "AD REM ICMS"), ("aliquota_icms", "% ICMS"), ("reducao_bc_icms", "% Red BC ICMS"),
            ("reducao_bc_icms_st", "% Red BC ICMS ST"), ("aliquota_icms_st", "% ICMS ST"),
            ("aliquota_red_bc_icms", "% Aliq Red BC ICMS"), ("mva", "MVA"), ("fcp", "FCP"),
            ("codigo_beneficio_fiscal", "Cod. Beneficio Fiscal"), ("antecipado", "Antecipado"),
            ("percentual_diferimento", "% Diferimento"), ("percentual_isencao", "% Isencao"),
            ("codigo_anp", "Codigo ANP"), ("base_legal_icms", "Base legal ICMS"),
        ]
        self.add_single_catalog_form(page, self.ncm_table, self.ncm_fields, ncm_form_fields, self.save_ncm_rule, self.clear_ncm_form, self.delete_ncm_rule)
        self.ncm_table.itemSelectionChanged.connect(self.handle_ncm_select)
        return page

    def build_catalog_product_page(self) -> QWidget:
        page = self.create_single_catalog_page("Produtos")
        header_panel = page.layout().itemAt(0).widget()
        if header_panel is not None and header_panel.layout() is not None:
            header_panel.layout().addWidget(
                self.create_button("Importar Planilha", self.open_planilha_import)
            )
        self.product_page_fields = self.create_line_fields(
            (
                "id",
                "chave_nfe_origem",
                "codigo_fornecedor",
                "codigo_empresa",
                "descricao",
                "ean",
                "ncm",
                "cest",
                "origem_entrada",
                "cfop_saida_fornecedor",
                "cfop_entrada",
                "cfop_saida",
                "natureza_receita_entrada",
                "c_classtrib",
                "c_benef",
                "origem_saida",
                "cst_icms_saida",
                "cfop_saida_empresa",
                "aliquota_icms_saida",
                "cst_icms",
                "reducao_bc_icms",
                "aliquota_icms",
                "cst_ipi",
                "aliquota_ipi",
                "cst_pis_cofins",
                "cst_pis",
                "cst_pis_saida",
                "cst_cofins",
                "cst_cofins_saida",
                "natureza_receita_saida",
                "aliquota_pis",
                "aliquota_cofins",
                "bc_st",
                "mva",
                "valor_icms_st",
                "aliquota_icms_st",
            )
        )
        self.product_page_supplier_combo = QComboBox()
        self.product_page_type_combo = QComboBox()
        self.product_page_filter_company_combo = QComboBox()
        self.product_page_filter_supplier_combo = QComboBox()
        for field in self.product_page_fields.values():
            self.make_product_input_compact(field)
        self.make_product_input_compact(self.product_page_supplier_combo)
        self.make_product_input_compact(self.product_page_type_combo)
        self.product_page_full_headers = [
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
        self.product_page_table = self.create_product_table_view(self.product_page_full_headers[:12])
        self.product_page_table.setObjectName("produtos")
        self.product_page_tabs = QTabWidget()

        consult_tab = QWidget()
        consult_layout = QVBoxLayout(consult_tab)
        consult_layout.setContentsMargins(10, 10, 10, 10)
        consult_layout.setSpacing(8)
        search_grid = QGridLayout()
        search_grid.setHorizontalSpacing(8)
        search_grid.setVerticalSpacing(6)
        self.product_page_filter_company_combo.addItem("Todas Empresas", 0)
        self.product_page_filter_supplier_combo.addItem("Todos Fornecedores", 0)
        self.product_page_filter_company_combo.setMinimumWidth(260)
        self.product_page_filter_supplier_combo.setMinimumWidth(320)
        self.product_page_filter_company_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.product_page_filter_supplier_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.configure_searchable_combo(self.product_page_filter_company_combo, "Digite para localizar empresa")
        self.configure_searchable_combo(self.product_page_filter_supplier_combo, "Digite para localizar fornecedor")
        self.product_page_filter_company_combo.currentIndexChanged.connect(self.on_product_filter_company_changed)
        self.product_page_filter_supplier_combo.currentIndexChanged.connect(lambda _idx: self.apply_product_page_filters())
        search_grid.addWidget(QLabel("Empresa"), 0, 0)
        search_grid.addWidget(QLabel("Fornecedor"), 0, 1)
        search_grid.addWidget(QLabel("Produto"), 0, 2)
        search_grid.addWidget(self.product_page_filter_company_combo, 1, 0)
        search_grid.addWidget(self.product_page_filter_supplier_combo, 1, 1)
        self.product_page_search_input = QLineEdit()
        self.product_page_search_input.setPlaceholderText("Codigo, descricao, classificacao, EAN, NCM ou CST")
        self.product_page_search_input.textChanged.connect(lambda _text: self.apply_product_page_filters())
        search_grid.addWidget(self.product_page_search_input, 1, 2)
        self.product_page_xml_only_check = QCheckBox("Somente via XML (tem Chave NF-e)")
        self.product_page_xml_only_check.stateChanged.connect(lambda _: self.apply_product_page_filters())
        search_grid.addWidget(self.product_page_xml_only_check, 2, 0, 1, 2)
        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar", self.export_product_page_filtered))
        actions.addWidget(self.create_button("Limpar Consulta", self.clear_product_page_consultation))
        self.product_page_refresh_button = self.create_button(
            "Atualizar",
            self.start_refresh_product_page_fast,
            primary=True,
        )
        actions.addWidget(self.product_page_refresh_button)
        actions.addWidget(self.create_button("Editar Selecionado", self.edit_product_page_selected))
        actions.addWidget(self.create_button("Excluir", self.delete_product_page))
        actions.addStretch()
        search_grid.addLayout(actions, 3, 0, 1, 3)
        search_grid.setColumnStretch(0, 2)
        search_grid.setColumnStretch(1, 2)
        search_grid.setColumnStretch(2, 5)
        consult_layout.addLayout(search_grid)
        product_metrics_layout = QGridLayout()
        product_metrics_layout.setHorizontalSpacing(8)
        self.product_metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate(
            (
                ("products", "Produtos"),
                ("classes", "Classificacoes"),
                ("top_class", "Classificacao Lider"),
                ("scope", "Escopo"),
            )
        ):
            card, value_label = self.create_metric_card_with_label(label, "-")
            self.product_metric_labels[key] = value_label
            product_metrics_layout.addWidget(card, 0, column)
        consult_layout.addLayout(product_metrics_layout)
        self.product_page_table.doubleClicked.connect(lambda _index: self.edit_product_page_selected())
        consult_layout.addWidget(self.product_page_table, 1)
        self.product_page_tabs.addTab(consult_tab, "Consulta")

        form_tab = QWidget()
        form_tab.setObjectName("productFormTab")
        form_layout = QVBoxLayout(form_tab)
        form_layout.setContentsMargins(8, 6, 8, 6)
        form_layout.setSpacing(5)
        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(4)
        form.addWidget(QLabel("Fornecedor"), 0, 0)
        form.addWidget(self.product_page_supplier_combo, 0, 1, 1, 3)
        form.addWidget(QLabel("Classificacao do Produto"), 1, 0)
        form.addWidget(self.product_page_type_combo, 1, 1, 1, 3)
        gerais_fields = [
            ("chave_nfe_origem", "Chave NFe origem"),
            ("codigo_fornecedor", "Cod. Produto Fornecedor"),
            ("codigo_empresa", "Cod. Produto Empresa"),
            ("descricao", "Descricao"),
            ("ean", "EAN"),
            ("ncm", "NCM"),
            ("cest", "CEST"),
            ("c_classtrib", "cClassTrib"),
            ("c_benef", "cBenef"),
        ]
        importacao_xml_fields = [
            ("origem_entrada", "Origem (entrada)"),
            ("cst_icms", "CST ICMS (entrada)"),
            ("reducao_bc_icms", "% Red BC ICMS"),
            ("cfop_saida_fornecedor", "CFOP saida fornecedor"),
            ("aliquota_icms", "% ICMS (entrada)"),
            ("cst_ipi", "CST IPI"),
            ("aliquota_ipi", "% IPI"),
            ("cst_pis", "CST PIS (entrada)"),
            ("aliquota_pis", "% PIS"),
            ("cst_cofins", "CST COFINS (entrada)"),
            ("aliquota_cofins", "% COFINS"),
            ("mva", "MVA"),
            ("valor_icms_st", "Valor ICMS-ST"),
        ]
        entrada_empresa_fields = [
            ("cfop_entrada", "CFOP entrada empresa"),
            ("cst_pis_cofins", "CST PIS_COFINS entrada empresa"),
            ("natureza_receita_entrada", "Natureza da receita"),
        ]
        saida_fields = [
            ("origem_saida", "Origem (saida)"),
            ("cst_icms_saida", "CST ICMS (saida)"),
            ("cfop_saida_empresa", "CFOP saida empresa"),
            ("aliquota_icms_saida", "% ICMS (saida)"),
            ("cst_pis_saida", "CST PIS (saida)"),
            ("cst_cofins_saida", "CST COFINS (saida)"),
            ("natureza_receita_saida", "Natureza da receita"),
        ]
        current_row = 2
        gerais_label = QLabel("Dados Gerais")
        gerais_label.setObjectName("sectionTitle")
        form.addWidget(gerais_label, current_row, 0, 1, 4)
        current_row += 1
        for index, (key, label) in enumerate(gerais_fields):
            row = current_row + (index // 2)
            column = 0 if index % 2 == 0 else 2
            form.addWidget(QLabel(label), row, column)
            form.addWidget(self.product_page_fields[key], row, column + 1)
        current_row += (len(gerais_fields) + 1) // 2 + 1

        trib_tabs = QTabWidget()
        trib_tabs.setObjectName("productTribTabs")
        importacao_xml_tab = QWidget()
        importacao_xml_grid = QGridLayout(importacao_xml_tab)
        importacao_xml_grid.setContentsMargins(8, 6, 8, 6)
        importacao_xml_grid.setHorizontalSpacing(8)
        importacao_xml_grid.setVerticalSpacing(4)
        for index, (key, label) in enumerate(importacao_xml_fields):
            row = index // 2
            column = 0 if index % 2 == 0 else 2
            importacao_xml_grid.addWidget(QLabel(label), row, column)
            importacao_xml_grid.addWidget(self.product_page_fields[key], row, column + 1)

        entrada_empresa_tab = QWidget()
        entrada_empresa_grid = QGridLayout(entrada_empresa_tab)
        entrada_empresa_grid.setContentsMargins(8, 6, 8, 6)
        entrada_empresa_grid.setHorizontalSpacing(8)
        entrada_empresa_grid.setVerticalSpacing(4)
        for index, (key, label) in enumerate(entrada_empresa_fields):
            row = index // 2
            column = 0 if index % 2 == 0 else 2
            entrada_empresa_grid.addWidget(QLabel(label), row, column)
            entrada_empresa_grid.addWidget(self.product_page_fields[key], row, column + 1)

        saida_tab = QWidget()
        saida_grid = QGridLayout(saida_tab)
        saida_grid.setContentsMargins(8, 6, 8, 6)
        saida_grid.setHorizontalSpacing(8)
        saida_grid.setVerticalSpacing(4)
        for index, (key, label) in enumerate(saida_fields):
            row = index // 2
            column = 0 if index % 2 == 0 else 2
            saida_grid.addWidget(QLabel(label), row, column)
            saida_grid.addWidget(self.product_page_fields[key], row, column + 1)

        trib_tabs.addTab(importacao_xml_tab, "Importacao XML")
        trib_tabs.addTab(entrada_empresa_tab, "Entrada Empresa")
        trib_tabs.addTab(saida_tab, "Saida Empresa")
        form.addWidget(trib_tabs, current_row, 0, 1, 4)
        current_row += 1
        self.product_review_import_status = QLabel("Importe a planilha revisada selecionando a empresa e localizando por Fornecedor e Cod. Forn.")
        self.product_review_import_status.setObjectName("muted")
        form.addWidget(self.product_review_import_status, current_row, 0, 1, 4)
        current_row += 1
        self.product_review_import_progress = QProgressBar()
        self.product_review_import_progress.setMinimum(0)
        self.product_review_import_progress.setMaximum(100)
        self.product_review_import_progress.setValue(0)
        form.addWidget(self.product_review_import_progress, current_row, 0, 1, 4)
        current_row += 1
        actions = QHBoxLayout()
        actions.addStretch()
        self.product_review_import_button = self.create_button("Importar Revisao", self.import_product_page_review_excel)
        actions.addWidget(self.product_review_import_button)
        actions.addWidget(self.create_button("Salvar Alteracoes", self.save_product_page, primary=True))
        actions.addWidget(self.create_button("Novo", self.clear_product_page_form))
        actions.addWidget(self.create_button("Excluir", self.delete_product_page))
        form.addLayout(actions, current_row, 0, 1, 4)
        form_layout.addLayout(form)
        form_layout.addStretch()
        self.product_page_tabs.addTab(form_tab, "Cadastro / Edicao")
        self.product_page_tabs.addTab(self._build_sped_catalog_check_tab(), "Conferência SPED x Cadastro")
        page.layout().addWidget(self.product_page_tabs, 1)
        self.product_page_table.selectionModel().selectionChanged.connect(lambda _selected, _deselected: self.handle_product_page_select())
        return page

    # ------------------------------------------------------------------
    # Conferência SPED x Cadastro
    # ------------------------------------------------------------------
    def _build_sped_catalog_check_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        title = QLabel("Conferência: Produtos C170 do SPED x Cadastro de Produtos")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("Arquivo SPED:"))
        self.sped_check_file_input = QLineEdit()
        self.sped_check_file_input.setPlaceholderText("Selecione um arquivo SPED (.txt / .efd)...")
        self.sped_check_file_input.setReadOnly(True)
        file_row.addWidget(self.sped_check_file_input, 1)
        file_row.addWidget(self.create_button("Selecionar", self._select_sped_check_file))
        file_row.addWidget(self.create_button("Gerar Relatorio", self.run_sped_catalog_check, primary=True))
        file_row.addWidget(self.create_button("Exportar Excel", self.export_sped_catalog_check))
        layout.addLayout(file_row)

        self.sped_check_info_label = QLabel("")
        self.sped_check_info_label.setObjectName("statusLabel")
        layout.addWidget(self.sped_check_info_label)

        metrics_row = QHBoxLayout()
        self.sped_check_metric_labels: dict[str, QLabel] = {}
        for key, label in (
            ("total", "Produtos C170"),
            ("found", "Cadastrados"),
            ("missing", "Nao Cadastrados"),
            ("via_xml", "Via XML"),
        ):
            card, value_label = self.create_metric_card_with_label(label, "-")
            self.sped_check_metric_labels[key] = value_label
            metrics_row.addWidget(card)
        metrics_row.addStretch()
        layout.addLayout(metrics_row)

        headers = [
            "Status",
            "Cod. SPED",
            "Descricao SPED",
            "NCM SPED",
            "CST ICMS SPED",
            "% ICMS SPED",
            "CEST SPED",
            "Fornecedor",
            "Cod. Fornecedor",
            "Cod. Empresa",
            "Descricao Cadastro",
            "NCM Cadastro",
            "CST ICMS Cad.",
            "% ICMS Cad.",
            "CST PIS Cad.",
            "CST COFINS Cad.",
            "Via XML",
            "Fornecedor Entrada (SPED)",
            "CNPJ Fornecedor Entrada",
        ]
        self.sped_check_headers = headers
        self.sped_check_rows: list[list[object]] = []
        self.sped_check_table = self.create_data_table(headers)
        layout.addWidget(self.sped_check_table, 1)
        return tab

    def _select_sped_check_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo SPED", "",
            "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)",
        )
        if file_path:
            self.sped_check_file_input.setText(file_path)
            self.sped_check_info_label.setText("Arquivo selecionado. Clique em 'Gerar Relatorio'.")

    def reload_sped_catalog_check_archives(self) -> None:
        pass  # not used in file-based approach

    def run_sped_catalog_check(self) -> None:
        from app.parsers.sped_fiscal_parser import read_sped_0200_products

        file_path_str = self.sped_check_file_input.text().strip()
        if not file_path_str:
            QMessageBox.warning(self, "Conferência SPED x Cadastro", "Selecione um arquivo SPED antes de gerar o relatório.")
            return
        sped_path = Path(file_path_str)
        if not sped_path.exists():
            QMessageBox.warning(self, "Conferência SPED x Cadastro", "Arquivo SPED não encontrado no caminho informado.")
            return
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            company_cnpj, company_name, periodo_ini, periodo_fim, sped_products, entry_suppliers, c170_codes = read_sped_0200_products(sped_path)
            if not company_cnpj:
                QApplication.restoreOverrideCursor()
                QMessageBox.warning(self, "Conferência SPED x Cadastro", "Não foi possível identificar o CNPJ da empresa no arquivo SPED (registro 0000 não encontrado).")
                return
            catalog_rows = self.mysql_repo.get_catalog_products_by_company_cnpj(self.environment, company_cnpj)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Conferência SPED x Cadastro", f"Erro ao processar:\n{exc}")
            return
        finally:
            QApplication.restoreOverrideCursor()

        # Index catalog products by codigo_empresa
        catalog_by_code: dict[str, list[dict]] = {}
        for prod in catalog_rows:
            code = str(prod.get("codigo_empresa") or "").strip()
            if code:
                catalog_by_code.setdefault(code, []).append(prod)

        table_rows: list[list[object]] = []
        seen_sped_codes: set[str] = set()
        for sped_prod in sped_products:
            sped_code = str(sped_prod.get("codigo") or "").strip()
            if sped_code in seen_sped_codes:
                continue
            if c170_codes and sped_code not in c170_codes:
                continue
            seen_sped_codes.add(sped_code)
            matches = catalog_by_code.get(sped_code, [])
            if matches:
                for cad in matches:
                    chave = str(cad.get("chave_nfe_origem") or "").strip()
                    table_rows.append([
                        "Cadastrado",
                        sped_code,
                        sped_prod.get("descricao", ""),
                        sped_prod.get("ncm", ""),
                        sped_prod.get("cst_icms", ""),
                        str(sped_prod.get("aliquota_icms") or ""),
                        sped_prod.get("cest", ""),
                        str(cad.get("fornecedor_nome") or cad.get("fornecedor") or ""),
                        str(cad.get("codigo_fornecedor") or ""),
                        str(cad.get("codigo_empresa") or ""),
                        str(cad.get("descricao") or ""),
                        str(cad.get("ncm") or ""),
                        str(cad.get("cst_icms") or ""),
                        str(cad.get("aliquota_icms") or ""),
                        str(cad.get("cst_pis") or ""),
                        str(cad.get("cst_cofins") or ""),
                        "Sim" if chave else "Nao",
                        "",
                        "",
                    ])
            else:
                suppliers = entry_suppliers.get(sped_code, [])
                sup_names = " ; ".join(s.get("nome", "") for s in suppliers if s.get("nome"))
                sup_cnpjs = " ; ".join(s.get("cnpj", "") for s in suppliers if s.get("cnpj"))
                table_rows.append([
                    "Nao Cadastrado",
                    sped_code,
                    sped_prod.get("descricao", ""),
                    sped_prod.get("ncm", ""),
                    sped_prod.get("cst_icms", ""),
                    str(sped_prod.get("aliquota_icms") or ""),
                    sped_prod.get("cest", ""),
                    "", "", "", "", "", "", "", "", "",
                    "",
                    sup_names,
                    sup_cnpjs,
                ])

        self.sped_check_rows = table_rows
        self.set_table_rows(self.sped_check_table, table_rows)

        total = len(seen_sped_codes)
        found = len({r[1] for r in table_rows if r[0] == "Cadastrado"})
        missing = total - found
        via_xml = len({r[1] for r in table_rows if r[16] == "Sim"})
        for key, value in (("total", total), ("found", found), ("missing", missing), ("via_xml", via_xml)):
            lbl = self.sped_check_metric_labels.get(key)
            if lbl:
                lbl.setText(str(value))

        periodo = ""
        if periodo_ini and periodo_fim:
            try:
                def fmt(d: str) -> str:
                    return f"{d[:2]}/{d[2:4]}/{d[4:]}" if len(d) == 8 else d
                periodo = f"{fmt(periodo_ini)} a {fmt(periodo_fim)}"
            except Exception:
                periodo = f"{periodo_ini} - {periodo_fim}"
        self.sped_check_info_label.setText(
            f"Empresa: {company_name}  |  CNPJ: {company_cnpj}  |  Periodo: {periodo}  |  {total} produtos com C170"
        )

    def export_sped_catalog_check(self) -> None:
        rows = getattr(self, "sped_check_rows", [])
        if not rows:
            QMessageBox.warning(self, "Exportar", "Gere o relatorio antes de exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self, "Salvar relatorio", "conferencia_sped_cadastro.xlsx",
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        output_path = Path(output)
        headers = list(getattr(self, "sped_check_headers", []))
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Conferencia", headers, rows, {"include_total": False})])
            self.handle_export_success("Conferência SPED x Cadastro", output_path, "Relatório exportado")
        except Exception as exc:
            self.handle_export_failure("Conferência SPED x Cadastro", "conferencia", exc)

    def create_single_catalog_page(self, title: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        header = self.create_panel()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 10, 12, 10)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.create_button("Consultar/Atualizar", lambda: self.refresh_current_catalog_page(), primary=True))
        layout.addWidget(header)
        return page

    def add_single_catalog_form(
        self,
        page: QWidget,
        table: QTableWidget,
        fields: dict[str, QLineEdit],
        form_fields: list[tuple[str, str]],
        save_callback: Callable[[], None],
        clear_callback: Callable[[], None],
        delete_callback: Callable[[], None],
        combos: list[tuple[str, QComboBox]] | None = None,
    ) -> None:
        layout = page.layout()
        panel = self.create_panel()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(12, 12, 12, 12)
        tabs = QTabWidget()
        if not hasattr(self, "catalog_page_tabs_by_table"):
            self.catalog_page_tabs_by_table: dict[int, QTabWidget] = {}
        self.catalog_page_tabs_by_table[id(table)] = tabs
        if not hasattr(self, "catalog_page_search_inputs_by_table"):
            self.catalog_page_search_inputs_by_table: dict[int, QLineEdit] = {}

        consult_tab = QWidget()
        consult_layout = QVBoxLayout(consult_tab)
        consult_layout.setContentsMargins(10, 10, 10, 10)
        consult_layout.setSpacing(8)
        search_row = QHBoxLayout()
        search_label = QLabel("Consultar")
        search_input = QLineEdit()
        search_input.setPlaceholderText("Pesquisar na lista")
        search_input.textChanged.connect(lambda text, current_table=table: self.filter_table_rows(current_table, text))
        self.catalog_page_search_inputs_by_table[id(table)] = search_input
        search_row.addWidget(search_label)
        search_row.addWidget(search_input, 1)
        search_row.addWidget(
            self.create_button(
                "Exportar",
                lambda current_table=table, name=(table.objectName() or "cadastros"): self.export_catalog_table_filtered(current_table, name),
            )
        )
        search_row.addWidget(self.create_button("Limpar Consulta", lambda current_table=table: self.clear_catalog_consultation(current_table)))
        search_row.addWidget(self.create_button("Editar Selecionado", lambda current_table=table: self.edit_single_catalog_selected(current_table)))
        search_row.addWidget(self.create_button("Excluir", delete_callback))
        consult_layout.addLayout(search_row)
        table.doubleClicked.connect(lambda _index, current_table=table: self.edit_single_catalog_selected(current_table))
        consult_layout.addWidget(table, 1)
        tabs.addTab(consult_tab, "Consulta")

        form_tab = QWidget()
        form_tab_layout = QVBoxLayout(form_tab)
        form_tab_layout.setContentsMargins(10, 10, 10, 10)
        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(6)
        row_index = 0
        for label, combo in combos or []:
            form.addWidget(QLabel(label), row_index, 0)
            form.addWidget(combo, row_index, 1, 1, 3)
            row_index += 1
        for index, (key, label) in enumerate(form_fields):
            row = row_index + index // 2
            column = 0 if index % 2 == 0 else 2
            form.addWidget(QLabel(label), row, column)
            form.addWidget(fields[key], row, column + 1)
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(self.create_button("Salvar Alteracoes", save_callback, primary=True))
        actions.addWidget(self.create_button("Novo", lambda current_table=table, command=clear_callback: self.clear_single_catalog_form(current_table, command)))
        actions.addWidget(self.create_button("Excluir", delete_callback))
        form.addLayout(actions, row_index + ((len(form_fields) + 1) // 2), 0, 1, 4)
        form_tab_layout.addLayout(form)
        form_tab_layout.addStretch()
        tabs.addTab(form_tab, "Cadastro / Edicao")

        panel_layout.addWidget(tabs, 1)
        layout.addWidget(panel, 1)

    def create_line_fields(self, keys: tuple[str, ...]) -> dict[str, QLineEdit]:
        fields = {key: QLineEdit() for key in keys}
        if "id" in fields:
            fields["id"].setReadOnly(True)
        if "chave_nfe_origem" in fields:
            fields["chave_nfe_origem"].setReadOnly(True)
        if "cnpj" in fields:
            fields["cnpj"].setInputMask("00.000.000/0000-00;_")
        if "inscricao_estadual" in fields:
            # IE varia por UF; aqui aplicamos mascara generica numerica.
            fields["inscricao_estadual"].setInputMask("000000000000000;_")
        if "uf" in fields:
            fields["uf"].setMaxLength(2)
        return fields

    def create_catalog_section(
        self,
        grid: QGridLayout,
        row: int,
        column: int,
        title: str,
        headers: list[str],
        fields: dict[str, QLineEdit],
        form_fields: list[tuple[str, str]],
        save_callback: Callable[[], None],
        clear_callback: Callable[[], None],
        delete_callback: Callable[[], None],
    ) -> QTableWidget:
        panel = self.create_panel()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        layout.addWidget(label)
        table = self.create_data_table(headers)
        table.setColumnWidth(0, 55)
        layout.addWidget(table, 1)
        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(6)
        for index, (key, field_label) in enumerate(form_fields):
            form.addWidget(QLabel(field_label), index, 0)
            form.addWidget(fields[key], index, 1)
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(self.create_button("Salvar", save_callback, primary=True))
        actions.addWidget(self.create_button("Novo", clear_callback))
        actions.addWidget(self.create_button("Excluir", delete_callback))
        form.addLayout(actions, len(form_fields), 0, 1, 2)
        layout.addLayout(form)
        grid.addWidget(panel, row, column)
        return table

    def build_user_permissions_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        list_panel = self.create_panel()
        list_layout = QVBoxLayout(list_panel)
        list_layout.setContentsMargins(16, 16, 16, 16)
        list_layout.setSpacing(8)
        title = QLabel("Usuarios E Permissoes")
        title.setObjectName("sectionTitle")
        list_layout.addWidget(title)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self.create_button("Atualizar", self.refresh_system_users, primary=True))
        toolbar.addWidget(self.create_button("Novo Usuario", self.clear_system_user_form))
        toolbar.addStretch()
        list_layout.addLayout(toolbar)

        self.system_user_table = self.create_data_table(["ID", "Nome", "Login", "Ativo", "Permissoes", "Observacao"])
        self.system_user_table.selectionModel().selectionChanged.connect(lambda _selected, _deselected: self.handle_system_user_select())
        list_layout.addWidget(self.system_user_table, 1)
        layout.addWidget(list_panel, 1)

        form_panel = self.create_panel()
        form_layout = QVBoxLayout(form_panel)
        form_layout.setContentsMargins(16, 16, 16, 16)
        form_layout.setSpacing(10)
        form_title = QLabel("Cadastro Do Usuario")
        form_title.setObjectName("sectionTitle")
        form_layout.addWidget(form_title)

        self.system_user_id = 0
        self.system_user_name_input = QLineEdit()
        self.system_user_login_input = QLineEdit()
        self.system_user_password_input = QLineEdit()
        self.system_user_password_input.setEchoMode(QLineEdit.Password)
        self.system_user_password_input.setPlaceholderText("Preencha para criar ou alterar a senha")
        self.system_user_active_check = QCheckBox("Usuario Ativo")
        self.system_user_active_check.setChecked(True)
        self.system_user_note_input = QTextEdit()
        self.system_user_note_input.setFixedHeight(70)

        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        form.addWidget(QLabel("Nome"), 0, 0)
        form.addWidget(self.system_user_name_input, 0, 1)
        form.addWidget(QLabel("Login"), 0, 2)
        form.addWidget(self.system_user_login_input, 0, 3)
        form.addWidget(QLabel("Senha"), 1, 0)
        form.addWidget(self.system_user_password_input, 1, 1, 1, 3)
        form.addWidget(QLabel("Situacao"), 2, 0)
        form.addWidget(self.system_user_active_check, 2, 1)
        form.addWidget(QLabel("Observacao"), 3, 0)
        form.addWidget(self.system_user_note_input, 3, 1, 1, 3)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)
        form_layout.addLayout(form)

        permissions_layout = QGridLayout()
        permissions_layout.setHorizontalSpacing(10)
        permissions_layout.setVerticalSpacing(10)
        self.system_user_permission_checks: dict[str, QCheckBox] = {}
        for group_index, (group_name, permissions) in enumerate(USER_PERMISSION_GROUPS):
            group_box = QGroupBox(group_name)
            group_inner = QVBoxLayout(group_box)
            group_inner.setContentsMargins(10, 10, 10, 10)
            group_inner.setSpacing(5)
            for permission_key, permission_label in permissions:
                check = QCheckBox(permission_label)
                self.system_user_permission_checks[permission_key] = check
                group_inner.addWidget(check)
            group_inner.addStretch()
            permissions_layout.addWidget(group_box, group_index // 3, group_index % 3)
        form_layout.addLayout(permissions_layout)

        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(self.create_button("Salvar Usuario", self.save_system_user, primary=True))
        actions.addWidget(self.create_button("Novo", self.clear_system_user_form))
        actions.addWidget(self.create_button("Excluir", self.delete_system_user))
        form_layout.addLayout(actions)
        layout.addWidget(form_panel)

        QTimer.singleShot(0, self.refresh_system_users)
        return page

    def refresh_system_users(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            rows = self.mysql_repo.list_system_users(self.environment)
            table_rows = []
            for row in rows:
                table_rows.append([
                    row.get("id", ""),
                    row.get("nome", ""),
                    row.get("login", ""),
                    "Sim" if int(row.get("ativo") or 0) else "Nao",
                    row.get("permissoes", "") or "",
                    row.get("observacao", "") or "",
                ])
            self.set_table_rows(self.system_user_table, table_rows)
            self.statusBar().showMessage(f"{len(table_rows)} usuario(s) carregado(s).")
        except Exception as exc:
            QMessageBox.critical(self, "Usuarios E Permissoes", str(exc))

    def ensure_default_admin_user(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            users = self.mysql_repo.list_system_users(self.environment)
            for user in users:
                if str(user.get("login") or "").strip().lower() == "admin":
                    return
            permissions = {
                permission_key
                for _group_name, group_permissions in USER_PERMISSION_GROUPS
                for permission_key, _permission_label in group_permissions
            }
            self.mysql_repo.save_system_user(
                self.environment,
                None,
                {
                    "nome": "Administrador Padrao",
                    "login": "admin",
                    "senha": "admin",
                    "ativo": True,
                    "observacao": "Usuario administrador criado automaticamente para primeiro acesso.",
                },
                permissions,
            )
            if hasattr(self, "system_user_table"):
                self.refresh_system_users()
            self.statusBar().showMessage("Usuario admin criado para primeiro acesso.")
        except Exception as exc:
            self.statusBar().showMessage(f"Nao foi possivel criar usuario admin: {exc}")

    def show_login_dialog(self) -> bool:
        while True:
            dialog = LoginDialog(self.app_home_title, self.environment, self)
            dialog.setStyleSheet(self.styleSheet())
            if dialog.exec() != QDialog.Accepted:
                return False
            login, password = dialog.credentials()
            try:
                user = self.mysql_repo.authenticate_system_user(self.environment, login, password)
            except Exception as exc:
                QMessageBox.critical(self, "Acesso Ao Sistema", str(exc))
                continue
            if user is None:
                QMessageBox.warning(self, "Acesso Ao Sistema", "Usuario ou senha invalidos.")
                continue
            self.current_user = user
            user_name = str(user.get("nome") or user.get("login") or "").strip()
            self.statusBar().showMessage(f"Usuario autenticado: {user_name}.")
            return True

    def selected_system_user_id(self) -> int:
        selected = self.system_user_table.selectionModel().selectedRows()
        if not selected:
            return 0
        item = self.system_user_table.item(selected[0].row(), 0)
        return int(item.text()) if item and item.text().isdigit() else 0

    def handle_system_user_select(self) -> None:
        selected = self.system_user_table.selectionModel().selectedRows()
        if not selected:
            return
        row_index = selected[0].row()
        self.system_user_id = self.selected_system_user_id()
        self.system_user_name_input.setText(self.system_user_table.item(row_index, 1).text() if self.system_user_table.item(row_index, 1) else "")
        self.system_user_login_input.setText(self.system_user_table.item(row_index, 2).text() if self.system_user_table.item(row_index, 2) else "")
        active_text = self.system_user_table.item(row_index, 3).text() if self.system_user_table.item(row_index, 3) else ""
        self.system_user_active_check.setChecked(active_text.lower() == "sim")
        self.system_user_password_input.clear()
        self.system_user_note_input.setPlainText(self.system_user_table.item(row_index, 5).text() if self.system_user_table.item(row_index, 5) else "")
        permissions = self.mysql_repo.get_system_user_permissions(self.system_user_id)
        for permission_key, check in self.system_user_permission_checks.items():
            check.setChecked(permission_key in permissions)

    def clear_system_user_form(self) -> None:
        self.system_user_id = 0
        self.system_user_name_input.clear()
        self.system_user_login_input.clear()
        self.system_user_password_input.clear()
        self.system_user_active_check.setChecked(True)
        self.system_user_note_input.clear()
        for check in self.system_user_permission_checks.values():
            check.setChecked(False)
        self.system_user_table.clearSelection()

    def save_system_user(self) -> None:
        permissions = {
            permission_key
            for permission_key, check in self.system_user_permission_checks.items()
            if check.isChecked()
        }
        data = {
            "nome": self.system_user_name_input.text(),
            "login": self.system_user_login_input.text(),
            "senha": self.system_user_password_input.text(),
            "ativo": self.system_user_active_check.isChecked(),
            "observacao": self.system_user_note_input.toPlainText(),
        }
        try:
            self.mysql_repo.ensure_schema()
            self.system_user_id = self.mysql_repo.save_system_user(
                self.environment,
                self.system_user_id or None,
                data,
                permissions,
            )
            self.system_user_password_input.clear()
            self.refresh_system_users()
            QMessageBox.information(self, "Usuarios E Permissoes", "Usuario salvo com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Usuarios E Permissoes", str(exc))

    def delete_system_user(self) -> None:
        user_id = self.selected_system_user_id() or self.system_user_id
        if not user_id:
            QMessageBox.information(self, "Usuarios E Permissoes", "Selecione um usuario para excluir.")
            return
        if QMessageBox.question(self, "Usuarios E Permissoes", "Excluir este usuario?") != QMessageBox.Yes:
            return
        try:
            self.mysql_repo.delete_system_user(self.environment, user_id)
            self.clear_system_user_form()
            self.refresh_system_users()
            QMessageBox.information(self, "Usuarios E Permissoes", "Usuario excluido com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Usuarios E Permissoes", str(exc))

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
        self.mysql_host_input = QLineEdit()
        self.mysql_port_input = QLineEdit()
        self.mysql_database_input = QLineEdit()
        self.mysql_user_input = QLineEdit()
        self.mysql_password_input = QLineEdit()
        self.mysql_password_input.setEchoMode(QLineEdit.Password)
        self.mysql_status_label = QLabel("")
        self.mysql_status_label.setObjectName("muted")
        panel_layout.addWidget(title)
        panel_layout.addWidget(config_label)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        form.addWidget(QLabel("Host"), 0, 0)
        form.addWidget(self.mysql_host_input, 0, 1)
        form.addWidget(QLabel("Porta"), 0, 2)
        form.addWidget(self.mysql_port_input, 0, 3)
        form.addWidget(QLabel("Banco"), 1, 0)
        form.addWidget(self.mysql_database_input, 1, 1)
        form.addWidget(QLabel("Usuario"), 1, 2)
        form.addWidget(self.mysql_user_input, 1, 3)
        form.addWidget(QLabel("Senha"), 2, 0)
        form.addWidget(self.mysql_password_input, 2, 1, 1, 3)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)
        panel_layout.addLayout(form)
        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Salvar Config", self.save_mysql_settings))
        actions.addWidget(self.create_button("Criar/Atualizar Banco", self.create_schema, primary=True))
        actions.addWidget(self.create_button("Testar Conexao", self.test_mysql_connection))
        actions.addStretch()
        panel_layout.addLayout(actions)
        panel_layout.addWidget(self.mysql_status_label)
        self.load_mysql_settings()
        layout.addWidget(panel)

        backup_panel = self.create_panel()
        backup_layout = QVBoxLayout(backup_panel)
        backup_layout.setContentsMargins(16, 16, 16, 16)
        backup_layout.setSpacing(8)
        backup_title = QLabel("Backup do Banco de Dados")
        backup_title.setObjectName("sectionTitle")
        backup_layout.addWidget(backup_title)
        backup_desc = QLabel("Gera um arquivo .sql com toda a estrutura e dados do banco atual.")
        backup_desc.setObjectName("muted")
        backup_layout.addWidget(backup_desc)
        backup_row = QHBoxLayout()
        self.backup_btn = self.create_button("Fazer Backup do Banco...", self.run_database_backup, primary=True)
        backup_row.addWidget(self.backup_btn)
        self.backup_progress_bar = QProgressBar()
        self.backup_progress_bar.setVisible(False)
        self.backup_progress_bar.setMaximum(100)
        self.backup_progress_bar.setFixedWidth(260)
        backup_row.addWidget(self.backup_progress_bar)
        backup_row.addStretch()
        backup_layout.addLayout(backup_row)
        self.backup_status_label = QLabel("")
        self.backup_status_label.setObjectName("muted")
        backup_layout.addWidget(self.backup_status_label)
        layout.addWidget(backup_panel)

        layout.addStretch()
        return page

    def build_catalog_import_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Importacao XML para Cadastros")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Gerar Previa", self.build_catalog_import_preview_ui))
        toolbar_layout.addWidget(self.create_button("Importar", self.import_catalogs_from_xml_sources, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.catalog_import_sources_input = QLineEdit()
        self.catalog_import_sources_input.setPlaceholderText("Selecione XML(s) ou pasta com XMLs")
        setup_layout.addWidget(QLabel("Fontes XML"), 0, 0)
        setup_layout.addWidget(self.catalog_import_sources_input, 0, 1, 1, 4)
        setup_layout.addWidget(self.create_button("Selecionar XMLs", self.select_catalog_import_xml_files), 1, 1)
        setup_layout.addWidget(self.create_button("Selecionar Pasta", self.select_catalog_import_folder), 1, 2)
        setup_layout.addWidget(self.create_button("Limpar", lambda: self.catalog_import_sources_input.clear()), 1, 3)
        setup_layout.addWidget(self.create_button("Gerar Previa", self.build_catalog_import_preview_ui), 1, 4)
        setup_layout.addWidget(self.create_button("Importar", self.import_catalogs_from_xml_sources, primary=True), 1, 5)
        self.catalog_import_ignored_button = self.create_button("Ver Ignorados", self.open_catalog_import_ignored_popup)
        self.catalog_import_ignored_button.setEnabled(False)
        setup_layout.addWidget(self.catalog_import_ignored_button, 1, 6)
        setup_layout.setColumnStretch(1, 1)
        self.catalog_import_status = QLabel("Importe XMLs para cadastrar Empresa, Fornecedor e Produtos automaticamente.")
        self.catalog_import_status.setObjectName("muted")
        setup_layout.addWidget(self.catalog_import_status, 2, 0, 1, 6)
        self.catalog_import_progress = QProgressBar()
        self.catalog_import_progress.setMinimum(0)
        self.catalog_import_progress.setMaximum(100)
        self.catalog_import_progress.setValue(0)
        setup_layout.addWidget(self.catalog_import_progress, 3, 0, 1, 6)
        layout.addWidget(setup)

        tips = self.create_panel()
        tips_layout = QVBoxLayout(tips)
        tips_layout.setContentsMargins(12, 12, 12, 12)
        tips_layout.addWidget(QLabel("Regras da importacao"))
        tips_label = QLabel(
            "1) Emitente da NF-e vira Fornecedor.\n"
            "2) Destinatario da NF-e vira Empresa.\n"
            "3) Produto vincula por Fornecedor + EAN (ou codigo quando EAN vazio).\n"
            "4) Classificacao do produto fica em branco para selecionar depois."
        )
        tips_label.setObjectName("muted")
        tips_layout.addWidget(tips_label)
        layout.addWidget(tips)

        preview_panel = self.create_panel()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(12, 12, 12, 12)
        preview_layout.setSpacing(8)
        preview_layout.addWidget(QLabel("Previa da Importacao"))
        self.catalog_import_preview_table = self.create_data_table(
            [
                "XML",
                "Empresa",
                "Empresa Existe?",
                "Fornecedor",
                "UF",
                "Fornecedor Existe?",
                "Cod. Forn.",
                "Descricao",
                "EAN",
                "NCM",
                "CST ICMS",
                "% Red BC ICMS",
                "% ICMS",
                "BC ST",
                "Valor ICMS ST",
                "% ICMS ST",
                "CST PIS",
                "CST COFINS",
                "% PIS",
                "% COFINS",
                "Classificacao",
                "Ja Existe?",
            ]
        )
        self.catalog_import_preview_table.setObjectName("catalog_import_preview")
        self.catalog_import_preview_rows: list[CatalogImportPreviewRow] = []
        self.catalog_import_ignored_rows: list[list[str]] = []
        self.catalog_import_skipped_product_rows: list[list[str]] = []
        self.catalog_import_preview_table.cellDoubleClicked.connect(self.open_catalog_import_xml_file)
        preview_layout.addWidget(self.catalog_import_preview_table, 1)
        layout.addWidget(preview_panel, 1)
        layout.addStretch()
        return page

    def load_mysql_settings(self) -> None:
        config = self.mysql_repo.load_config()
        self.mysql_host_input.setText(str(config.get("host", "")))
        self.mysql_port_input.setText(str(config.get("port", "3306")))
        self.mysql_database_input.setText(str(config.get("database", "")))
        self.mysql_user_input.setText(str(config.get("user", "")))
        self.mysql_password_input.setText(str(config.get("password", "")))
        self.mysql_status_label.setText("Configuracao carregada.")

    def mysql_form_payload(self) -> dict[str, str]:
        defaults = self.get_mysql_default_config()
        return {
            "host": self.mysql_host_input.text().strip() or defaults["host"],
            "port": self.mysql_port_input.text().strip() or "3306",
            "database": self.mysql_database_input.text().strip() or defaults["database"],
            "user": self.mysql_user_input.text().strip() or defaults["user"],
            "password": self.mysql_password_input.text(),
        }

    def save_mysql_settings(self) -> None:
        try:
            self.mysql_repo.save_config(self.mysql_form_payload())
            self.mysql_status_label.setText("Configuracao MySQL salva.")
            self.statusBar().showMessage("Configuracao MySQL salva.")
            QMessageBox.information(self, "Configuracoes", "Configuracao MySQL salva com sucesso.")
        except Exception as exc:
            self.mysql_status_label.setText(f"Falha ao salvar configuracao: {exc}")
            self.statusBar().showMessage("Falha ao salvar configuracao MySQL.")
            QMessageBox.critical(self, "Configuracoes", f"Nao foi possivel salvar a configuracao MySQL.\n\n{exc}")

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
        self.apply_table_column_policy(table)
        self.enable_table_sorting(table)
        return table

    def create_product_table_view(self, headers: list[str]) -> QTableView:
        table = QTableView()
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setWordWrap(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSortIndicatorShown(False)
        table.horizontalHeader().setSectionsClickable(False)
        table.horizontalHeader().setMinimumSectionSize(70)
        table.horizontalHeader().setStretchLastSection(False)
        table.setModel(ProductTableModel(headers))
        self.apply_product_table_column_widths(table)
        return table

    def apply_product_table_column_widths(self, table: QTableView) -> None:
        base_widths = [90, 90, 230, 260, 70, 140, 110, 120, 360, 150, 110, 110]
        available_width = max(0, table.viewport().width() - 6)
        total_base_width = sum(base_widths)
        if available_width > total_base_width:
            extra_width = available_width - total_base_width
            weights = [1, 1, 4, 5, 1, 2, 1, 1, 6, 2, 1, 1]
            total_weight = sum(weights)
            widths = [
                base_width + int(extra_width * weight / total_weight)
                for base_width, weight in zip(base_widths, weights)
            ]
        else:
            widths = base_widths
        for column_index, width in enumerate(widths):
            table.setColumnWidth(column_index, width)

    def resizeEvent(self, event: object) -> None:
        super().resizeEvent(event)
        if hasattr(self, "product_page_table"):
            QTimer.singleShot(0, lambda: self.apply_product_table_column_widths(self.product_page_table))

    def configure_searchable_combo(self, combo: QComboBox, placeholder: str) -> None:
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMaxVisibleItems(18)
        if combo.lineEdit() is not None:
            combo.lineEdit().setPlaceholderText(placeholder)
        completer = combo.completer()
        if completer is not None:
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setFilterMode(Qt.MatchContains)
            completer.setCompletionMode(QCompleter.PopupCompletion)

    def enable_table_sorting(self, table: QTableWidget) -> None:
        table.setSortingEnabled(True)
        header = table.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSectionsClickable(True)
        header.setMinimumSectionSize(70)
        self.ensure_full_header_visibility(table)

    def ensure_full_header_visibility(self, table: QTableWidget) -> None:
        header = table.horizontalHeader()
        font_metrics = header.fontMetrics()
        for column_index in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(column_index)
            header_text = header_item.text() if header_item else ""
            minimum_width = max(70, font_metrics.horizontalAdvance(header_text) + 28)
            if table.columnWidth(column_index) < minimum_width:
                table.setColumnWidth(column_index, minimum_width)

    def apply_table_column_policy(self, table: QTableWidget) -> None:
        short_labels = (
            "id",
            "ie",
            "cst",
            "cfop",
            "ncm",
            "cest",
            "ean",
            "aliq",
            "%",
            "bc",
            "st",
            "vl",
            "valor",
            "data",
            "serie",
            "modelo",
            "cod",
            "codigo",
            "num",
            "numero",
            "qtd",
            "qtde",
        )
        header = table.horizontalHeader()
        for index in range(table.columnCount()):
            label_item = table.horizontalHeaderItem(index)
            label = (label_item.text() if label_item else "").strip().lower()
            if any(token in label for token in short_labels):
                header.setSectionResizeMode(index, QHeaderView.ResizeToContents)
            else:
                header.setSectionResizeMode(index, QHeaderView.Stretch)
        header.setStretchLastSection(True)

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
        self.entry_reprocess_button = self.create_button("Reprocessar SPED", lambda: self.generate_adjusted_sped_from_consultation("Entrada"))
        self.entry_reprocess_button.setEnabled(False)
        toolbar_layout.addWidget(self.entry_reprocess_button)
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
        filter_actions.addWidget(self.create_button("Confronto", self.open_confronto_entry, primary=True))
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
                ("base_ratio", "% Base/Oper"),
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
        self.entry_table.horizontalHeader().setStretchLastSection(True)
        self.entry_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
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
        self.entry_table.itemDoubleClicked.connect(self.handle_entry_table_double_click)
        self.entry_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_table.customContextMenuRequested.connect(lambda pos: self.show_runtime_rule_context_menu_for_table(self.entry_table, self.filtered_entry_rows, "Entrada", pos))
        self.enable_table_sorting(self.entry_table)
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
        self.sale_reprocess_button = self.create_button("Reprocessar SPED", lambda: self.generate_adjusted_sped_from_consultation("Saida"))
        self.sale_reprocess_button.setEnabled(False)
        toolbar_layout.addWidget(self.sale_reprocess_button)
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
        filter_actions.addWidget(self.create_button("Confronto", self.open_confronto_sale, primary=True))
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
        self.sale_table.horizontalHeader().setStretchLastSection(True)
        self.sale_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for column, width in enumerate((88, 72, 105, 330, 260, 90, 90, 85, 80, 85, 105, 65, 75)):
            self.sale_table.setColumnWidth(column, width)
        self.sale_table.itemDoubleClicked.connect(self.handle_sale_table_double_click)
        self.sale_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.sale_table.customContextMenuRequested.connect(lambda pos: self.show_runtime_rule_context_menu_for_table(self.sale_table, self.filtered_sale_rows, "Saida", pos))
        self.enable_table_sorting(self.sale_table)
        layout.addWidget(self.sale_table, 1)
        self.bind_sale_live_filters()
        return page

    def build_entry_exit_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Analise Entrada e Saida")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_entry_exit_page))
        toolbar_layout.addWidget(self.create_button("Exportar Excel", self.export_entry_exit_analysis))
        toolbar_layout.addWidget(self.create_button("Processar", self.process_entry_exit_analysis, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.entry_exit_sped_input = QLineEdit()
        self.entry_exit_sped_input.setPlaceholderText("Selecione um ou mais arquivos SPED Fiscal")
        self.add_path_row(setup_layout, 0, "SPEDs", self.entry_exit_sped_input, self.select_entry_exit_sped_files, lambda: self.entry_exit_sped_input.clear())
        layout.addWidget(setup)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(8)
        self.entry_exit_metric_labels: dict[str, QLabel] = {}
        for column, (key, label) in enumerate(
            (
                ("saida_icms", "Saidas ICMS"),
                ("entrada_icms", "Entradas ICMS"),
                ("recolher", "A recolher"),
                ("percent_sale", "% sobre venda"),
                ("source_files", "Arquivos"),
            )
        ):
            card, value_label = self.create_metric_card_with_label(label, "0")
            self.entry_exit_metric_labels[key] = value_label
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        self.entry_exit_tabs = QTabWidget()
        self.entry_exit_sale_table = self.create_data_table(self.entry_exit_detail_headers())
        self.entry_exit_entry_table = self.create_data_table(self.entry_exit_detail_headers())
        self.entry_exit_sale_footer_table = self.create_data_table(self.entry_exit_footer_headers())
        self.entry_exit_entry_footer_table = self.create_data_table(self.entry_exit_footer_headers())
        self.entry_exit_tabs.addTab(self.entry_exit_sale_table, "Saidas")
        self.entry_exit_tabs.addTab(self.entry_exit_entry_table, "Entradas")
        self.entry_exit_tabs.addTab(self.entry_exit_sale_footer_table, "Resumo Saidas")
        self.entry_exit_tabs.addTab(self.entry_exit_entry_footer_table, "Resumo Entradas")
        layout.addWidget(self.entry_exit_tabs, 1)

        self.entry_exit_status_label = QLabel("Selecione os SPEDs e clique em Processar.")
        self.entry_exit_status_label.setObjectName("muted")
        layout.addWidget(self.entry_exit_status_label)
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
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for column, width in enumerate((88, 72, 105, 330, 260, 90, 90, 85, 85, 105, 90, 105, 65, 75)):
            table.setColumnWidth(column, width)
        table.itemDoubleClicked.connect(lambda _item, op=operation_type: self.open_contrib_docs_popup(op))
        self.enable_table_sorting(table)
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
        toolbar_layout.addWidget(self.create_button("Adicionar Regra Assistida", lambda: self.open_runtime_rule_builder(), primary=True))
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_runtime_rules))
        toolbar_layout.addWidget(self.create_button("Validar regras", self.validate_runtime_rules))
        layout.addWidget(toolbar)

        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)

        editor_panel = self.create_panel()
        editor_layout = QVBoxLayout(editor_panel)
        editor_layout.setContentsMargins(12, 12, 12, 12)
        editor_title = QLabel("Regras para reprocessamento")
        editor_title.setObjectName("metricKey")
        editor_layout.addWidget(editor_title)
        self.runtime_rules_text = QTextEdit()
        self.runtime_rules_text.setPlaceholderText(
            "Uma regra por linha. Exemplo: tipo=entrada; cst=000; cfop=1102; novo_cfop=1101; recalcular_valor_icms=sim"
        )
        editor_layout.addWidget(self.runtime_rules_text, 1)
        content_layout.addWidget(editor_panel, 3)

        history_panel = self.create_panel()
        history_layout = QVBoxLayout(history_panel)
        history_layout.setContentsMargins(12, 12, 12, 12)
        history_title = QLabel("Historico de regras")
        history_title.setObjectName("metricKey")
        history_layout.addWidget(history_title)

        history_filter_row = QHBoxLayout()
        self.runtime_rule_filter_input = QLineEdit()
        self.runtime_rule_filter_input.setPlaceholderText("Filtrar por CST, CFOP, CNPJ, codigo ou acao")
        self.runtime_rule_filter_input.returnPressed.connect(self.refresh_runtime_rule_history_list)
        history_filter_row.addWidget(self.runtime_rule_filter_input, 1)
        history_filter_row.addWidget(self.create_button("Filtrar", self.refresh_runtime_rule_history_list))
        history_filter_row.addWidget(self.create_button("Limpar", self.clear_runtime_rule_filter))
        history_layout.addLayout(history_filter_row)

        self.runtime_rule_history_table = self.create_data_table(["Regra", "Resumo", "Usos", "Ultimo uso"])
        self.runtime_rule_history_table.setColumnWidth(0, 360)
        self.runtime_rule_history_table.setColumnWidth(1, 300)
        self.runtime_rule_history_table.setColumnWidth(2, 70)
        self.runtime_rule_history_table.setColumnWidth(3, 145)
        self.runtime_rule_history_table.itemSelectionChanged.connect(self.update_runtime_rule_history_preview)
        self.runtime_rule_history_table.cellDoubleClicked.connect(lambda _row, _column: self.insert_selected_runtime_rule())
        history_layout.addWidget(self.runtime_rule_history_table, 1)

        self.runtime_rule_history_preview = QTextEdit()
        self.runtime_rule_history_preview.setReadOnly(True)
        self.runtime_rule_history_preview.setMaximumHeight(82)
        history_layout.addWidget(self.runtime_rule_history_preview)

        history_buttons = QHBoxLayout()
        history_buttons.addWidget(self.create_button("Inserir", self.insert_selected_runtime_rule, primary=True))
        history_buttons.addWidget(self.create_button("Editar", self.edit_selected_runtime_rule))
        history_buttons.addWidget(self.create_button("Excluir", self.delete_selected_runtime_rule))
        history_buttons.addStretch()
        history_layout.addLayout(history_buttons)
        content_layout.addWidget(history_panel, 2)
        layout.addLayout(content_layout, 2)

        self.runtime_rules_status = QLabel("Informe as regras e clique em Validar regras.")
        self.runtime_rules_status.setObjectName("muted")
        layout.addWidget(self.runtime_rules_status)

        self.runtime_rules_table = self.create_data_table(["Linha", "Resumo", "Regra"])
        self.runtime_rules_table.setColumnWidth(0, 70)
        self.runtime_rules_table.setColumnWidth(1, 560)
        self.runtime_rules_table.setColumnWidth(2, 760)
        layout.addWidget(self.runtime_rules_table, 1)
        self.refresh_runtime_rule_history_list()
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
        key.setWordWrap(True)
        value_label = QLabel(value)
        value_label.setObjectName("metricValue")
        value_label.setWordWrap(True)
        value_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(key)
        layout.addWidget(value_label)
        return card, value_label

    def make_product_input_compact(self, widget: QWidget) -> None:
        widget.setMinimumHeight(24)
        widget.setMaximumHeight(30)

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
        was_sorting_enabled = table.isSortingEnabled()
        if was_sorting_enabled:
            table.setSortingEnabled(False)
        table.setUpdatesEnabled(False)
        header_labels = [
            (table.horizontalHeaderItem(index).text().strip().lower() if table.horizontalHeaderItem(index) else "")
            for index in range(table.columnCount())
        ]
        try:
            table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    display_value = str(value)
                    header_label = header_labels[column_index] if column_index < len(header_labels) else ""
                    if header_label == "cnpj":
                        display_value = self.format_cnpj(display_value)
                    elif header_label == "ie":
                        display_value = self.format_ie(display_value)
                    item = QTableWidgetItem(display_value)
                    item.setForeground(QColor(COLORS["text"]))
                    if header_label in {"descricao", "produto"}:
                        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                    else:
                        item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(row_index, column_index, item)
            # Evita recalculo pesado de largura no refresh de Produtos.
            if table.objectName() not in {"catalog_import_preview", "produtos"} or len(rows) <= 1000:
                self.apply_table_column_policy(table)
        finally:
            table.setUpdatesEnabled(True)
            if was_sorting_enabled:
                table.setSortingEnabled(True)

    def format_cnpj(self, value: object) -> str:
        digits = self.digits_only(str(value or ""))
        if len(digits) != 14:
            return str(value or "")
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"

    def format_ie(self, value: object) -> str:
        digits = self.digits_only(str(value or ""))
        return digits

    def selected_catalog_id(self, table: QTableWidget) -> int:
        selected = table.selectionModel().selectedRows()
        if not selected:
            return 0
        id_column = self.get_table_id_column_index(table)
        item = table.item(selected[0].row(), id_column)
        return int(item.text()) if item and item.text().isdigit() else 0

    def get_table_id_column_index(self, table: QTableWidget) -> int:
        for index in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(index)
            header_text = str(header_item.text() if header_item else "").strip().lower()
            if header_text == "id":
                return index
        return 0

    def run_guarded_ui_operation(self, operation_key: str, message: str, callback: Callable[[], None]) -> None:
        running_operations = getattr(self, "_running_ui_operations", set())
        if operation_key in running_operations:
            self.statusBar().showMessage("Aguarde a consulta atual terminar.")
            return
        running_operations.add(operation_key)
        self._running_ui_operations = running_operations
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.statusBar().showMessage(message)
        QApplication.processEvents()
        try:
            callback()
        except Exception as exc:
            self.statusBar().showMessage("Falha na consulta.")
            QMessageBox.critical(self, "Consulta", str(exc))
        finally:
            running_operations.discard(operation_key)
            QApplication.restoreOverrideCursor()
            QApplication.processEvents()

    def company_display_name(self, row: dict[str, object]) -> str:
        company_name = str(row.get("nome", row.get("empresa_nome", "")) or "").strip()
        company_cnpj = self.format_cnpj(str(row.get("cnpj", row.get("empresa_cnpj", "")) or ""))
        return f"{company_name} - {company_cnpj}" if company_cnpj else company_name

    def refresh_current_catalog_page(self) -> None:
        if self.page_title.text() == "Produtos":
            self.start_refresh_product_page_fast()
            return

        def refresh() -> None:
            title = self.page_title.text()
            if title == "Empresas":
                self.refresh_company_page()
            elif title == "Fornecedores":
                self.refresh_supplier_page()
            elif title == "Classificacao do Produto":
                self.refresh_type_page()
            elif title == "NCM":
                self.refresh_ncm_page()
            else:
                self.refresh_catalog()
            self.statusBar().showMessage("Consulta carregada.")

        self.run_guarded_ui_operation("refresh_current_catalog_page", "Carregando consulta...", refresh)

    def fill_company_combo(self, combo: QComboBox) -> None:
        combo.clear()
        for row in self.mysql_repo.list_companies(self.environment):
            combo.addItem(self.company_display_name(row), int(row["id"]))

    def fill_supplier_combo(self, combo: QComboBox) -> None:
        self.fill_supplier_combo_from_rows(combo, self.mysql_repo.list_suppliers_catalog(self.environment))

    def fill_supplier_combo_from_rows(self, combo: QComboBox, rows: list[dict[str, object]]) -> None:
        combo.clear()
        self.catalog_supplier_company_by_id = {}
        for row in rows:
            self.catalog_supplier_company_by_id[int(row["id"])] = int(row.get("empresa_id") or 0)
            combo.addItem(f"{row.get('empresa_nome', '')} / {row.get('nome', '')}", int(row["id"]))

    def populate_regime_tributario_combo(self, combo: QComboBox) -> None:
        combo.clear()
        combo.addItem("Lucro Real/Presumido", "LUCRO_REAL_PRESUMIDO")
        combo.addItem("Simples Nacional", "SIMPLES_NACIONAL")

    def fill_type_combo(self, combo: QComboBox) -> None:
        self.fill_type_combo_from_rows(combo, self.mysql_repo.list_product_types(self.environment))

    def fill_type_combo_from_rows(self, combo: QComboBox, rows: list[dict[str, object]]) -> None:
        combo.clear()
        combo.addItem("", 0)
        for row in rows:
            combo.addItem(str(row.get("nome", "")), int(row["id"]))

    def refresh_company_page(self) -> None:
        self.mysql_repo.ensure_schema()
        rows = self.mysql_repo.list_companies(self.environment)
        self.company_page_rows = {int(row["id"]): row for row in rows}
        self.set_table_rows(
            self.company_page_table,
            [[int(row["id"]), row.get("nome", ""), row.get("cnpj", ""), row.get("inscricao_estadual", "")] for row in rows],
        )

    def handle_company_page_select(self) -> None:
        self.fill_line_fields(self.company_page_fields, getattr(self, "company_page_rows", {}).get(self.selected_catalog_id(self.company_page_table), {}))

    def clear_company_page_form(self) -> None:
        self.fill_line_fields(self.company_page_fields, {})

    def save_company_page(self) -> None:
        try:
            payload = self.line_field_payload(self.company_page_fields)
            row_before = int(payload.get("id") or 0)
            if not self.validate_required_cnpj(str(payload.get("cnpj", "")), "Empresas"):
                return
            row_id = self.mysql_repo.save_company(self.environment, payload)
            self.refresh_company_page()
            self.select_table_row_by_id(self.company_page_table, row_id)
            action = "atualizada" if row_before else "cadastrada"
            QMessageBox.information(self, "Empresas", f"Empresa {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Empresas", str(exc))

    def delete_company_page(self) -> None:
        row_id = self.selected_catalog_id(self.company_page_table)
        if row_id and QMessageBox.question(self, "Empresas", "Excluir a empresa e seus fornecedores/produtos?") == QMessageBox.Yes:
            self.mysql_repo.delete_company(row_id)
            self.refresh_company_page()

    def refresh_supplier_page(self) -> None:
        self.mysql_repo.ensure_schema()
        self.fill_company_combo(self.supplier_page_company_combo)
        rows = self.mysql_repo.list_suppliers_catalog(self.environment)
        self.supplier_page_rows = {int(row["id"]): row for row in rows}
        self.set_table_rows(
            self.supplier_page_table,
            [
                [
                    int(row["id"]),
                    row.get("empresa_nome", ""),
                    row.get("nome", ""),
                    row.get("cnpj", ""),
                    row.get("inscricao_estadual", ""),
                    row.get("uf", ""),
                    row.get("codigo", ""),
                    self.regime_tributario_label(str(row.get("regime_tributario", ""))),
                ]
                for row in rows
            ],
        )

    def handle_supplier_page_select(self) -> None:
        row = getattr(self, "supplier_page_rows", {}).get(self.selected_catalog_id(self.supplier_page_table), {})
        self.fill_line_fields(self.supplier_page_fields, row)
        self.select_combo_data(self.supplier_page_company_combo, int(row.get("empresa_id") or 0))
        self.select_regime_tributario_combo(self.supplier_page_regime_combo, str(row.get("regime_tributario", "")))

    def clear_supplier_page_form(self) -> None:
        self.fill_line_fields(self.supplier_page_fields, {})
        self.select_regime_tributario_combo(self.supplier_page_regime_combo, "")

    def save_supplier_page(self) -> None:
        try:
            payload = self.line_field_payload(self.supplier_page_fields)
            row_before = int(payload.get("id") or 0)
            if not self.validate_required_cnpj(str(payload.get("cnpj", "")), "Fornecedores"):
                return
            payload["regime_tributario"] = self.normalize_regime_tributario_value(self.supplier_page_regime_combo.currentData())
            company_id = int(self.supplier_page_company_combo.currentData() or 0)
            row_id = self.mysql_repo.save_supplier(company_id, payload)
            self.refresh_supplier_page()
            self.select_table_row_by_id(self.supplier_page_table, row_id)
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Fornecedores", f"Fornecedor {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Fornecedores", str(exc))

    def delete_supplier_page(self) -> None:
        row_id = self.selected_catalog_id(self.supplier_page_table)
        if row_id and QMessageBox.question(self, "Fornecedores", "Excluir o fornecedor e seus produtos?") == QMessageBox.Yes:
            self.mysql_repo.delete_supplier(row_id)
            self.refresh_supplier_page()

    def refresh_type_page(self) -> None:
        self.mysql_repo.ensure_schema()
        rows = self.mysql_repo.list_product_types_catalog(self.environment)
        self.type_page_rows = {int(row["id"]): row for row in rows}
        self.set_table_rows(
            self.type_page_table,
            [[int(row["id"]), row.get("nome", ""), row.get("descricao", "")] for row in rows],
        )

    def handle_type_page_select(self) -> None:
        row = getattr(self, "type_page_rows", {}).get(self.selected_catalog_id(self.type_page_table), {})
        self.fill_line_fields(self.type_page_fields, row)

    def clear_type_page_form(self) -> None:
        self.fill_line_fields(self.type_page_fields, {})

    def save_type_page(self) -> None:
        try:
            payload = self.line_field_payload(self.type_page_fields)
            row_before = int(payload.get("id") or 0)
            row_id = self.mysql_repo.save_product_type(self.environment, payload)
            self.refresh_type_page()
            self.select_table_row_by_id(self.type_page_table, row_id)
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Classificacao do Produto", f"Classificacao do produto {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Classificacao do Produto", str(exc))

    def delete_type_page(self) -> None:
        row_id = self.selected_catalog_id(self.type_page_table)
        if row_id and QMessageBox.question(self, "Classificacao do Produto", "Excluir esta classificacao? Produtos vinculados ficarao sem classificacao.") == QMessageBox.Yes:
            self.mysql_repo.delete_product_type(row_id)
            self.refresh_type_page()

    def refresh_ncm_page(self) -> None:
        self.mysql_repo.ensure_schema()
        rows = self.mysql_repo.list_ncm_catalog(self.environment)
        self.ncm_page_rows = {int(row["id"]): row for row in rows}
        self.set_table_rows(
            self.ncm_table,
            [
                [
                    int(row["id"]),
                    row.get("atividade", ""),
                    row.get("regime_tributario", ""),
                    row.get("uf", ""),
                    row.get("data_vigencia", ""),
                    row.get("ncm", ""),
                    row.get("descricao", ""),
                    row.get("cest", ""),
                    row.get("aliquota_ipi", ""),
                    row.get("cst_ipi", ""),
                    row.get("ex_tipi", ""),
                    row.get("cst_pis_cofins_entrada", ""),
                    row.get("cst_pis_cofins_saida", ""),
                    row.get("codigo_sped", ""),
                    row.get("aliquota_pis", ""),
                    row.get("aliquota_cofins", ""),
                    row.get("base_legal_pis_cofins", ""),
                    row.get("cfop_entrada", ""),
                    row.get("cfop_saida", ""),
                    row.get("cst_csosn", ""),
                    row.get("ad_rem_icms", ""),
                    row.get("aliquota_icms", ""),
                    row.get("reducao_bc_icms", ""),
                    row.get("reducao_bc_icms_st", ""),
                    row.get("aliquota_icms_st", ""),
                    row.get("aliquota_red_bc_icms", ""),
                    row.get("mva", ""),
                    row.get("fcp", ""),
                    row.get("codigo_beneficio_fiscal", ""),
                    row.get("antecipado", ""),
                    row.get("percentual_diferimento", ""),
                    row.get("percentual_isencao", ""),
                    row.get("codigo_anp", ""),
                    row.get("base_legal_icms", ""),
                ]
                for row in rows
            ],
        )

    def handle_ncm_select(self) -> None:
        row = getattr(self, "ncm_page_rows", {}).get(self.selected_catalog_id(self.ncm_table), {})
        self.fill_line_fields(self.ncm_fields, row)

    def clear_ncm_form(self) -> None:
        self.fill_line_fields(self.ncm_fields, {})

    def save_ncm_rule(self) -> None:
        try:
            payload = self.line_field_payload(self.ncm_fields)
            row_before = int(payload.get("id") or 0)
            row_id = self.mysql_repo.save_ncm_rule(self.environment, payload)
            self.refresh_ncm_page()
            self.select_table_row_by_id(self.ncm_table, row_id)
            action = "atualizada" if row_before else "cadastrada"
            QMessageBox.information(self, "NCM", f"Regra NCM {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "NCM", str(exc))

    def delete_ncm_rule(self) -> None:
        row_id = self.selected_catalog_id(self.ncm_table)
        if row_id and QMessageBox.question(self, "NCM", "Excluir esta regra de NCM?") == QMessageBox.Yes:
            self.mysql_repo.delete_ncm_rule(row_id)
            self.refresh_ncm_page()

    def select_ncm_import_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar planilha de NCM", "", "Planilhas Excel (*.xlsx *.xlsm)")
        if file_path:
            self.ncm_import_path_input.setText(file_path)

    def import_ncm_rules_from_excel(self) -> None:
        path = Path(self.ncm_import_path_input.text().strip())
        if not path.exists():
            QMessageBox.warning(self, "NCM", "Selecione uma planilha valida para importar.")
            return
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            stats = self.mysql_repo.import_ncm_rules_from_excel(self.environment, path)
            self.refresh_ncm_page()
            QMessageBox.information(
                self,
                "NCM",
                f"Importacao concluida.\nLinhas processadas: {stats['rows']}\nIgnoradas: {stats['ignored']}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "NCM", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def import_product_page_review_excel(self) -> None:
        if getattr(self, "product_review_import_thread", None) is not None:
            QMessageBox.information(self, "Importar Revisao de Produtos", "Ja existe uma importacao em andamento.")
            return
        selected_company = self.select_product_review_import_company()
        if not selected_company:
            return
        company_id, company_name = selected_company
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar planilha revisada de produtos",
            "",
            "Planilhas Excel (*.xlsx)",
        )
        if not file_path:
            return
        if QMessageBox.question(
            self,
            "Importar Revisao de Produtos",
            f'Esta rotina vai atualizar produtos da empresa "{company_name}" lendo a aba "BASE_COMPLETA" e localizando cada produto por Fornecedor e Cod. Forn. Deseja continuar?',
        ) != QMessageBox.Yes:
            return
        if hasattr(self, "product_review_import_button"):
            self.product_review_import_button.setEnabled(False)
        if hasattr(self, "product_review_import_progress"):
            self.product_review_import_progress.setValue(0)
        if hasattr(self, "product_review_import_status"):
            self.product_review_import_status.setText(f"Iniciando importacao para {company_name}...")
        self.statusBar().showMessage("Importando revisao de produtos...")
        self.product_review_import_thread = QThread(self)
        self.product_review_import_worker = ProductReviewImportWorker(self.mysql_repo, self.environment, company_id, Path(file_path))
        self.product_review_import_worker.moveToThread(self.product_review_import_thread)
        self.product_review_import_thread.started.connect(self.product_review_import_worker.run)
        self.product_review_import_worker.progress.connect(self.update_product_review_import_progress)
        self.product_review_import_worker.finished.connect(self.handle_product_review_import_finished)
        self.product_review_import_worker.failed.connect(self.handle_product_review_import_failed)
        self.product_review_import_worker.finished.connect(self.product_review_import_thread.quit)
        self.product_review_import_worker.failed.connect(self.product_review_import_thread.quit)
        self.product_review_import_thread.finished.connect(self.product_review_import_worker.deleteLater)
        self.product_review_import_thread.finished.connect(self.product_review_import_thread.deleteLater)
        self.product_review_import_thread.finished.connect(lambda: setattr(self, "product_review_import_worker", None))
        self.product_review_import_thread.finished.connect(lambda: setattr(self, "product_review_import_thread", None))
        self.product_review_import_thread.start()

    def select_product_review_import_company(self) -> tuple[int, str] | None:
        try:
            companies = self.mysql_repo.list_companies(self.environment)
        except Exception as exc:
            QMessageBox.critical(self, "Importar Revisao de Produtos", str(exc))
            return None
        if not companies:
            QMessageBox.warning(self, "Importar Revisao de Produtos", "Nao ha empresas cadastradas para selecionar.")
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar Empresa")
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)
        layout.addWidget(QLabel("Selecione a empresa que tera os produtos atualizados:"))
        combo = QComboBox()
        combo.setMinimumWidth(420)
        self.configure_searchable_combo(combo, "Digite para localizar empresa")
        for company in companies:
            company_id = int(company.get("id") or 0)
            company_name = str(company.get("nome", "") or "").strip()
            if company_id and company_name:
                combo.addItem(self.company_display_name(company), company_id)
        layout.addWidget(combo)
        actions = QHBoxLayout()
        actions.addStretch()
        cancel_button = self.create_button("Cancelar", dialog.reject)
        confirm_button = self.create_button("Selecionar", dialog.accept, primary=True)
        actions.addWidget(cancel_button)
        actions.addWidget(confirm_button)
        layout.addLayout(actions)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected_company_id = int(combo.currentData() or 0)
        selected_company_name = combo.currentText().strip()
        if not selected_company_id:
            return None
        return selected_company_id, selected_company_name

    def update_product_review_import_progress(self, current: int, total: int, message: str) -> None:
        if hasattr(self, "product_review_import_progress"):
            value = 0 if total <= 0 else int((max(0, current) / max(1, total)) * 100)
            self.product_review_import_progress.setValue(max(0, min(100, value)))
        if hasattr(self, "product_review_import_status"):
            self.product_review_import_status.setText(message)
        self.statusBar().showMessage(message)

    def handle_product_review_import_finished(self, stats: dict[str, object]) -> None:
        if hasattr(self, "product_review_import_button"):
            self.product_review_import_button.setEnabled(True)
        if hasattr(self, "product_review_import_progress"):
            self.product_review_import_progress.setValue(100)
        log_path = str(stats.get("log_path", "") or "")
        message = (
            "Importacao concluida.\n"
            f"Linhas processadas: {stats['rows']}\n"
            f"Produtos atualizados: {stats['updated']}\n"
            f"Valores alterados: {stats['changed_values']}\n"
            f"Classificacoes criadas: {stats['created_types']}\n"
            f"Produtos nao encontrados: {stats['missing_products']}\n"
            f"Chaves duplicadas: {stats['duplicate_products']}\n"
            f"Linhas ignoradas: {stats['ignored']}\n"
            f"Erros: {stats['errors']}"
        )
        if log_path:
            message += f"\n\nLog de alteracoes:\n{log_path}"
        error_messages = list(stats.get("error_messages", [])) if isinstance(stats, dict) else []
        if error_messages:
            message += "\n\nPrimeiros erros:\n" + "\n".join(str(error) for error in error_messages)
        if hasattr(self, "product_review_import_status"):
            self.product_review_import_status.setText("Revisao importada. Use Consultar/Atualizar para recarregar a consulta quando precisar.")
        QMessageBox.information(self, "Importar Revisao de Produtos", message)
        self.statusBar().showMessage("Revisao de produtos importada. Consulta nao foi recarregada automaticamente.")

    def handle_product_review_import_failed(self, error_text: str) -> None:
        if hasattr(self, "product_review_import_button"):
            self.product_review_import_button.setEnabled(True)
        if hasattr(self, "product_review_import_status"):
            self.product_review_import_status.setText(f"Falha na importacao: {error_text}")
        QMessageBox.critical(self, "Importar Revisao de Produtos", error_text)
        self.statusBar().showMessage("Falha na importacao da revisao de produtos.")

    def refresh_product_page(self) -> None:
        self.start_refresh_product_page_fast()

    def refresh_product_page_fast(self) -> None:
        self.start_refresh_product_page_fast()

    def start_refresh_product_page_fast(self) -> None:
        if getattr(self, "product_page_refresh_thread", None) is not None:
            self.statusBar().showMessage("Aguarde a consulta atual terminar.")
            return
        if hasattr(self, "product_page_refresh_button"):
            self.product_page_refresh_button.setEnabled(False)
        self.statusBar().showMessage("Atualizando produtos...")
        QApplication.processEvents()
        self.product_page_refresh_thread = QThread(self)
        self.product_page_refresh_worker = ProductCatalogWorker(self.mysql_repo, self.environment)
        self.product_page_refresh_worker.moveToThread(self.product_page_refresh_thread)
        self.product_page_refresh_thread.started.connect(self.product_page_refresh_worker.run)
        self.product_page_refresh_worker.finished.connect(self.handle_product_page_refresh_finished)
        self.product_page_refresh_worker.failed.connect(self.handle_product_page_refresh_failed)
        self.product_page_refresh_worker.finished.connect(self.product_page_refresh_thread.quit)
        self.product_page_refresh_worker.failed.connect(self.product_page_refresh_thread.quit)
        self.product_page_refresh_thread.finished.connect(self.product_page_refresh_worker.deleteLater)
        self.product_page_refresh_thread.finished.connect(self.product_page_refresh_thread.deleteLater)
        self.product_page_refresh_thread.finished.connect(lambda: setattr(self, "product_page_refresh_worker", None))
        self.product_page_refresh_thread.finished.connect(lambda: setattr(self, "product_page_refresh_thread", None))
        self.product_page_refresh_thread.start()

    def handle_product_page_refresh_finished(self, payload: object) -> None:
        try:
            data = payload if isinstance(payload, dict) else {}
            rows = list(data.get("products", []))
            supplier_rows = list(data.get("suppliers", []))
            type_rows = list(data.get("types", []))
            self.fill_supplier_combo_from_rows(self.product_page_supplier_combo, supplier_rows)
            self.fill_type_combo_from_rows(self.product_page_type_combo, type_rows)
            self.product_page_rows = {int(row["id"]): row for row in rows}
            full_table_rows = [
                [
                    row.get("status_produto", ""),
                    int(row["id"]),
                    row.get("empresa_nome", ""),
                    row.get("fornecedor_nome", ""),
                    row.get("fornecedor_uf", ""),
                    row.get("tipo_produto", ""),
                    row.get("codigo_fornecedor", ""),
                    row.get("codigo_empresa", ""),
                    row.get("descricao", ""),
                    row.get("ean", ""),
                    row.get("ncm", ""),
                    row.get("cest", ""),
                    row.get("origem_entrada", ""),
                    row.get("cst_icms", ""),
                    row.get("reducao_bc_icms", ""),
                    row.get("cfop_saida_fornecedor", ""),
                    row.get("aliquota_icms", ""),
                    row.get("cfop_entrada", ""),
                    row.get("cst_ipi", ""),
                    row.get("aliquota_ipi", ""),
                    row.get("cst_pis", ""),
                    row.get("cst_pis_cofins", ""),
                    row.get("aliquota_pis", ""),
                    row.get("cst_cofins", ""),
                    row.get("aliquota_cofins", ""),
                    row.get("natureza_receita_entrada", ""),
                    row.get("mva", ""),
                    row.get("valor_icms_st", ""),
                    row.get("c_classtrib", ""),
                    row.get("c_benef", ""),
                    row.get("origem_saida", ""),
                    row.get("cst_icms_saida", ""),
                    row.get("cfop_saida_empresa", ""),
                    row.get("aliquota_icms_saida", ""),
                    row.get("cst_pis_saida", ""),
                    row.get("cst_cofins_saida", ""),
                    row.get("natureza_receita_saida", ""),
                    row.get("chave_nfe_origem", ""),
                ]
                for row in rows
            ]
            self.product_page_search_text_by_id = {
                int(row_values[1]): " ".join(str(value).lower() for value in row_values)
                for row_values in full_table_rows
            }
            self.product_page_export_rows_by_id = {
                int(row_values[1]): row_values
                for row_values in full_table_rows
            }
            self.product_page_display_rows_by_id = {
                int(row_values[1]): row_values[:12]
                for row_values in full_table_rows
            }
            self.product_page_all_display_rows = [row_values[:12] for row_values in full_table_rows]
            self.populate_product_filter_combos(rows)
            self.apply_product_page_filters()
            pending_select_id = int(getattr(self, "_product_page_pending_select_id", 0) or 0)
            if pending_select_id:
                self.select_product_page_row_by_id(pending_select_id)
                self._product_page_pending_select_id = 0
            self.statusBar().showMessage("Produtos atualizados.")
            if hasattr(self, "product_page_refresh_button"):
                self.product_page_refresh_button.setEnabled(True)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
            self.statusBar().showMessage("Falha na atualizacao de produtos.")
            if hasattr(self, "product_page_refresh_button"):
                self.product_page_refresh_button.setEnabled(True)

    def handle_product_page_refresh_failed(self, error: str) -> None:
        QMessageBox.critical(self, "Produtos", error)
        self.statusBar().showMessage("Falha na atualizacao de produtos.")
        if hasattr(self, "product_page_refresh_button"):
            self.product_page_refresh_button.setEnabled(True)
        QApplication.processEvents()

    def _start_product_page_table_render(self, table_rows: list[list[object]], source_rows: list[dict[str, object]]) -> None:
        table = self.product_page_table
        self._product_render_rows = table_rows
        self._product_render_source_rows = source_rows
        self._product_render_index = 0
        self._product_render_chunk_size = 100
        self._product_render_header_labels = [
            (table.horizontalHeaderItem(index).text().strip().lower() if table.horizontalHeaderItem(index) else "")
            for index in range(table.columnCount())
        ]
        self._product_render_sorting = table.isSortingEnabled()
        if self._product_render_sorting:
            table.setSortingEnabled(False)
        table.setUpdatesEnabled(False)
        table.setRowCount(len(table_rows))
        self._product_rows_all_visible = True
        QTimer.singleShot(0, self._continue_product_page_table_render)

    def _continue_product_page_table_render(self) -> None:
        table = self.product_page_table
        rows = getattr(self, "_product_render_rows", [])
        header_labels = getattr(self, "_product_render_header_labels", [])
        start_index = int(getattr(self, "_product_render_index", 0))
        chunk_size = int(getattr(self, "_product_render_chunk_size", 100))
        end_index = min(start_index + chunk_size, len(rows))
        for row_index in range(start_index, end_index):
            row = rows[row_index]
            for column_index, value in enumerate(row):
                display_value = str(value)
                header_label = header_labels[column_index] if column_index < len(header_labels) else ""
                if header_label == "cnpj":
                    display_value = self.format_cnpj(display_value)
                elif header_label == "ie":
                    display_value = self.format_ie(display_value)
                item = QTableWidgetItem(display_value)
                if header_label in {"descricao", "produto"}:
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                table.setItem(row_index, column_index, item)
        self._product_render_index = end_index
        if end_index < len(rows):
            self.statusBar().showMessage(f"Atualizando produtos... ({end_index}/{len(rows)})")
            QTimer.singleShot(1, self._continue_product_page_table_render)
            return
        try:
            self.populate_product_filter_combos(getattr(self, "_product_render_source_rows", []))
            self.apply_product_page_filters()
            pending_select_id = int(getattr(self, "_product_page_pending_select_id", 0) or 0)
            if pending_select_id:
                self.select_product_page_row_by_id(pending_select_id)
                self._product_page_pending_select_id = 0
            self.statusBar().showMessage("Produtos atualizados.")
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
            self.statusBar().showMessage("Falha na atualizacao de produtos.")
        finally:
            table.setUpdatesEnabled(True)
            if bool(getattr(self, "_product_render_sorting", False)):
                table.setSortingEnabled(True)
            self._product_render_rows = []
            self._product_render_source_rows = []
            self._product_render_header_labels = []
            self._product_render_index = 0
            if hasattr(self, "product_page_refresh_button"):
                self.product_page_refresh_button.setEnabled(True)
            QApplication.processEvents()

    def populate_product_filter_combos(self, rows: list[dict[str, object]]) -> None:
        selected_company_id = int(self.product_page_filter_company_combo.currentData() or 0)
        selected_supplier_id = int(self.product_page_filter_supplier_combo.currentData() or 0)

        companies: dict[int, dict[str, object]] = {}
        suppliers: dict[int, tuple[str, int]] = {}
        for row in rows:
            company_id = int(row.get("empresa_id") or 0)
            supplier_id = int(row.get("fornecedor_id") or 0)
            if company_id:
                companies[company_id] = {"empresa_nome": row.get("empresa_nome", ""), "empresa_cnpj": row.get("empresa_cnpj", "")}
            if supplier_id:
                suppliers[supplier_id] = (str(row.get("fornecedor_nome", "")), company_id)

        self.product_page_filter_company_combo.blockSignals(True)
        self.product_page_filter_company_combo.clear()
        self.product_page_filter_company_combo.addItem("Todas Empresas", 0)
        for company_id, company_row in sorted(companies.items(), key=lambda item: str(item[1].get("empresa_nome", "")).upper()):
            self.product_page_filter_company_combo.addItem(self.company_display_name(company_row), company_id)
        self.select_combo_data(self.product_page_filter_company_combo, selected_company_id)
        self.product_page_filter_company_combo.blockSignals(False)

        self.product_page_suppliers_filter_index = suppliers
        self.populate_product_supplier_filter_options(selected_company_id, selected_supplier_id)

    def populate_product_supplier_filter_options(self, company_id: int, selected_supplier_id: int = 0) -> None:
        suppliers = getattr(self, "product_page_suppliers_filter_index", {})
        self.product_page_filter_supplier_combo.blockSignals(True)
        self.product_page_filter_supplier_combo.clear()
        self.product_page_filter_supplier_combo.addItem("Todos Fornecedores", 0)
        for supplier_id, (supplier_name, supplier_company_id) in sorted(suppliers.items(), key=lambda item: item[1][0].upper()):
            if company_id and supplier_company_id != company_id:
                continue
            self.product_page_filter_supplier_combo.addItem(supplier_name, supplier_id)
        self.select_combo_data(self.product_page_filter_supplier_combo, selected_supplier_id)
        self.product_page_filter_supplier_combo.blockSignals(False)

    def on_product_filter_company_changed(self, _index: int) -> None:
        selected_company_id = int(self.product_page_filter_company_combo.currentData() or 0)
        self.populate_product_supplier_filter_options(selected_company_id)
        self.apply_product_page_filters()

    def apply_product_page_filters(self) -> None:
        selected_company_id = int(self.product_page_filter_company_combo.currentData() or 0)
        selected_supplier_id = int(self.product_page_filter_supplier_combo.currentData() or 0)
        search_text = self.product_page_search_input.text().strip().lower() if hasattr(self, "product_page_search_input") else ""
        xml_only = hasattr(self, "product_page_xml_only_check") and self.product_page_xml_only_check.isChecked()
        display_rows_by_id = getattr(self, "product_page_display_rows_by_id", {})
        visible_ids: list[int] = []
        visible_rows: list[list[object]] = []
        if not selected_company_id and not selected_supplier_id and not search_text and not xml_only:
            self._product_rows_all_visible = True
            visible_rows = list(getattr(self, "product_page_all_display_rows", []))
            visible_ids = [int(row_values[1]) for row_values in visible_rows if len(row_values) > 1 and str(row_values[1]).isdigit()]
            self.set_product_page_visible_rows(visible_rows, visible_ids)
            self.refresh_product_metrics()
            return
        self._product_rows_all_visible = False
        for row_id, row in getattr(self, "product_page_rows", {}).items():
            row_company_id = int(row.get("empresa_id") or 0)
            row_supplier_id = int(row.get("fornecedor_id") or 0)
            company_ok = not selected_company_id or row_company_id == selected_company_id
            supplier_ok = not selected_supplier_id or row_supplier_id == selected_supplier_id
            row_text = getattr(self, "product_page_search_text_by_id", {}).get(row_id, "")
            text_ok = not search_text or search_text in row_text
            xml_ok = not xml_only or bool(str(row.get("chave_nfe_origem") or "").strip())
            if company_ok and supplier_ok and text_ok and xml_ok and row_id in display_rows_by_id:
                visible_ids.append(row_id)
                visible_rows.append(display_rows_by_id[row_id])
        self.set_product_page_visible_rows(visible_rows, visible_ids)
        self.refresh_product_metrics()

    def set_product_page_visible_rows(self, rows: list[list[object]], row_ids: list[int]) -> None:
        model = self.product_page_table.model()
        if isinstance(model, ProductTableModel):
            model.set_rows(rows)
            self.apply_product_table_column_widths(self.product_page_table)
        self.product_page_visible_ids = row_ids

    def refresh_product_metrics(self) -> None:
        if not hasattr(self, "product_metric_labels"):
            return
        visible_ids = list(getattr(self, "product_page_visible_ids", []))
        total_products = len(visible_ids)
        class_counter: dict[str, int] = {}
        for row_id in visible_ids:
            row = getattr(self, "product_page_rows", {}).get(row_id, {})
            class_name = str(row.get("tipo_produto", "") or "").strip() or "Sem classificacao"
            class_counter[class_name] = class_counter.get(class_name, 0) + 1
        total_classes = len(class_counter)
        top_class = "-"
        if class_counter:
            class_name, qty = max(class_counter.items(), key=lambda item: item[1])
            top_class = f"{class_name} ({qty})"
        selected_supplier_id = int(self.product_page_filter_supplier_combo.currentData() or 0) if hasattr(self, "product_page_filter_supplier_combo") else 0
        selected_company_id = int(self.product_page_filter_company_combo.currentData() or 0) if hasattr(self, "product_page_filter_company_combo") else 0
        scope_text = "Geral"
        if selected_supplier_id:
            scope_text = self.product_page_filter_supplier_combo.currentText().strip() or "Fornecedor"
        elif selected_company_id:
            scope_text = self.product_page_filter_company_combo.currentText().strip() or "Empresa"
        self.product_metric_labels["products"].setText(str(total_products))
        self.product_metric_labels["classes"].setText(str(total_classes))
        self.product_metric_labels["top_class"].setText(top_class)
        self.product_metric_labels["scope"].setText(scope_text)

    def handle_product_page_select(self) -> None:
        row = getattr(self, "product_page_rows", {}).get(self.selected_product_page_id(), {})
        self.fill_line_fields(self.product_page_fields, row)
        self.select_combo_data(self.product_page_supplier_combo, int(row.get("fornecedor_id") or 0))
        type_id = int(row.get("tipo_produto_id") or 0)
        if type_id:
            self.select_combo_data(self.product_page_type_combo, type_id)
            return
        selected_rows = self.product_page_table.selectionModel().selectedRows()
        if not selected_rows:
            self.product_page_type_combo.setCurrentIndex(0)
            return
        classification_text = str(row.get("tipo_produto", "") or "").strip()
        combo_index = self.product_page_type_combo.findText(classification_text)
        self.product_page_type_combo.setCurrentIndex(combo_index if combo_index >= 0 else 0)

    def edit_product_page_selected(self) -> None:
        self.confirm_selected_for_edit(self.product_page_table)
        if self.selected_product_page_id():
            self.product_page_tabs.setCurrentIndex(1)

    def clear_product_page_form(self) -> None:
        self.fill_line_fields(self.product_page_fields, {})
        self.product_page_type_combo.setCurrentIndex(0)
        if hasattr(self, "product_page_tabs"):
            self.product_page_tabs.setCurrentIndex(1)

    def clear_product_page_consultation(self) -> None:
        if hasattr(self, "product_page_filter_company_combo"):
            self.product_page_filter_company_combo.setCurrentIndex(0)
        if hasattr(self, "product_page_filter_supplier_combo"):
            self.product_page_filter_supplier_combo.setCurrentIndex(0)
        if hasattr(self, "product_page_search_input"):
            self.product_page_search_input.clear()
        if hasattr(self, "product_page_xml_only_check"):
            self.product_page_xml_only_check.setChecked(False)
        self.apply_product_page_filters()

    def save_product_page(self) -> None:
        try:
            supplier_id = int(self.product_page_supplier_combo.currentData() or 0)
            payload = self.line_field_payload(self.product_page_fields)
            row_before = int(payload.get("id") or 0)
            payload["tipo_produto_id"] = self.product_page_type_combo.currentData()
            row_id = self.mysql_repo.save_supplier_product(supplier_id, payload)
            self._product_page_pending_select_id = row_id
            self.start_refresh_product_page_fast()
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Produtos", f"Produto {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))

    def open_planilha_import(self) -> None:
        dialog = ImportPlanilhaDialog(self.mysql_repo, self.environment, self)
        dialog.exec()
        self.start_refresh_product_page_fast()

    def delete_product_page(self) -> None:
        row_id = self.selected_product_page_id()
        if row_id and QMessageBox.question(self, "Produtos", "Excluir este produto?") == QMessageBox.Yes:
            self.mysql_repo.delete_supplier_product(row_id)
            self.start_refresh_product_page_fast()

    def selected_product_page_id(self) -> int:
        selected_rows = self.product_page_table.selectionModel().selectedRows()
        if not selected_rows:
            return 0
        model = self.product_page_table.model()
        if isinstance(model, ProductTableModel):
            return model.row_id(selected_rows[0].row())
        return 0

    def select_product_page_row_by_id(self, row_id: int) -> None:
        model = self.product_page_table.model()
        if not isinstance(model, ProductTableModel):
            return
        for row_index in range(model.rowCount()):
            if model.row_id(row_index) == row_id:
                self.product_page_table.selectRow(row_index)
                return

    def select_combo_data(self, combo: QComboBox, value: int) -> None:
        index = combo.findData(value)
        combo.setCurrentIndex(index if index >= 0 else 0)

    def normalize_regime_tributario_value(self, value: object) -> str:
        normalized = str(value or "").strip().upper()
        if normalized in {"SIMPLES_NACIONAL", "SIMPLES NACIONAL"}:
            return "SIMPLES_NACIONAL"
        return "LUCRO_REAL_PRESUMIDO"

    def regime_tributario_label(self, value: str) -> str:
        normalized = self.normalize_regime_tributario_value(value)
        return "Simples Nacional" if normalized == "SIMPLES_NACIONAL" else "Lucro Real/Presumido"

    def select_regime_tributario_combo(self, combo: QComboBox, value: str) -> None:
        normalized = self.normalize_regime_tributario_value(value)
        index = combo.findData(normalized)
        combo.setCurrentIndex(index if index >= 0 else 0)

    def filter_table_rows(self, table: QTableWidget, text: str) -> None:
        normalized = text.strip().lower()
        for row_index in range(table.rowCount()):
            row_text = " ".join(
                table.item(row_index, column_index).text().lower()
                for column_index in range(table.columnCount())
                if table.item(row_index, column_index) is not None
            )
            table.setRowHidden(row_index, normalized not in row_text)

    def confirm_selected_for_edit(self, table: QTableWidget) -> None:
        selected = table.selectionModel().selectedRows()
        if not selected:
            QMessageBox.information(self, "Editar Cadastro", "Selecione uma linha na lista para editar.")
            return
        table.setFocus()
        self.statusBar().showMessage("Registro selecionado. Edite os campos do formulario e clique em Salvar Alteracoes.")

    def edit_single_catalog_selected(self, table: QTableWidget) -> None:
        self.confirm_selected_for_edit(table)
        if self.selected_catalog_id(table):
            tabs = getattr(self, "catalog_page_tabs_by_table", {}).get(id(table))
            if tabs is not None:
                tabs.setCurrentIndex(1)

    def clear_single_catalog_form(self, table: QTableWidget, clear_callback: Callable[[], None]) -> None:
        clear_callback()
        tabs = getattr(self, "catalog_page_tabs_by_table", {}).get(id(table))
        if tabs is not None:
            tabs.setCurrentIndex(1)

    def clear_catalog_consultation(self, table: QTableWidget) -> None:
        search_input = getattr(self, "catalog_page_search_inputs_by_table", {}).get(id(table))
        if search_input is not None:
            search_input.clear()
        for row_index in range(table.rowCount()):
            table.setRowHidden(row_index, False)

    def refresh_catalog(self) -> None:
        if not hasattr(self, "catalog_company_table"):
            return
        try:
            self.mysql_repo.ensure_schema()
            companies = self.mysql_repo.list_companies(self.environment)
            self.catalog_companies_by_id = {int(row["id"]): row for row in companies}
            self.set_table_rows(
            self.catalog_company_table,
            [[int(row["id"]), row.get("nome", ""), row.get("cnpj", ""), row.get("inscricao_estadual", "")] for row in companies],
        )
            if companies:
                self.catalog_company_table.selectRow(0)
                self.handle_catalog_company_select()
            else:
                self.refresh_catalog_children(0)
            self.catalog_status.setText(f"{len(companies)} empresa(s) cadastrada(s).")
        except Exception as exc:
            self.catalog_status.setText(f"Falha no cadastro: {exc}")
            QMessageBox.critical(self, "Cadastro Empresa", str(exc))

    def refresh_catalog_children(self, company_id: int) -> None:
        suppliers = self.mysql_repo.list_suppliers(company_id) if company_id else []
        types = self.mysql_repo.list_product_types(self.environment)
        self.catalog_suppliers_by_id = {int(row["id"]): row for row in suppliers}
        self.catalog_types_by_id = {int(row["id"]): row for row in types}
        self.set_table_rows(
            self.catalog_supplier_table,
            [
                [
                    int(row["id"]),
                    row.get("nome", ""),
                    row.get("cnpj", ""),
                    row.get("inscricao_estadual", ""),
                    row.get("uf", ""),
                    row.get("codigo", ""),
                    self.regime_tributario_label(str(row.get("regime_tributario", ""))),
                ]
                for row in suppliers
            ],
        )
        self.set_table_rows(
            self.catalog_type_table,
            [[int(row["id"]), row.get("nome", ""), row.get("descricao", "")] for row in types],
        )
        self.catalog_product_type_combo.clear()
        self.catalog_product_type_combo.addItem("", 0)
        for type_id, row in self.catalog_types_by_id.items():
            self.catalog_product_type_combo.addItem(str(row.get("nome", "")), type_id)
        self.set_table_rows(self.catalog_product_table, [])
        self.clear_catalog_supplier_form()
        self.clear_catalog_type_form()
        self.clear_catalog_product_form()

    def handle_catalog_company_select(self) -> None:
        company_id = self.selected_catalog_id(self.catalog_company_table)
        row = self.catalog_companies_by_id.get(company_id, {})
        self.fill_line_fields(self.catalog_company_fields, row)
        self.refresh_catalog_children(company_id)

    def handle_catalog_supplier_select(self) -> None:
        supplier_id = self.selected_catalog_id(self.catalog_supplier_table)
        row = self.catalog_suppliers_by_id.get(supplier_id, {})
        self.fill_line_fields(self.catalog_supplier_fields, row)
        self.refresh_catalog_products(supplier_id)

    def handle_catalog_type_select(self) -> None:
        type_id = self.selected_catalog_id(self.catalog_type_table)
        self.fill_line_fields(self.catalog_type_fields, self.catalog_types_by_id.get(type_id, {}))

    def refresh_catalog_products(self, supplier_id: int) -> None:
        products = self.mysql_repo.list_supplier_products(supplier_id) if supplier_id else []
        self.catalog_products_by_id = {int(row["id"]): row for row in products}
        self.set_table_rows(
            self.catalog_product_table,
            [
                [
                    int(row["id"]),
                    row.get("codigo_fornecedor", ""),
                    row.get("codigo_empresa", ""),
                    row.get("descricao", ""),
                    row.get("tipo_produto", ""),
                    row.get("ncm", ""),
                    row.get("origem_entrada", ""),
                    row.get("cst_icms", ""),
                    row.get("reducao_bc_icms", ""),
                    row.get("cfop_saida_fornecedor", ""),
                    row.get("aliquota_icms", ""),
                    row.get("cfop_entrada", ""),
                    row.get("cst_ipi", ""),
                    row.get("aliquota_ipi", ""),
                    row.get("cst_pis", ""),
                    row.get("cst_pis_cofins", ""),
                    row.get("aliquota_pis", ""),
                    row.get("cst_cofins", ""),
                    row.get("aliquota_cofins", ""),
                    row.get("natureza_receita_entrada", ""),
                    row.get("mva", ""),
                    row.get("valor_icms_st", ""),
                    row.get("c_classtrib", ""),
                    row.get("c_benef", ""),
                    row.get("origem_saida", ""),
                    row.get("cst_icms_saida", ""),
                    row.get("cfop_saida_empresa", ""),
                    row.get("aliquota_icms_saida", ""),
                    row.get("cst_pis_saida", ""),
                    row.get("cst_cofins_saida", ""),
                    row.get("natureza_receita_saida", ""),
                ]
                for row in products
            ],
        )
        self.clear_catalog_product_form()

    def handle_catalog_product_select(self) -> None:
        product_id = self.selected_catalog_id(self.catalog_product_table)
        row = self.catalog_products_by_id.get(product_id, {})
        self.fill_line_fields(self.catalog_product_fields, row)
        type_id = int(row.get("tipo_produto_id") or 0)
        index = self.catalog_product_type_combo.findData(type_id)
        self.catalog_product_type_combo.setCurrentIndex(index if index >= 0 else 0)

    def fill_line_fields(self, fields: dict[str, QLineEdit], row: dict[str, object]) -> None:
        for key, field in fields.items():
            field.setText(str(row.get(key, "")))

    def line_field_payload(self, fields: dict[str, QLineEdit]) -> dict[str, object]:
        return {key: field.text().strip() for key, field in fields.items()}

    def digits_only(self, value: str) -> str:
        return "".join(char for char in str(value or "") if char.isdigit())

    def is_valid_cnpj(self, cnpj: str) -> bool:
        digits = self.digits_only(cnpj)
        if len(digits) != 14:
            return False
        if digits == digits[0] * 14:
            return False

        def calc_digit(base: str, multipliers: list[int]) -> int:
            total = sum(int(number) * weight for number, weight in zip(base, multipliers))
            remainder = total % 11
            return 0 if remainder < 2 else 11 - remainder

        first = calc_digit(digits[:12], [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
        second = calc_digit(digits[:12] + str(first), [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
        return digits.endswith(f"{first}{second}")

    def validate_required_cnpj(self, cnpj: str, title: str) -> bool:
        if not self.digits_only(cnpj):
            QMessageBox.warning(self, title, "CNPJ obrigatorio. Informe os 14 digitos.")
            return False
        if self.is_valid_cnpj(cnpj):
            return True
        QMessageBox.warning(self, title, "CNPJ informado e invalido. Confira os 14 digitos.")
        return False

    def clear_catalog_company_form(self) -> None:
        self.fill_line_fields(self.catalog_company_fields, {})

    def clear_catalog_supplier_form(self) -> None:
        self.fill_line_fields(self.catalog_supplier_fields, {})

    def clear_catalog_type_form(self) -> None:
        self.fill_line_fields(self.catalog_type_fields, {})

    def clear_catalog_product_form(self) -> None:
        self.fill_line_fields(self.catalog_product_fields, {})
        self.catalog_product_type_combo.setCurrentIndex(0)

    def save_catalog_company(self) -> None:
        try:
            payload = self.line_field_payload(self.catalog_company_fields)
            row_before = int(payload.get("id") or 0)
            if not self.validate_required_cnpj(str(payload.get("cnpj", "")), "Cadastro Empresa"):
                return
            company_id = self.mysql_repo.save_company(self.environment, payload)
            self.refresh_catalog()
            self.select_table_row_by_id(self.catalog_company_table, company_id)
            action = "atualizada" if row_before else "cadastrada"
            QMessageBox.information(self, "Cadastro Empresa", f"Empresa {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Cadastro Empresa", str(exc))

    def save_catalog_supplier(self) -> None:
        try:
            payload = self.line_field_payload(self.catalog_supplier_fields)
            row_before = int(payload.get("id") or 0)
            if not self.validate_required_cnpj(str(payload.get("cnpj", "")), "Cadastro Fornecedor"):
                return
            payload["regime_tributario"] = self.normalize_regime_tributario_value(payload.get("regime_tributario", ""))
            company_id = self.selected_catalog_id(self.catalog_company_table)
            supplier_id = self.mysql_repo.save_supplier(company_id, payload)
            self.refresh_catalog_children(company_id)
            self.select_table_row_by_id(self.catalog_supplier_table, supplier_id)
            self.handle_catalog_supplier_select()
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Cadastro Fornecedor", f"Fornecedor {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Cadastro Fornecedor", str(exc))

    def save_catalog_type(self) -> None:
        try:
            company_id = self.selected_catalog_id(self.catalog_company_table)
            payload = self.line_field_payload(self.catalog_type_fields)
            row_before = int(payload.get("id") or 0)
            type_id = self.mysql_repo.save_product_type(self.environment, payload)
            self.refresh_catalog_children(company_id)
            self.select_table_row_by_id(self.catalog_type_table, type_id)
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Classificacao do Produto", f"Classificacao do produto {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Classificacao do Produto", str(exc))

    def save_catalog_product(self) -> None:
        try:
            supplier_id = self.selected_catalog_id(self.catalog_supplier_table)
            payload = self.line_field_payload(self.catalog_product_fields)
            row_before = int(payload.get("id") or 0)
            payload["tipo_produto_id"] = self.catalog_product_type_combo.currentData()
            product_id = self.mysql_repo.save_supplier_product(supplier_id, payload)
            self.refresh_catalog_products(supplier_id)
            self.select_table_row_by_id(self.catalog_product_table, product_id)
            action = "atualizado" if row_before else "cadastrado"
            QMessageBox.information(self, "Cadastro Produto", f"Produto {action} com sucesso.")
        except Exception as exc:
            QMessageBox.critical(self, "Cadastro Produto", str(exc))

    def delete_catalog_company(self) -> None:
        company_id = self.selected_catalog_id(self.catalog_company_table)
        if company_id and QMessageBox.question(self, "Cadastro Empresa", "Excluir a empresa e seus fornecedores/produtos?") == QMessageBox.Yes:
            self.mysql_repo.delete_company(company_id)
            self.refresh_catalog()

    def delete_catalog_supplier(self) -> None:
        supplier_id = self.selected_catalog_id(self.catalog_supplier_table)
        if supplier_id and QMessageBox.question(self, "Cadastro Fornecedor", "Excluir o fornecedor e seus produtos?") == QMessageBox.Yes:
            self.mysql_repo.delete_supplier(supplier_id)
            self.refresh_catalog_children(self.selected_catalog_id(self.catalog_company_table))

    def delete_catalog_type(self) -> None:
        type_id = self.selected_catalog_id(self.catalog_type_table)
        if type_id and QMessageBox.question(self, "Classificacao do Produto", "Excluir esta classificacao? Produtos vinculados ficarao sem classificacao.") == QMessageBox.Yes:
            self.mysql_repo.delete_product_type(type_id)
            self.refresh_catalog_children(self.selected_catalog_id(self.catalog_company_table))

    def delete_catalog_product(self) -> None:
        product_id = self.selected_catalog_id(self.catalog_product_table)
        if product_id and QMessageBox.question(self, "Cadastro Produto", "Excluir este produto?") == QMessageBox.Yes:
            self.mysql_repo.delete_supplier_product(product_id)
            self.refresh_catalog_products(self.selected_catalog_id(self.catalog_supplier_table))

    def select_table_row_by_id(self, table: QTableWidget, row_id: int) -> None:
        id_column = self.get_table_id_column_index(table)
        for row_index in range(table.rowCount()):
            item = table.item(row_index, id_column)
            if item and item.text() == str(row_id):
                table.selectRow(row_index)
                return

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

    def build_duplicates_cleanup_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Limpeza de Duplicados")
        title.setObjectName("sectionTitle")
        self.duplicates_status = QLabel("Analise e remova duplicados de Codigo/EAN por fornecedor.")
        self.duplicates_status.setObjectName("muted")
        toolbar_layout.addWidget(title)
        toolbar_layout.addWidget(self.duplicates_status, 1)
        toolbar_layout.addWidget(self.create_button("Atualizar", self.refresh_duplicates_cleanup_page))
        toolbar_layout.addWidget(self.create_button("Excluir Selecionados", self.delete_selected_duplicates))
        toolbar_layout.addWidget(self.create_button("Limpar Automatico", self.delete_duplicates_automatic, primary=True))
        layout.addWidget(toolbar)
        self.duplicates_table = self.create_data_table(["ID", "Empresa", "Fornecedor", "Tipo", "Cod. Forn.", "EAN", "Descricao"])
        self.duplicates_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.duplicates_table.setSelectionMode(QAbstractItemView.MultiSelection)
        layout.addWidget(self.duplicates_table, 1)
        return page

    def build_nfe_key_extract_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar = self.create_panel()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        title = QLabel("Extrair chave NFe")
        title.setObjectName("sectionTitle")
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.create_button("Limpar", self.clear_nfe_key_extract_page))
        toolbar_layout.addWidget(self.create_button("Exportar", self.export_nfe_key_extract_rows))
        toolbar_layout.addWidget(self.create_button("Carregar", self.process_nfe_key_extract, primary=True))
        layout.addWidget(toolbar)

        setup = self.create_panel()
        setup_layout = QGridLayout(setup)
        setup_layout.setContentsMargins(12, 12, 12, 12)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(8)
        self.nfe_extract_sped_input = QLineEdit()
        self.nfe_extract_sped_input.setPlaceholderText("Selecione um ou mais arquivos SPED (.txt, .sped, .efd)")
        setup_layout.addWidget(QLabel("SPEDs"), 0, 0)
        setup_layout.addWidget(self.nfe_extract_sped_input, 0, 1, 1, 5)
        setup_layout.addWidget(self.create_button("Selecionar SPEDs", self.select_nfe_extract_sped_files), 1, 1)
        setup_layout.addWidget(self.create_button("Selecionar Pasta", self.select_nfe_extract_sped_folder), 1, 2)
        setup_layout.addWidget(self.create_button("Limpar SPEDs", lambda: self.nfe_extract_sped_input.clear()), 1, 3)
        self.nfe_extract_scope_combo = QComboBox()
        self.nfe_extract_scope_combo.addItems(["Todos", "Entrada", "Saida"])
        self.nfe_extract_scope_combo.currentTextChanged.connect(self.apply_nfe_key_filter)
        setup_layout.addWidget(QLabel("Tipo de documento"), 1, 4)
        setup_layout.addWidget(self.nfe_extract_scope_combo, 1, 5)
        setup_layout.addWidget(self.create_button("Carregar", self.process_nfe_key_extract, primary=True), 1, 6)
        setup_layout.setColumnStretch(1, 1)
        self.nfe_extract_status = QLabel("Selecione os SPEDs, escolha o tipo e clique em Carregar.")
        self.nfe_extract_status.setObjectName("muted")
        setup_layout.addWidget(self.nfe_extract_status, 2, 0, 1, 6)
        self.nfe_extract_progress = QProgressBar()
        self.nfe_extract_progress.setMinimum(0)
        self.nfe_extract_progress.setMaximum(100)
        self.nfe_extract_progress.setValue(0)
        setup_layout.addWidget(self.nfe_extract_progress, 3, 0, 1, 6)
        layout.addWidget(setup)

        table_panel = self.create_panel()
        table_layout = QVBoxLayout(table_panel)
        table_layout.setContentsMargins(12, 12, 12, 12)
        table_layout.setSpacing(8)
        table_layout.addWidget(QLabel("Chaves carregadas"))
        self.nfe_extract_table = self.create_data_table(
            ["Tipo Documento", "Operacao", "Chave", "Documento", "Serie", "Data", "Participante", "CNPJ/CPF", "Arquivo SPED"]
        )
        table_layout.addWidget(self.nfe_extract_table, 1)
        layout.addWidget(table_panel, 1)
        return page

    def refresh_duplicates_cleanup_page(self) -> None:
        rows = self.mysql_repo.list_supplier_product_duplicates_catalog(self.environment)
        unique: dict[int, dict[str, object]] = {}
        for row in rows:
            unique[int(row["id"])] = row
        ordered = sorted(unique.values(), key=lambda item: (str(item.get("empresa_nome", "")), str(item.get("fornecedor_nome", "")), int(item.get("id", 0))))
        self.duplicates_rows_by_id = {int(row["id"]): row for row in ordered}
        self.set_table_rows(
            self.duplicates_table,
            [[int(row["id"]), row.get("empresa_nome", ""), row.get("fornecedor_nome", ""), row.get("duplicate_type", ""), row.get("codigo_fornecedor", ""), row.get("ean", ""), row.get("descricao", "")] for row in ordered],
        )
        self.duplicates_status.setText(f"Duplicados encontrados: {len(ordered)} item(ns).")

    def delete_selected_duplicates(self) -> None:
        selected = self.duplicates_table.selectionModel().selectedRows()
        ids = [self.selected_catalog_id(self.duplicates_table)] if len(selected) == 1 else []
        if len(selected) > 1:
            ids = []
            for row in selected:
                item = self.duplicates_table.item(row.row(), 0)
                if item and item.text().isdigit():
                    ids.append(int(item.text()))
        if not ids:
            QMessageBox.information(self, "Limpeza Duplicados", "Selecione ao menos uma linha.")
            return
        if QMessageBox.question(self, "Limpeza Duplicados", f"Excluir {len(ids)} item(ns) selecionado(s)?") != QMessageBox.Yes:
            return
        removed = self.mysql_repo.delete_supplier_products_by_ids(ids)
        self.refresh_duplicates_cleanup_page()
        QMessageBox.information(self, "Limpeza Duplicados", f"Removidos: {removed} item(ns).")

    def delete_duplicates_automatic(self) -> None:
        rows = self.mysql_repo.list_supplier_product_duplicates_catalog(self.environment)
        if not rows:
            QMessageBox.information(self, "Limpeza Duplicados", "Nao ha duplicados para limpar.")
            return
        grouped: dict[tuple[int, str, str], list[int]] = {}
        for row in rows:
            key = (int(row.get("fornecedor_id", 0)), str(row.get("duplicate_type", "")), str(row.get("codigo_fornecedor", "") if row.get("duplicate_type") == "codigo_fornecedor" else row.get("ean", "")))
            grouped.setdefault(key, []).append(int(row["id"]))
        to_delete: list[int] = []
        for _key, ids in grouped.items():
            ordered = sorted(set(ids))
            to_delete.extend(ordered[1:])
        if not to_delete:
            QMessageBox.information(self, "Limpeza Duplicados", "Nao ha itens excedentes para excluir.")
            return
        if QMessageBox.question(self, "Limpeza Duplicados", f"Excluir automaticamente {len(to_delete)} duplicado(s), mantendo o menor ID de cada grupo?") != QMessageBox.Yes:
            return
        removed = self.mysql_repo.delete_supplier_products_by_ids(sorted(set(to_delete)))
        self.refresh_duplicates_cleanup_page()
        QMessageBox.information(self, "Limpeza Duplicados", f"Limpeza automatica concluida. Removidos: {removed} item(ns).")

    def select_sped_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Fiscais", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        paths = append_unique_paths(parse_selected_paths(self.sped_input.text()), files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        self.sped_input.setText(format_selected_paths(paths))
        if limit_exceeded:
            QMessageBox.warning(self, "SPED Qt", "A consulta aceita no maximo 12 SPEDs.")

    def select_nfe_extract_sped_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Fiscais", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        current_paths = parse_selected_paths(self.nfe_extract_sped_input.text())
        paths = append_unique_paths(current_paths, files)
        self.nfe_extract_sped_input.setText(format_selected_paths(paths))

    def select_nfe_extract_sped_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta com SPEDs")
        if not directory:
            return
        folder_path = Path(directory)
        files = sorted(
            [
                file_path
                for file_path in folder_path.iterdir()
                if file_path.is_file() and file_path.suffix.lower() in {".txt", ".sped", ".efd"}
            ],
            key=lambda path: path.name.lower(),
        )
        if not files:
            QMessageBox.information(self, "Extrair chave NFe", "Nenhum arquivo SPED (.txt/.sped/.efd) foi encontrado na pasta selecionada.")
            return
        current_paths = parse_selected_paths(self.nfe_extract_sped_input.text())
        paths = append_unique_paths(current_paths, files)
        self.nfe_extract_sped_input.setText(format_selected_paths(paths))

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

    def select_entry_exit_sped_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar SPEDs Fiscais", "", "Arquivos SPED (*.txt *.sped *.efd);;Todos os arquivos (*.*)")
        if not files:
            return
        paths = append_unique_paths(parse_selected_paths(self.entry_exit_sped_input.text()), files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        self.entry_exit_sped_input.setText(format_selected_paths(paths))
        if limit_exceeded:
            QMessageBox.warning(self, "SPED Qt", "A analise aceita no maximo 12 SPEDs.")

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

    def select_catalog_import_xml_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar XMLs para importar", "", "Arquivos XML (*.xml);;Todos os arquivos (*.*)")
        if not files:
            return
        current_paths = parse_selected_paths(self.catalog_import_sources_input.text())
        collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(files)
        if collapsed_to_folder:
            current_paths = [path for path in current_paths if not path.is_file()]
        self.catalog_import_sources_input.setText(format_selected_paths(append_unique_paths(current_paths, collapsed_paths)))

    def select_catalog_import_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Selecionar pasta de XMLs para importar")
        if directory:
            current_paths = parse_selected_paths(self.catalog_import_sources_input.text())
            self.catalog_import_sources_input.setText(format_selected_paths(append_unique_paths(current_paths, [Path(directory)])))

    def import_catalogs_from_xml_sources(self) -> None:
        if getattr(self, "catalog_import_thread", None) is not None:
            QMessageBox.information(self, "Importacao XML Cadastros", "Ja existe uma operacao de importacao em andamento.")
            return
        if not getattr(self, "catalog_import_preview_rows", []):
            QMessageBox.information(self, "Importacao XML Cadastros", "Gere a previa antes de importar.")
            self.build_catalog_import_preview_ui()
            return
        preview_rows = getattr(self, "catalog_import_preview_rows", [])
        if not preview_rows:
            QMessageBox.warning(self, "Importacao XML Cadastros", "Nao ha itens na previa para importar.")
            return
        mode = self.ask_catalog_import_mode()
        if mode is None:
            return
        allow_insert_new, allow_update_existing = mode
        selected_update_fields = set(UPDATABLE_PRODUCT_FIELDS)
        if allow_update_existing:
            selected_update_fields = self.ask_catalog_import_update_fields()
            if not selected_update_fields:
                QMessageBox.information(self, "Importacao XML Cadastros", "Nenhum campo selecionado para atualizacao.")
                return
        self.catalog_import_progress.setValue(0)
        self.catalog_import_status.setText("Iniciando importacao dos cadastros...")
        self.statusBar().showMessage("Importando cadastros via XML...")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        self.catalog_import_thread = QThread(self)
        self.catalog_import_worker = CatalogImportWorker(
            self.mysql_repo,
            self.environment,
            preview_rows,
            allow_update_existing,
            allow_insert_new,
            selected_update_fields,
        )
        self.catalog_import_worker.moveToThread(self.catalog_import_thread)
        self.catalog_import_thread.started.connect(self.catalog_import_worker.run)
        self.catalog_import_worker.progress.connect(self.update_catalog_import_progress)
        self.catalog_import_worker.finished.connect(self.render_catalog_import_results)
        self.catalog_import_worker.failed.connect(self.handle_catalog_import_error)
        self.catalog_import_worker.finished.connect(self.catalog_import_thread.quit)
        self.catalog_import_worker.failed.connect(self.catalog_import_thread.quit)
        self.catalog_import_thread.finished.connect(self.catalog_import_worker.deleteLater)
        self.catalog_import_thread.finished.connect(self.catalog_import_thread.deleteLater)
        self.catalog_import_thread.finished.connect(lambda: setattr(self, "catalog_import_worker", None))
        self.catalog_import_thread.finished.connect(lambda: setattr(self, "catalog_import_thread", None))
        self.catalog_import_thread.start()

    def render_catalog_import_results(self, stats: dict[str, int], duplicate_rows: list[list[object]], change_rows: list[list[object]], backup_path: str) -> None:
        def limited_rows(rows: list[list[object]], headers: list[str]) -> list[list[object]]:
            if len(rows) <= CATALOG_IMPORT_POPUP_ROW_LIMIT:
                return rows
            hidden = len(rows) - CATALOG_IMPORT_POPUP_ROW_LIMIT
            message_row = [""] * max(1, len(headers))
            message_row[-1] = f"{hidden} linha(s) omitida(s) para preservar desempenho da tela."
            return [*rows[:CATALOG_IMPORT_POPUP_ROW_LIMIT], message_row]

        try:
            if backup_path:
                self.statusBar().showMessage(f"Backup de seguranca criado: {backup_path}")
            skipped_new = stats.get("products_skipped_new", 0)
            summary_parts = [
                f"Registros na previa: {stats['rows_total']}",
                f"Empresas processadas: {stats['companies_processed']} | Fornecedores processados: {stats['suppliers_processed']}",
                f"Produtos criados: {stats['products_created']} | atualizados: {stats['products_updated']} | existentes ignorados: {stats['products_skipped_existing']}",
            ]
            if skipped_new:
                summary_parts.append(f"Novos ignorados (modo somente atualizar): {skipped_new}")
            summary_parts.append("Classificacao do produto: em branco para selecionar depois")
            summary = "\n".join(summary_parts)
            self.catalog_import_status.setText(summary)
            self.catalog_import_skipped_product_rows = list(stats.get("skipped_rows", [])) if isinstance(stats, dict) else []
            self.statusBar().showMessage("Importacao de cadastros via XML concluida.")
            self.catalog_import_progress.setValue(100)
            QMessageBox.information(self, "Importacao XML Cadastros", summary)
            groups: list[tuple[str, list[str], list[list[object]]]] = []
            if duplicate_rows:
                headers = ["XML", "Fornecedor ID", "Codigo Forn.", "EAN", "Produto ID", "Descricao", "Motivo"]
                groups.append(("Duplicidades no cadastro", headers, limited_rows(duplicate_rows, headers)))
            if change_rows:
                headers = ["XML", "Codigo", "Campo", "Valor Atual", "Valor Novo", "Motivo"]
                groups.append(("Mudancas detectadas", headers, limited_rows(change_rows, headers)))
            if self.catalog_import_skipped_product_rows:
                ignored_rows = [row for row in self.catalog_import_skipped_product_rows if str(row[4]).strip().lower() == "ignorado"]
                unchanged_rows = [row for row in self.catalog_import_skipped_product_rows if str(row[4]).strip().lower() == "sem alteracao"]
                if ignored_rows:
                    headers = ["XML", "Codigo", "EAN", "Descricao", "Status", "Motivo"]
                    groups.append(("Ignorados", headers, limited_rows(ignored_rows, headers)))
                if unchanged_rows:
                    headers = ["XML", "Codigo", "EAN", "Descricao", "Status", "Motivo"]
                    groups.append(("Sem alteracao", headers, limited_rows(unchanged_rows, headers)))
            if groups:
                self.open_multi_table_popup("Resultado consolidado da importacao", groups, width=1560, height=720)
        finally:
            QApplication.restoreOverrideCursor()

    def handle_catalog_import_error(self, error_text: str) -> None:
        try:
            self.catalog_import_status.setText(f"Falha na importacao: {error_text}")
            self.statusBar().showMessage("Falha na importacao de cadastros via XML.")
            QMessageBox.critical(self, "Importacao XML Cadastros", error_text)
        finally:
            QApplication.restoreOverrideCursor()

    def ask_catalog_import_mode(self) -> tuple[bool, bool] | None:
        """Returns (allow_insert_new, allow_update_existing) or None if cancelled."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Modo de Importacao")
        dialog.setObjectName("popupDialog")
        self.resize_dialog_to_screen(dialog, 480, 260)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        title = QLabel("Como deseja importar os produtos encontrados nos XMLs?")
        title.setObjectName("sectionTitle")
        title.setWordWrap(True)
        layout.addWidget(title)

        group = QButtonGroup(dialog)
        rb_novos = QRadioButton("Somente cadastrar novos  (produtos existentes sao ignorados)")
        rb_atualizar = QRadioButton("Somente atualizar existentes  (novos produtos nao sao cadastrados)")
        rb_ambos = QRadioButton("Cadastrar novos e atualizar existentes")
        rb_novos.setChecked(True)
        group.addButton(rb_novos, 0)
        group.addButton(rb_atualizar, 1)
        group.addButton(rb_ambos, 2)

        panel = self.create_panel()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(14, 10, 14, 10)
        panel_layout.setSpacing(8)
        panel_layout.addWidget(rb_novos)
        panel_layout.addWidget(rb_atualizar)
        panel_layout.addWidget(rb_ambos)
        layout.addWidget(panel)

        actions = QHBoxLayout()
        ok_btn = self.create_button("Continuar", dialog.accept, primary=True)
        cancel_btn = self.create_button("Cancelar", dialog.reject)
        actions.addStretch()
        actions.addWidget(ok_btn)
        actions.addWidget(cancel_btn)
        layout.addLayout(actions)

        if dialog.exec() != QDialog.Accepted:
            return None
        selected = group.checkedId()
        if selected == 0:
            return (True, False)   # insert only
        if selected == 1:
            return (False, True)   # update only
        return (True, True)        # both

    def ask_catalog_import_update_fields(self) -> set[str]:
        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar Campos para Atualizacao")
        dialog.setObjectName("popupDialog")
        self.resize_dialog_to_screen(dialog, 700, 520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        title = QLabel("Escolha os campos que podem ser atualizados nos itens existentes:")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        checks: dict[str, QCheckBox] = {}
        labels = {
            "descricao": "Descricao",
            "chave_nfe_origem": "Chave NF-e origem",
            "ncm": "NCM",
            "cest": "CEST",
            "origem_entrada": "Origem (entrada)",
            "cfop_saida_fornecedor": "CFOP saida fornecedor",
            "c_classtrib": "cClassTrib",
            "c_benef": "cBenef",
            "cst_icms": "CST ICMS",
            "reducao_bc_icms": "% Red BC ICMS",
            "aliquota_icms": "% ICMS",
            "cst_ipi": "CST IPI",
            "aliquota_ipi": "% IPI",
            "cst_pis": "CST PIS",
            "cst_cofins": "CST COFINS",
            "cst_pis_cofins": "CST PIS/COFINS",
            "aliquota_pis_cofins": "% PIS/COFINS",
            "aliquota_pis": "% PIS",
            "aliquota_cofins": "% COFINS",
            "bc_st": "BC ST",
            "mva": "MVA",
            "valor_icms_st": "Valor ICMS ST",
            "aliquota_icms_st": "% ICMS ST",
        }
        toggle_actions = QHBoxLayout()
        mark_all_btn = self.create_button("Marcar Todos", lambda: [check.setChecked(True) for check in checks.values()])
        unmark_all_btn = self.create_button("Desmarcar Todos", lambda: [check.setChecked(False) for check in checks.values()])
        toggle_actions.addWidget(mark_all_btn)
        toggle_actions.addWidget(unmark_all_btn)
        toggle_actions.addStretch()
        layout.addLayout(toggle_actions)

        checks_panel = self.create_panel()
        checks_panel_layout = QGridLayout(checks_panel)
        checks_panel_layout.setContentsMargins(12, 12, 12, 12)
        checks_panel_layout.setHorizontalSpacing(22)
        checks_panel_layout.setVerticalSpacing(8)
        for field_name in UPDATABLE_PRODUCT_FIELDS:
            check = QCheckBox(labels.get(field_name, field_name))
            check.setChecked(True)
            checks[field_name] = check
        half = (len(UPDATABLE_PRODUCT_FIELDS) + 1) // 2
        for index, field_name in enumerate(UPDATABLE_PRODUCT_FIELDS):
            row = index if index < half else index - half
            column = 0 if index < half else 1
            checks_panel_layout.addWidget(checks[field_name], row, column)
        layout.addWidget(checks_panel)
        actions = QHBoxLayout()
        ok_btn = self.create_button("Confirmar", dialog.accept, primary=True)
        cancel_btn = self.create_button("Cancelar", dialog.reject)
        actions.addStretch()
        actions.addWidget(ok_btn)
        actions.addWidget(cancel_btn)
        layout.addLayout(actions)
        if dialog.exec() != QDialog.Accepted:
            return set()
        return {field for field, check in checks.items() if check.isChecked()}

    def build_catalog_import_preview_ui(self) -> None:
        if getattr(self, "catalog_import_thread", None) is not None:
            QMessageBox.information(self, "Importacao XML Cadastros", "Ja existe uma operacao de importacao em andamento.")
            return
        source_value = self.catalog_import_sources_input.text().strip()
        if not source_value:
            QMessageBox.warning(self, "Importacao XML Cadastros", "Selecione XML(s) ou uma pasta com XMLs.")
            return
        self.catalog_import_preview_rows = []
        self.catalog_import_ignored_rows = []
        self.catalog_import_ignored_button.setEnabled(False)
        self.set_table_rows(self.catalog_import_preview_table, [])
        self.catalog_import_progress.setValue(0)
        self.catalog_import_status.setText("Iniciando leitura dos XMLs...")
        self.statusBar().showMessage("Gerando previa da importacao XML...")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        self.catalog_import_thread = QThread(self)
        self.catalog_import_worker = CatalogImportPreviewWorker(self.mysql_repo, self.environment, source_value)
        self.catalog_import_worker.moveToThread(self.catalog_import_thread)
        self.catalog_import_thread.started.connect(self.catalog_import_worker.run)
        self.catalog_import_worker.progress.connect(self.update_catalog_import_progress)
        self.catalog_import_worker.finished.connect(self.render_catalog_import_preview_results)
        self.catalog_import_worker.failed.connect(self.handle_catalog_import_preview_error)
        self.catalog_import_worker.finished.connect(self.catalog_import_thread.quit)
        self.catalog_import_worker.failed.connect(self.catalog_import_thread.quit)
        self.catalog_import_thread.finished.connect(self.catalog_import_worker.deleteLater)
        self.catalog_import_thread.finished.connect(self.catalog_import_thread.deleteLater)
        self.catalog_import_thread.finished.connect(lambda: setattr(self, "catalog_import_worker", None))
        self.catalog_import_thread.finished.connect(lambda: setattr(self, "catalog_import_thread", None))
        self.catalog_import_thread.start()

    def render_catalog_import_preview_results(self, preview_rows: list[CatalogImportPreviewRow], stats: dict[str, int]) -> None:
        try:
            self.catalog_import_preview_rows = preview_rows
            self.catalog_import_ignored_rows = list(stats.get("ignored_rows", [])) if isinstance(stats, dict) else []
            self.catalog_import_ignored_button.setEnabled(bool(self.catalog_import_ignored_rows))
            table_rows = [
                [
                    row.xml_file,
                    row.empresa,
                    "SIM" if row.company_exists else "NAO",
                    row.fornecedor,
                    row.fornecedor_uf,
                    "SIM" if row.supplier_exists else "NAO",
                    row.codigo_fornecedor,
                    row.descricao,
                    row.ean,
                    row.ncm,
                    row.cst_icms,
                    row.reducao_bc_icms,
                    row.aliquota_icms,
                    row.bc_st,
                    row.valor_icms_st,
                    row.aliquota_icms_st,
                    row.cst_pis,
                    row.cst_cofins,
                    row.aliquota_pis,
                    row.aliquota_cofins,
                    row.classificacao,
                    "SIM" if row.exists else "NAO",
                ]
                for row in preview_rows
            ]
            self.set_table_rows(self.catalog_import_preview_table, table_rows)
            self.catalog_import_progress.setValue(100)
            self.catalog_import_status.setText(
                f"Previa pronta: XMLs {stats['xml_total']} | validos {stats['xml_valid']} | ignorados {stats['xml_ignored']} | "
                f"produtos na previa {stats['products_previewed']} | sem chave {stats['products_skipped_without_key']}."
            )
            self.statusBar().showMessage("Previa da importacao concluida.")
        finally:
            QApplication.restoreOverrideCursor()

    def handle_catalog_import_preview_error(self, error_text: str) -> None:
        try:
            if "duplicate entry" in error_text.lower() and "cad_produtos_fornecedor" in error_text.lower():
                friendly = (
                    "Foram encontrados registros duplicados antigos no cadastro de produtos por fornecedor.\n"
                    "Use a tela 'Limpeza Duplicados' para corrigir e gere a previa novamente."
                )
                self.catalog_import_status.setText(f"Previa gerada com alerta: {friendly}")
                self.statusBar().showMessage("Alerta de consistencia encontrado. Use Limpeza Duplicados.")
                QMessageBox.warning(self, "Importacao XML Cadastros", friendly)
            else:
                self.catalog_import_status.setText(f"Falha ao gerar previa: {error_text}")
                self.statusBar().showMessage("Falha ao gerar previa da importacao.")
                QMessageBox.critical(self, "Importacao XML Cadastros", error_text)
        finally:
            QApplication.restoreOverrideCursor()

    def open_catalog_import_ignored_popup(self) -> None:
        rows = getattr(self, "catalog_import_ignored_rows", [])
        if not rows:
            QMessageBox.information(self, "Importacao XML Cadastros", "Nao ha XMLs ignorados na ultima previa.")
            return
        self.open_table_popup(
            "XMLs Ignorados na Previa",
            ["Arquivo", "Caminho", "Motivo"],
            rows,
            1500,
            700,
        )

    def open_catalog_import_xml_file(self, row_index: int, column_index: int) -> None:
        if column_index != 0:
            return
        if row_index < 0 or row_index >= len(self.catalog_import_preview_rows):
            return
        selected_row = self.catalog_import_preview_rows[row_index]
        xml_path = Path(selected_row.xml_file_path)
        if not xml_path.exists():
            QMessageBox.warning(self, "Importacao XML Cadastros", f"Arquivo XML nao encontrado:\n{xml_path}")
            return
        try:
            subprocess.Popen(["explorer", "/select,", str(xml_path)])
        except Exception:
            os.startfile(str(xml_path.parent))

    def update_catalog_import_progress(self, current: int, total: int, message: str) -> None:
        if total <= 0:
            self.catalog_import_progress.setValue(0)
        else:
            self.catalog_import_progress.setValue(int((current / total) * 100))
        self.catalog_import_status.setText(message)

    def select_single_file(self, field: QLineEdit, title: str, file_filter: str) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        if file_path:
            field.setText(file_path)

    def select_directory(self, field: QLineEdit, title: str) -> None:
        directory = QFileDialog.getExistingDirectory(self, title)
        if directory:
            field.setText(directory)

    def write_audit_log(self, event_type: str, message: str) -> None:
        try:
            timestamp = dt.datetime.now().isoformat(timespec="seconds")
            event = str(event_type or "").strip() or "EVENTO"
            text = str(message or "").replace("\n", " ")
            self.audit_log_path.parent.mkdir(parents=True, exist_ok=True)
            with self.audit_log_path.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{timestamp}\t{self.audit_session_id}\t{event}\t{text}\n")
        except Exception:
            pass

    def get_current_runtime_rules(self) -> list[dict[str, object]]:
        if not hasattr(self, "runtime_rules_text"):
            return []
        self.current_runtime_rules = parse_runtime_rule_lines(self.runtime_rules_text.toPlainText())
        return self.current_runtime_rules

    def runtime_rule_lines(self) -> list[str]:
        if not hasattr(self, "runtime_rules_text"):
            return []
        return [line.strip() for line in self.runtime_rules_text.toPlainText().splitlines() if line.strip()]

    def validate_runtime_rules(self) -> None:
        try:
            rules = self.get_current_runtime_rules()
            lines = self.runtime_rule_lines()
            rows = [[index + 1, runtime_rule_summary(line), line] for index, line in enumerate(lines)]
            self.set_table_rows(self.runtime_rules_table, rows)
            self.runtime_rules_status.setText(f"Regras validas: {len(rules)}.")
            self.remember_runtime_rules(lines, "validacao_qt")
            self.statusBar().showMessage(f"{len(rules)} regra(s) dinamica(s) validada(s).")
            self.update_reprocess_buttons_state()
        except Exception as exc:
            self.update_reprocess_buttons_state()
            self.runtime_rules_status.setText(f"Erro: {exc}")
            QMessageBox.critical(self, "Regras Dinamicas", str(exc))

    def clear_runtime_rules(self) -> None:
        self.runtime_rules_text.clear()
        self.runtime_rules_table.setRowCount(0)
        self.current_runtime_rules = []
        self.runtime_rules_status.setText("Informe as regras e clique em Validar regras.")
        self.update_reprocess_buttons_state()

    def update_reprocess_buttons_state(self) -> None:
        if getattr(self, "reprocess_thread", None) is not None:
            if hasattr(self, "entry_reprocess_button"):
                self.entry_reprocess_button.setEnabled(False)
            if hasattr(self, "sale_reprocess_button"):
                self.sale_reprocess_button.setEnabled(False)
            return
        entry_enabled = False
        sale_enabled = False
        if hasattr(self, "runtime_rules_text"):
            try:
                rules = parse_runtime_rule_lines(self.runtime_rules_text.toPlainText())
                entry_enabled = any(str(rule.get("operation_type", "")).strip().lower() == "entrada" for rule in rules)
                sale_enabled = any(str(rule.get("operation_type", "")).strip().lower() == "saida" for rule in rules)
            except Exception:
                entry_enabled = False
                sale_enabled = False
        if hasattr(self, "entry_reprocess_button"):
            self.entry_reprocess_button.setEnabled(entry_enabled)
        if hasattr(self, "sale_reprocess_button"):
            self.sale_reprocess_button.setEnabled(sale_enabled)

    def enriched_runtime_rule_row(self, row: dict[str, object], operation_type: str = "") -> dict[str, object]:
        enriched = dict(row)
        if operation_type and not str(enriched.get("operation_type", "")).strip():
            enriched["operation_type"] = operation_type
        if "cst_icms" not in enriched and "cst" in enriched:
            enriched["cst_icms"] = enriched.get("cst")
        if "sale_value" not in enriched:
            enriched["sale_value"] = self.first_row_value(enriched, "total_operation_value", "operation_value", default=Decimal("0"))
        if "document_tax_id" not in enriched:
            document_key = str(enriched.get("document_key", "")).strip()
            tax_id = extract_tax_id_from_document_key(document_key)
            if not tax_id:
                tax_id = self.digits_only(str(enriched.get("participant_tax_id", "")))
            if len(tax_id) in {11, 14}:
                enriched["document_tax_id"] = tax_id
        return enriched

    def find_matching_runtime_rule_line(self, row: dict[str, object], operation_type: str = "") -> str:
        candidate = self.enriched_runtime_rule_row(row, operation_type)
        for rule_line in self.runtime_rule_lines():
            try:
                parsed_rules = parse_runtime_rule_lines(rule_line)
            except Exception:
                continue
            if get_first_matching_icms_rule(parsed_rules, candidate) is not None:
                return rule_line
        return ""

    def remove_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        answer = QMessageBox.question(self, "Remover Regra Dinamica", f"Deseja remover esta regra dinamica?\n\n{normalized_line}")
        if answer != QMessageBox.Yes:
            return
        updated_lines = [line for line in self.runtime_rule_lines() if line != normalized_line]
        self.runtime_rules_text.setPlainText("\n".join(updated_lines))
        self.validate_runtime_rules()
        self.statusBar().showMessage("Regra dinamica removida.")

    def edit_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        try:
            prefill = self.build_runtime_rule_prefill_from_line(normalized_line)
        except Exception as exc:
            QMessageBox.critical(self, "Editar regra", f"Nao foi possivel carregar a regra:\n{exc}")
            return
        self.open_runtime_rule_builder(prefill=prefill, replace_rule_line=normalized_line)

    def build_runtime_rule_prefill_from_row(self, row: dict[str, object], operation_type: str, include_document: bool = False) -> dict[str, object]:
        enriched = self.enriched_runtime_rule_row(row, operation_type)
        rate_value = enriched.get("icms_rate", enriched.get("nominal_icms_rate", Decimal("0")))
        try:
            rate_text = format(self.decimal_value(rate_value).normalize(), "f")
        except Exception:
            rate_text = str(rate_value or "").strip()
        prefill = {
            "operation": str(enriched.get("operation_type", operation_type)).strip().lower() or operation_type.lower(),
            "cst": str(self.first_row_value(enriched, "cst_icms", "cst")).strip().split("|", 1)[0].strip(),
            "cfop": str(enriched.get("cfop", "")).strip().split("|", 1)[0].strip(),
            "rate": rate_text,
            "codes": str(enriched.get("code", "")).strip(),
        }
        if include_document:
            prefill["document"] = str(enriched.get("document_number", "")).strip()
            prefill["tax_id"] = str(enriched.get("document_tax_id", "")).strip()
        return prefill

    def add_runtime_rule_actions_to_menu(self, menu: QMenu, row: dict[str, object], operation_type: str, include_document: bool = False) -> None:
        matching_rule_line = self.find_matching_runtime_rule_line(row, operation_type)
        if matching_rule_line:
            menu.addAction("Editar Regra Dinamica", lambda current_rule=matching_rule_line: self.edit_runtime_rule_line(current_rule))
            menu.addAction("Excluir Regra Dinamica", lambda current_rule=matching_rule_line: self.remove_runtime_rule_line(current_rule))
        else:
            menu.addAction(
                "Criar Regra Dinamica",
                lambda current_row=dict(row): self.open_runtime_rule_builder(
                    prefill=self.build_runtime_rule_prefill_from_row(current_row, operation_type, include_document)
                ),
            )

    def show_runtime_rule_context_menu_for_table(
        self,
        table: QTableWidget,
        source_rows: list[dict[str, object]],
        operation_type: str,
        pos: object,
    ) -> None:
        item = table.itemAt(pos)
        if item is None:
            return
        data_index = item.data(Qt.UserRole)
        try:
            row_index = int(data_index)
        except (TypeError, ValueError):
            row_index = item.row()
        if row_index < 0 or row_index >= len(source_rows):
            return
        table.selectRow(item.row())
        menu = QMenu(table)
        self.add_runtime_rule_actions_to_menu(menu, source_rows[row_index], operation_type, include_document=False)
        menu.addSeparator()
        menu.addAction("Abrir Detalhamento", lambda current_row=source_rows[row_index]: self.open_row_detail_popup(current_row, operation_type))
        menu.exec(table.viewport().mapToGlobal(pos))

    def show_runtime_rule_context_menu_for_payloads(
        self,
        table: QTableWidget,
        payload_rows: list[dict[str, object]],
        operation_type: str,
        pos: object,
        include_document: bool = True,
    ) -> None:
        item = table.itemAt(pos)
        if item is None:
            return
        data_index = item.data(Qt.UserRole)
        try:
            row_index = int(data_index)
        except (TypeError, ValueError):
            row_index = item.row()
        if row_index < 0 or row_index >= len(payload_rows):
            return
        table.selectRow(item.row())
        menu = QMenu(table)
        self.add_runtime_rule_actions_to_menu(menu, payload_rows[row_index], operation_type, include_document=include_document)
        menu.exec(table.viewport().mapToGlobal(pos))

    def save_runtime_rule_history(self) -> None:
        save_runtime_rule_history_file(self.runtime_rule_history_path, self.runtime_rule_history)

    def remember_runtime_rules(self, rule_lines: list[str], source: str) -> None:
        if not rule_lines:
            return
        self.runtime_rule_history = remember_runtime_rules_in_history(self.runtime_rule_history, rule_lines, source)
        self.save_runtime_rule_history()
        self.refresh_runtime_rule_history_list()

    def clear_runtime_rule_filter(self) -> None:
        if hasattr(self, "runtime_rule_filter_input"):
            self.runtime_rule_filter_input.clear()
        self.refresh_runtime_rule_history_list()

    def refresh_runtime_rule_history_list(self) -> None:
        if not hasattr(self, "runtime_rule_history_table"):
            return
        filter_text = normalize_text(self.runtime_rule_filter_input.text() if hasattr(self, "runtime_rule_filter_input") else "")

        def sort_key(item: dict[str, object]) -> tuple[str, str]:
            return (str(item.get("last_used_at", "")), str(item.get("rule_line", "")))

        rows: list[list[object]] = []
        for item in sorted(self.runtime_rule_history, key=sort_key, reverse=True):
            rule_line = str(item.get("rule_line", "")).strip()
            searchable_text = normalize_text(f"{rule_line} {item.get('summary', '')}")
            if filter_text and filter_text not in searchable_text:
                continue
            rows.append([
                rule_line,
                str(item.get("summary", "")).strip(),
                int(item.get("use_count", 0) or 0),
                str(item.get("last_used_at", "")).strip(),
            ])
        self.set_table_rows(self.runtime_rule_history_table, rows)
        if rows:
            self.runtime_rule_history_table.selectRow(0)
        self.update_runtime_rule_history_preview()

    def selected_runtime_rule_history_line(self) -> str:
        if not hasattr(self, "runtime_rule_history_table"):
            return ""
        selected = self.runtime_rule_history_table.selectionModel().selectedRows()
        if not selected:
            return ""
        item = self.runtime_rule_history_table.item(selected[0].row(), 0)
        return item.text().strip() if item is not None else ""

    def update_runtime_rule_history_preview(self) -> None:
        if not hasattr(self, "runtime_rule_history_preview"):
            return
        self.runtime_rule_history_preview.setPlainText(self.selected_runtime_rule_history_line())

    def insert_selected_runtime_rule(self) -> None:
        rule_line = self.selected_runtime_rule_history_line()
        if not rule_line:
            QMessageBox.warning(self, "Historico", "Selecione uma regra do historico.")
            return
        self.append_runtime_rule_line(rule_line)

    def edit_selected_runtime_rule(self) -> None:
        rule_line = self.selected_runtime_rule_history_line()
        if not rule_line:
            QMessageBox.warning(self, "Historico", "Selecione uma regra do historico.")
            return
        try:
            prefill = self.build_runtime_rule_prefill_from_line(rule_line)
        except Exception as exc:
            QMessageBox.critical(self, "Editar regra", f"Nao foi possivel carregar a regra:\n{exc}")
            return
        self.open_runtime_rule_builder(prefill=prefill, replace_rule_line=rule_line)

    def delete_selected_runtime_rule(self) -> None:
        rule_line = self.selected_runtime_rule_history_line()
        if not rule_line:
            QMessageBox.warning(self, "Historico", "Selecione uma regra do historico.")
            return
        answer = QMessageBox.question(self, "Historico", "Deseja excluir a regra selecionada do historico?")
        if answer != QMessageBox.Yes:
            return
        self.runtime_rule_history = remove_runtime_rule(self.runtime_rule_history, rule_line)
        self.save_runtime_rule_history()
        self.refresh_runtime_rule_history_list()
        self.statusBar().showMessage("Regra removida do historico.")

    def append_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        current_text = self.runtime_rules_text.toPlainText().strip()
        next_text = f"{current_text}\n{normalized_line}" if current_text else normalized_line
        self.runtime_rules_text.setPlainText(next_text)
        self.validate_runtime_rules()

    def replace_runtime_rule_line(self, old_rule_line: str, new_rule_line: str) -> None:
        old_normalized = str(old_rule_line or "").strip()
        new_normalized = str(new_rule_line or "").strip()
        if not new_normalized:
            return
        lines = self.runtime_rules_text.toPlainText().splitlines()
        replaced = False
        updated_lines: list[str] = []
        for line in lines:
            if not replaced and line.strip() == old_normalized:
                updated_lines.append(new_normalized)
                replaced = True
            else:
                updated_lines.append(line)
        if not replaced:
            updated_lines.append(new_normalized)
        self.runtime_rules_text.setPlainText("\n".join(line for line in updated_lines if line.strip()))
        self.validate_runtime_rules()

    def build_runtime_rule_prefill_from_line(self, rule_line: str) -> dict[str, object]:
        parsed_rules = parse_runtime_rule_lines(rule_line)
        if not parsed_rules:
            return {}
        rule = parsed_rules[0]

        def decimal_text(key: str) -> str:
            value = rule.get(key, "")
            if isinstance(value, Decimal):
                return format(value.normalize(), "f")
            return str(value or "").strip()

        match_codes = rule.get("match_codes", "")
        if isinstance(match_codes, set):
            codes_text = ",".join(sorted(str(value).strip() for value in match_codes if str(value).strip()))
        else:
            codes_text = str(match_codes or "").strip()

        return {
            "operation": str(rule.get("operation_type", "entrada")).strip() or "entrada",
            "document": str(rule.get("document_number", "")).strip(),
            "tax_id": str(rule.get("document_tax_id", "")).strip(),
            "cst": str(rule.get("cst_icms", "")).strip(),
            "cfop": str(rule.get("cfop", "")).strip(),
            "rate": decimal_text("match_rate"),
            "codes": codes_text,
            "new_cst": str(rule.get("new_cst", "")).strip(),
            "new_cfop": str(rule.get("new_cfop", "")).strip(),
            "force_rate": decimal_text("force_rate"),
            "set_base": decimal_text("set_base_icms"),
            "set_icms_value": decimal_text("set_icms_value"),
            "base_reduction_percent": decimal_text("base_reduction_percent"),
            "zero_icms": bool(rule.get("zero_icms", False)),
            "use_sale_value_base": bool(rule.get("use_sale_value_as_base", False)),
            "recalculate_icms_value": bool(rule.get("recalculate_icms_value", False)),
            "recalculate_reduced_base": bool(rule.get("recalculate_reduced_base", False)),
        }

    def open_runtime_rule_builder(
        self,
        prefill: dict[str, object] | None = None,
        replace_rule_line: str | None = None,
    ) -> None:
        prefill = prefill or {}
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle("Editar Regra Dinamica" if replace_rule_line else "Adicionar Regra Dinamica")
        self.resize_dialog_to_screen(dialog, 980, 560)

        container = QVBoxLayout(dialog)
        container.setContentsMargins(16, 16, 16, 16)
        title = QLabel("Monte a regra separando o que sera filtrado do que sera alterado no reprocessamento.")
        title.setObjectName("sectionTitle")
        container.addWidget(title)

        help_label = QLabel("Campos vazios sao ignorados. Use os filtros para localizar os itens, e os campos de reprocessamento para informar a alteracao.")
        help_label.setObjectName("muted")
        container.addWidget(help_label)

        filter_box = QGroupBox("Campos para filtro - localizar quais itens serao alterados")
        filter_grid = QGridLayout(filter_box)
        filter_grid.setSpacing(10)
        container.addWidget(filter_box)

        reprocess_box = QGroupBox("Campos para reprocessamento - alteracoes que serao aplicadas")
        reprocess_grid = QGridLayout(reprocess_box)
        reprocess_grid.setSpacing(10)
        container.addWidget(reprocess_box)

        fields: dict[str, QLineEdit] = {}

        def add_entry(parent_grid: QGridLayout, row: int, column: int, key: str, label: str) -> QLineEdit:
            field_box = QWidget()
            field_layout = QVBoxLayout(field_box)
            field_layout.setContentsMargins(0, 0, 0, 0)
            field_layout.setSpacing(4)
            field_layout.addWidget(QLabel(label))
            line_edit = QLineEdit(str(prefill.get(key, "") or ""))
            field_layout.addWidget(line_edit)
            parent_grid.addWidget(field_box, row, column)
            fields[key] = line_edit
            return line_edit

        operation_box = QWidget()
        operation_layout = QVBoxLayout(operation_box)
        operation_layout.setContentsMargins(0, 0, 0, 0)
        operation_layout.setSpacing(4)
        operation_layout.addWidget(QLabel("Tipo de operacao"))
        operation_combo = QComboBox()
        operation_combo.addItems(["entrada", "saida"])
        operation_value = str(prefill.get("operation", "entrada") or "entrada").lower()
        operation_combo.setCurrentText(operation_value if operation_value in {"entrada", "saida"} else "entrada")
        operation_layout.addWidget(operation_combo)
        filter_grid.addWidget(operation_box, 0, 0)

        add_entry(filter_grid, 0, 1, "cst", "CST atual")
        add_entry(filter_grid, 0, 2, "cfop", "CFOP atual")
        add_entry(filter_grid, 1, 0, "rate", "Aliquota atual")
        add_entry(filter_grid, 1, 1, "codes", "Codigo(s) do produto")
        add_entry(filter_grid, 1, 2, "document", "Documento")
        add_entry(filter_grid, 2, 0, "tax_id", "CNPJ/CPF do documento")

        add_entry(reprocess_grid, 0, 0, "new_cst", "Novo CST")
        add_entry(reprocess_grid, 0, 1, "new_cfop", "Novo CFOP")
        add_entry(reprocess_grid, 0, 2, "force_rate", "Nova aliquota")
        add_entry(reprocess_grid, 1, 0, "set_base", "Definir base ICMS")
        add_entry(reprocess_grid, 1, 1, "set_icms_value", "Definir valor ICMS")
        add_entry(reprocess_grid, 1, 2, "base_reduction_percent", "% reducao de base")

        options_box = QGroupBox("Opcoes de recalculo")
        options = QGridLayout(options_box)
        options.setSpacing(8)
        container.addWidget(options_box)
        zero_icms_check = QCheckBox("Zerar ICMS/Base/Valor")
        use_sale_base_check = QCheckBox("Usar valor da operacao como base ICMS")
        recalc_value_check = QCheckBox("Recalcular valor do ICMS")
        recalc_reduced_base_check = QCheckBox("Recalcular base de calculo da reducao")
        zero_icms_check.setChecked(bool(prefill.get("zero_icms", False)))
        use_sale_base_check.setChecked(bool(prefill.get("use_sale_value_base", False)))
        recalc_value_check.setChecked(bool(prefill.get("recalculate_icms_value", False)))
        recalc_reduced_base_check.setChecked(bool(prefill.get("recalculate_reduced_base", False)))
        options.addWidget(zero_icms_check, 0, 0)
        options.addWidget(use_sale_base_check, 0, 1)
        options.addWidget(recalc_value_check, 1, 0)
        options.addWidget(recalc_reduced_base_check, 1, 1)

        example = QLabel("Exemplo: tipo=entrada; cnpj=12345678000199; cst=000; cfop=1102; novo_cfop=1101")
        example.setObjectName("muted")
        container.addWidget(example)

        buttons = QHBoxLayout()
        save_button = self.create_button("Salvar" if replace_rule_line else "Adicionar", lambda: save_runtime_rule(), primary=True)
        cancel_button = self.create_button("Cancelar", dialog.reject)
        buttons.addWidget(save_button)
        buttons.addWidget(cancel_button)
        buttons.addStretch()
        container.addLayout(buttons)

        def sync_reduced_base_option() -> None:
            if recalc_reduced_base_check.isChecked():
                recalc_value_check.setChecked(True)

        def sync_reduction_percent(text: str) -> None:
            if text.strip():
                recalc_reduced_base_check.setChecked(True)
                recalc_value_check.setChecked(True)

        recalc_reduced_base_check.toggled.connect(lambda _checked: sync_reduced_base_option())
        fields["base_reduction_percent"].textChanged.connect(sync_reduction_percent)

        def save_runtime_rule() -> None:
            fragments = [f"tipo={operation_combo.currentText().strip()}"]
            optional_fields = [
                ("documento", fields["document"].text()),
                ("cnpj", fields["tax_id"].text()),
                ("cst", fields["cst"].text()),
                ("cfop", fields["cfop"].text()),
                ("aliquota", fields["rate"].text()),
                ("codigos", fields["codes"].text()),
                ("novo_cst", fields["new_cst"].text()),
                ("novo_cfop", fields["new_cfop"].text()),
                ("nova_aliquota", fields["force_rate"].text()),
                ("definir_base_icms", fields["set_base"].text()),
                ("definir_valor_icms", fields["set_icms_value"].text()),
                ("percentual_reducao_base", fields["base_reduction_percent"].text()),
            ]
            for key, value in optional_fields:
                text = str(value or "").strip()
                if text:
                    fragments.append(f"{key}={text}")
            if zero_icms_check.isChecked():
                fragments.append("zerar_icms=sim")
            if use_sale_base_check.isChecked():
                fragments.append("usar_valor_operacao_como_base=sim")
            if recalc_value_check.isChecked():
                fragments.append("recalcular_valor_icms=sim")
            if recalc_reduced_base_check.isChecked():
                fragments.append("recalcular_base_reducao=sim")

            if len(fragments) <= 1:
                QMessageBox.warning(dialog, "Regra incompleta", "Informe pelo menos um filtro ou uma acao para montar a regra.")
                return
            if recalc_reduced_base_check.isChecked() and not fields["base_reduction_percent"].text().strip():
                QMessageBox.warning(dialog, "Regra incompleta", "Informe o percentual de reducao de base.")
                return

            action_keys = {
                "novo_cst", "novo_cfop", "nova_aliquota", "definir_base_icms", "definir_valor_icms",
                "percentual_reducao_base", "zerar_icms", "usar_valor_operacao_como_base",
                "recalcular_valor_icms", "recalcular_base_reducao",
            }
            if not any(fragment.split("=", 1)[0] in action_keys for fragment in fragments[1:]):
                QMessageBox.warning(dialog, "Regra sem acao", "Informe ao menos uma acao de recalculo para a regra.")
                return

            rule_line = "; ".join(fragments)
            try:
                parse_runtime_rule_lines(rule_line)
            except Exception as exc:
                QMessageBox.critical(dialog, "Regra invalida", str(exc))
                return

            if replace_rule_line:
                self.replace_runtime_rule_line(replace_rule_line, rule_line)
            else:
                self.append_runtime_rule_line(rule_line)
            dialog.accept()

        dialog.exec()

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
            self.save_mysql_settings()
            self.mysql_repo.ensure_schema()
            self.statusBar().showMessage("Banco e tabelas atualizados.")
            QMessageBox.information(self, "Configuracoes", "Banco e tabelas atualizados.")
        except Exception as exc:
            QMessageBox.critical(self, "Configuracoes", str(exc))

    def test_mysql_connection(self) -> None:
        try:
            self.save_mysql_settings()
            self.mysql_repo.test_connection()
            self.statusBar().showMessage("Conexao MySQL OK.")
            QMessageBox.information(self, "Configuracoes", "Conexao MySQL OK.")
        except Exception as exc:
            QMessageBox.critical(self, "Configuracoes", str(exc))

    def run_database_backup(self) -> None:
        from app.services.database_backup import suggest_backup_filename
        self.save_mysql_settings()
        config = self.mysql_repo.load_config()
        database_name = config.get("database", "banco")
        initial_filename = suggest_backup_filename(str(database_name))
        dest, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar backup do banco de dados",
            initial_filename,
            "SQL Dump (*.sql);;Todos os arquivos (*)",
        )
        if not dest:
            return
        dest_path = Path(dest)
        connection_config = {
            "host": config.get("host", "127.0.0.1"),
            "port": int(str(config.get("port", "3306")) or "3306"),
            "user": config.get("user", "root"),
            "password": config.get("password", ""),
            "database": database_name,
            "connection_timeout": 10,
        }
        self.backup_btn.setEnabled(False)
        self.backup_progress_bar.setVisible(True)
        self.backup_progress_bar.setValue(0)
        self.backup_status_label.setText("Iniciando backup...")
        self.statusBar().showMessage("Backup do banco em andamento...")

        self._backup_thread = QThread(self)
        self._backup_worker = BackupDatabaseWorker(connection_config, dest_path)
        self._backup_worker.moveToThread(self._backup_thread)
        self._backup_thread.started.connect(self._backup_worker.run)
        self._backup_worker.progress.connect(self._on_backup_progress)
        self._backup_worker.finished.connect(self._on_backup_success)
        self._backup_worker.failed.connect(self._on_backup_error)
        self._backup_worker.finished.connect(self._backup_thread.quit)
        self._backup_worker.failed.connect(self._backup_thread.quit)
        self._backup_thread.start()

    def _on_backup_progress(self, current: int, total: int, message: str) -> None:
        if total > 0:
            self.backup_progress_bar.setValue(int(current * 100 / total))
        self.backup_status_label.setText(message)

    def _on_backup_success(self, dest_path: Path) -> None:
        self.backup_btn.setEnabled(True)
        self.backup_progress_bar.setVisible(False)
        size_kb = round(dest_path.stat().st_size / 1024, 1)
        msg = f"Backup concluido: {dest_path.name} ({size_kb} KB)"
        self.backup_status_label.setText(msg)
        self.statusBar().showMessage(msg)
        reply = QMessageBox.question(self, "Backup", f"{msg}\n\nDeseja abrir a pasta do arquivo?")
        if reply == QMessageBox.Yes:
            os.startfile(str(dest_path.parent))

    def _on_backup_error(self, error: str) -> None:
        self.backup_btn.setEnabled(True)
        self.backup_progress_bar.setVisible(False)
        self.backup_status_label.setText(f"Erro: {error}")
        self.statusBar().showMessage("Falha no backup do banco.")
        QMessageBox.critical(self, "Backup", f"Falha ao gerar backup:\n\n{error}")

    def process_entry_query(self) -> None:
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            sped_paths = parse_selected_paths(self.sped_input.text())
            xml_sources = parse_selected_paths(self.xml_input.text())
            period_labels, rows = build_entry_period_comparison_rows(sped_paths, xml_sources)
            sale_period_labels, sale_rows = build_sale_period_comparison_rows(sped_paths, xml_sources)
            self.entry_rows = list(rows)
            # Mantem os dados de saida sincronizados com o mesmo processamento de entradas,
            # para permitir apuracao completa (credito + debito) na mesma tela.
            self.sale_rows = list(sale_rows)
            self.rebuild_period_checks(list(period_labels))
            self.rebuild_sale_period_checks(list(sale_period_labels))
            if hasattr(self, "sale_sped_input"):
                self.sale_sped_input.setText(self.sped_input.text())
            if hasattr(self, "sale_xml_input"):
                self.sale_xml_input.setText(self.xml_input.text())
            self.refresh_entry_table()
            self.statusBar().showMessage(
                f"Entradas processadas: {len(rows)} linha(s), {len(period_labels)} periodo(s). "
                f"Saidas sincronizadas: {len(sale_rows)} linha(s). {dt.datetime.now().strftime('%H:%M:%S')}"
            )
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
            entry_period_labels, entry_rows = build_entry_period_comparison_rows(sped_paths, xml_sources)
            period_labels, rows = build_sale_period_comparison_rows(sped_paths, xml_sources)
            self.entry_rows = list(entry_rows)
            self.sale_rows = list(rows)
            self.rebuild_period_checks(list(entry_period_labels))
            self.rebuild_sale_period_checks(list(period_labels))
            if hasattr(self, "sped_input"):
                self.sped_input.setText(self.sale_sped_input.text())
            if hasattr(self, "xml_input"):
                self.xml_input.setText(self.sale_xml_input.text())
            self.refresh_sale_table()
            self.statusBar().showMessage(
                f"Saidas processadas: {len(rows)} linha(s), {len(period_labels)} periodo(s). "
                f"Entradas sincronizadas: {len(entry_rows)} linha(s). {dt.datetime.now().strftime('%H:%M:%S')}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "SPED Qt", str(exc))
            self.statusBar().showMessage(f"Falha no processamento: {exc}")
        finally:
            QApplication.restoreOverrideCursor()

    def process_entry_exit_analysis(self) -> None:
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            sped_paths = parse_selected_paths(self.entry_exit_sped_input.text())
            if not sped_paths:
                QMessageBox.warning(self, "Analise Entrada e Saida", "Selecione ao menos um arquivo SPED Fiscal.")
                return
            detail_rows, excel_rows, totals = build_entry_exit_analysis_rows(sped_paths)
            self.entry_exit_detail_rows = list(detail_rows)
            self.entry_exit_excel_rows = list(excel_rows)
            self.entry_exit_totals = dict(totals)
            self.refresh_entry_exit_analysis()
            self.entry_exit_status_label.setText(f"Analise montada com {len(detail_rows)} agrupamento(s) de CST/CFOP/aliquota.")
            self.statusBar().showMessage(f"Analise Entrada e Saida processada: {len(detail_rows)} agrupamento(s). {dt.datetime.now().strftime('%H:%M:%S')}")
        except Exception as exc:
            QMessageBox.critical(self, "Analise Entrada e Saida", str(exc))
            self.statusBar().showMessage(f"Falha na analise Entrada e Saida: {exc}")
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

    def get_consultation_sped_paths_for_reprocess(self, operation_type: str) -> list[Path]:
        source_text = ""
        if operation_type == "Entrada" and hasattr(self, "sped_input"):
            source_text = self.sped_input.text()
        elif operation_type == "Saida" and hasattr(self, "sale_sped_input"):
            source_text = self.sale_sped_input.text()
        paths = parse_selected_paths(source_text)
        unique_paths: list[Path] = []
        seen: set[Path] = set()
        for path in paths:
            resolved = path.resolve() if path.exists() else path
            if resolved in seen:
                continue
            seen.add(resolved)
            unique_paths.append(path)
        return unique_paths

    def generate_adjusted_sped_from_consultation(
        self, operation_type: str, override_rules: list[dict] | None = None
    ) -> None:
        if getattr(self, "reprocess_thread", None) is not None:
            QMessageBox.information(self, "Reprocessar SPED", "Ja existe um reprocessamento em andamento.")
            return
        if override_rules is not None:
            # Regras do Confronto: combinam com as regras dinamicas existentes
            try:
                existing = self.get_current_runtime_rules()
            except Exception:
                existing = []
            runtime_rules = override_rules + existing
        else:
            try:
                runtime_rules = self.get_current_runtime_rules()
            except Exception as exc:
                QMessageBox.critical(self, "Reprocessar SPED", f"Corrija as regras dinamicas antes de reprocessar.\n\n{exc}")
                self.update_reprocess_buttons_state()
                return
        if not runtime_rules:
            QMessageBox.warning(self, "Reprocessar SPED", "Cadastre e valide ao menos uma regra dinamica.")
            self.update_reprocess_buttons_state()
            return

        sped_paths = self.get_consultation_sped_paths_for_reprocess(operation_type)
        if not sped_paths:
            QMessageBox.warning(self, "Reprocessar SPED", "Selecione ao menos um arquivo SPED na consulta.")
            return
        missing_paths = [str(path) for path in sped_paths if not path.exists()]
        if missing_paths:
            QMessageBox.critical(self, "Reprocessar SPED", "Os seguintes SPEDs nao existem mais:\n" + "\n".join(missing_paths))
            return

        output_paths: list[Path] = []
        if len(sped_paths) == 1:
            default_path = sped_paths[0].with_name(f"{sped_paths[0].stem}_ajustado{sped_paths[0].suffix or '.txt'}")
            selected_output, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar SPED reprocessado",
                str(default_path),
                "Arquivo SPED (*.txt *.sped);;Todos os arquivos (*.*)",
            )
            if not selected_output:
                return
            output_paths = [Path(selected_output)]
        else:
            selected_folder = QFileDialog.getExistingDirectory(self, "Selecione a pasta para salvar os SPEDs reprocessados")
            if not selected_folder:
                return
            output_dir = Path(selected_folder)
            output_paths = [output_dir / f"{path.stem}_ajustado{path.suffix or '.txt'}" for path in sped_paths]

        rule_lines = self.runtime_rule_lines()
        self.entry_reprocess_button.setEnabled(False)
        self.sale_reprocess_button.setEnabled(False)
        self.statusBar().showMessage("Preparando reprocessamento do SPED...")

        self.reprocess_thread = QThread()
        self.reprocess_worker = SpedReprocessWorker(
            sped_paths,
            output_paths,
            operation_type,
            runtime_rules,
            rule_lines,
            self.audit_log_path,
            self.audit_session_id,
        )
        self.reprocess_worker.moveToThread(self.reprocess_thread)
        self.reprocess_thread.started.connect(self.reprocess_worker.run)
        self.reprocess_worker.progress.connect(self.update_reprocess_status)
        self.reprocess_worker.finished.connect(lambda generated_files: self.handle_reprocess_finished(generated_files, rule_lines))
        self.reprocess_worker.failed.connect(self.handle_reprocess_error)
        self.reprocess_worker.finished.connect(self.reprocess_thread.quit)
        self.reprocess_worker.failed.connect(self.reprocess_thread.quit)
        self.reprocess_thread.finished.connect(self.reprocess_worker.deleteLater)
        self.reprocess_thread.finished.connect(self.reprocess_thread.deleteLater)
        self.reprocess_thread.finished.connect(lambda: setattr(self, "reprocess_worker", None))
        self.reprocess_thread.finished.connect(lambda: setattr(self, "reprocess_thread", None))
        self.reprocess_thread.finished.connect(self.update_reprocess_buttons_state)
        self.reprocess_thread.start()

    def update_reprocess_status(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.statusBar().showMessage(f"{message} ({percent}%)")

    def handle_reprocess_finished(self, generated_files: list[str], rule_lines: list[str]) -> None:
        self.remember_runtime_rules(rule_lines, "sped_ajustado_qt")
        self.statusBar().showMessage("SPED reprocessado com sucesso.")
        QMessageBox.information(self, "Reprocessar SPED", "Arquivos gerados com sucesso:\n" + "\n".join(generated_files))

    def handle_reprocess_error(self, message: str) -> None:
        self.statusBar().showMessage("Falha ao reprocessar SPED.")
        QMessageBox.critical(self, "Reprocessar SPED", f"Falha ao reprocessar SPED.\n\n{message}")

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
        replacements = {"Ã¡": "a", "Ã ": "a", "Ã£": "a", "Ã¢": "a", "Ã©": "e", "Ãª": "e", "Ã­": "i", "Ã³": "o", "Ã´": "o", "Ãµ": "o", "Ãº": "u", "Ã§": "c"}
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
                item.setData(Qt.UserRole, row_index)
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
        base_ratio = (total_base / total_operation * Decimal("100")) if total_operation > 0 else Decimal("0")
        self.metric_labels["base_ratio"].setText(self.format_number(base_ratio))

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
                item.setData(Qt.UserRole, row_index)
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

    def entry_exit_detail_headers(self) -> list[str]:
        return [
            "Tipo",
            "Classificacao",
            "CST",
            "CFOP",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Valor ICMS",
            "Total Operacao",
            "Base ICMS",
            "Base ICMS ST",
            "Valor ICMS ST",
            "Base IPI",
            "Aliq IPI",
            "Valor IPI",
        ]

    def entry_exit_footer_headers(self) -> list[str]:
        return [
            "Resumo",
            "Valor ICMS",
            "Total Operacao",
            "Base ICMS",
            "Base ICMS ST",
            "Valor ICMS ST",
            "Base IPI",
            "Valor IPI",
        ]

    def refresh_entry_exit_analysis(self) -> None:
        sale_rows = [row for row in self.entry_exit_detail_rows if str(row.get("operation_type", "")) == "Saida"]
        entry_rows = [row for row in self.entry_exit_detail_rows if str(row.get("operation_type", "")) == "Entrada"]
        self.set_table_rows(self.entry_exit_sale_table, [self.entry_exit_detail_display_row(row) for row in sale_rows])
        self.set_table_rows(self.entry_exit_entry_table, [self.entry_exit_detail_display_row(row) for row in entry_rows])
        self.set_table_rows(
            self.entry_exit_sale_footer_table,
            [self.entry_exit_footer_display_row(row) for row in build_entry_exit_footer_rows(sale_rows, "Saida")],
        )
        self.set_table_rows(
            self.entry_exit_entry_footer_table,
            [self.entry_exit_footer_display_row(row) for row in build_entry_exit_footer_rows(entry_rows, "Entrada")],
        )
        totals = self.entry_exit_totals
        self.entry_exit_metric_labels["saida_icms"].setText(self.format_number(totals.get("saida_icms", Decimal("0"))))
        self.entry_exit_metric_labels["entrada_icms"].setText(self.format_number(totals.get("entrada_icms", Decimal("0"))))
        self.entry_exit_metric_labels["recolher"].setText(self.format_number(totals.get("recolher", Decimal("0"))))
        self.entry_exit_metric_labels["percent_sale"].setText(f"{self.format_number(totals.get('percent_sale', Decimal('0')))}%")
        self.entry_exit_metric_labels["source_files"].setText(str(int(totals.get("source_files", Decimal("0")))))

    def entry_exit_detail_display_row(self, row: dict[str, object]) -> list[object]:
        return [
            row.get("marker", ""),
            row.get("category", ""),
            row.get("cst_icms", ""),
            row.get("cfop", ""),
            self.format_number(row.get("icms_rate", Decimal("0"))),
            self.format_number(row.get("effective_rate", Decimal("0"))),
            self.format_number(row.get("icms_value", Decimal("0"))),
            self.format_number(row.get("total_operation_value", Decimal("0"))),
            self.format_number(row.get("base_icms", Decimal("0"))),
            self.format_number(row.get("base_icms_st", Decimal("0"))),
            self.format_number(row.get("icms_st_value", Decimal("0"))),
            self.format_number(row.get("base_ipi", Decimal("0"))),
            self.format_number(row.get("ipi_rate", Decimal("0"))),
            self.format_number(row.get("ipi_value", Decimal("0"))),
        ]

    def entry_exit_footer_display_row(self, row: dict[str, object]) -> list[object]:
        return [
            row.get("label", ""),
            self.format_number(row.get("icms_value", Decimal("0"))),
            self.format_number(row.get("total_operation_value", Decimal("0"))),
            self.format_number(row.get("base_icms", Decimal("0"))),
            self.format_number(row.get("base_icms_st", Decimal("0"))),
            self.format_number(row.get("icms_st_value", Decimal("0"))),
            self.format_number(row.get("base_ipi", Decimal("0"))),
            self.format_number(row.get("ipi_value", Decimal("0"))),
        ]

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

    def clear_entry_exit_page(self) -> None:
        self.entry_exit_sped_input.clear()
        self.entry_exit_detail_rows = []
        self.entry_exit_excel_rows = []
        self.entry_exit_totals = {}
        self.refresh_entry_exit_analysis()
        self.entry_exit_status_label.setText("Selecione os SPEDs e clique em Processar.")
        self.statusBar().showMessage("Tela Analise Entrada e Saida limpa.")

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

    def clear_nfe_key_extract_page(self) -> None:
        self.statusBar().showMessage("Limpando tela Extrair chave NFe... aguarde.")
        self.nfe_extract_status.setText("Limpando dados... aguarde.")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        try:
            self.nfe_extract_sped_input.clear()
            self.nfe_extract_scope_combo.setCurrentText("Todos")
            self.nfe_extract_all_rows = []
            self.nfe_extract_rows = []
            self.nfe_extract_table.setUpdatesEnabled(False)
            self.nfe_extract_table.setRowCount(0)
            self.nfe_extract_table.clearContents()
            self.nfe_extract_table.setUpdatesEnabled(True)
            self.nfe_extract_progress.setValue(0)
            self.nfe_extract_status.setText("Tela limpa. Selecione os SPEDs, escolha o tipo e clique em Carregar.")
            self.statusBar().showMessage("Tela Extrair chave NFe limpa.")
        finally:
            QApplication.restoreOverrideCursor()

    def process_nfe_key_extract(self) -> None:
        sped_paths = [path for path in parse_selected_paths(self.nfe_extract_sped_input.text()) if path.exists() and path.is_file()]
        if not sped_paths:
            QMessageBox.warning(self, "Extrair chave NFe", "Selecione pelo menos um arquivo SPED valido.")
            return
        self.nfe_extract_all_rows = []
        self.nfe_extract_rows = []
        self.nfe_extract_table.setRowCount(0)
        self.nfe_extract_progress.setValue(0)
        self.nfe_extract_status.setText("Iniciando leitura dos SPEDs... aguarde.")
        self.statusBar().showMessage("Extraindo chaves NFe...")
        self.set_nfe_extract_busy(True)

        self.nfe_extract_thread = QThread(self)
        self.nfe_extract_worker = NFeKeyExtractWorker(sped_paths)
        self.nfe_extract_worker.moveToThread(self.nfe_extract_thread)
        self.nfe_extract_thread.started.connect(self.nfe_extract_worker.run)
        self.nfe_extract_worker.progress.connect(self.update_nfe_extract_progress)
        self.nfe_extract_worker.finished.connect(self.handle_nfe_extract_finished)
        self.nfe_extract_worker.failed.connect(self.handle_nfe_extract_error)
        self.nfe_extract_worker.finished.connect(self.nfe_extract_thread.quit)
        self.nfe_extract_worker.failed.connect(self.nfe_extract_thread.quit)
        self.nfe_extract_thread.finished.connect(self.nfe_extract_worker.deleteLater)
        self.nfe_extract_thread.finished.connect(self.nfe_extract_thread.deleteLater)
        self.nfe_extract_thread.finished.connect(lambda: setattr(self, "nfe_extract_worker", None))
        self.nfe_extract_thread.finished.connect(lambda: setattr(self, "nfe_extract_thread", None))
        self.nfe_extract_thread.start()

    def set_nfe_extract_busy(self, busy: bool) -> None:
        self.nfe_extract_scope_combo.setEnabled(not busy)
        self.nfe_extract_table.setEnabled(not busy)
        QApplication.setOverrideCursor(Qt.WaitCursor) if busy else QApplication.restoreOverrideCursor()

    def update_nfe_extract_progress(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.nfe_extract_progress.setValue(percent)
        self.nfe_extract_status.setText(f"{message} ({percent}%)")

    def handle_nfe_extract_finished(self, rows: list[dict[str, object]]) -> None:
        self.nfe_extract_all_rows = rows
        self.apply_nfe_key_filter()
        self.nfe_extract_progress.setValue(100)
        self.set_nfe_extract_busy(False)
        if not self.nfe_extract_all_rows:
            QMessageBox.information(self, "Extrair chave NFe", "Nenhuma chave NFe valida foi encontrada nos arquivos selecionados.")

    def handle_nfe_extract_error(self, message: str) -> None:
        self.set_nfe_extract_busy(False)
        self.nfe_extract_progress.setValue(0)
        self.nfe_extract_status.setText(f"Falha na extracao: {message}")
        self.statusBar().showMessage("Falha ao extrair chaves NFe.")
        QMessageBox.critical(self, "Extrair chave NFe", f"Falha ao extrair chaves NFe.\n\n{message}")

    def apply_nfe_key_filter(self) -> None:
        selected_scope = self.nfe_extract_scope_combo.currentText().strip()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.nfe_extract_status.setText("Aplicando filtro... aguarde.")
        QApplication.processEvents()
        try:
            if selected_scope == "Todos":
                self.nfe_extract_rows = list(self.nfe_extract_all_rows)
            else:
                self.nfe_extract_rows = [row for row in self.nfe_extract_all_rows if str(row.get("operation_type", "")) == selected_scope]
            self.refresh_nfe_key_extract_table()
            total_documents = len(self.nfe_extract_rows)
            total_all = len(self.nfe_extract_all_rows)
            self.nfe_extract_status.setText(
                f"{total_documents} documento(s) com chave NFe exibido(s) de {total_all} carregado(s)."
            )
            self.statusBar().showMessage(f"Filtro aplicado ({selected_scope}): {total_documents} registro(s).")
        finally:
            QApplication.restoreOverrideCursor()

    def refresh_nfe_key_extract_table(self) -> None:
        rows = [
            [
                row.get("document_type", ""),
                row.get("operation_type", ""),
                row.get("document_key", ""),
                row.get("document_number", ""),
                row.get("document_series", ""),
                row.get("document_date", ""),
                row.get("participant_name", ""),
                row.get("participant_tax_id", ""),
                row.get("file_name", ""),
            ]
            for row in self.nfe_extract_rows
        ]
        self.set_table_rows(self.nfe_extract_table, rows)

    def export_nfe_key_extract_rows(self) -> None:
        if not self.nfe_extract_rows:
            QMessageBox.warning(self, "Extrair chave NFe", "Nao ha dados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Exportar chaves NFe",
            "chaves_nfe_extraidas.xlsx",
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        headers = ["Arquivo SPED", "Tipo Documento", "Operacao", "Chave", "Documento", "Serie", "Modelo", "Data", "Participante", "CNPJ/CPF"]
        rows = [
            [
                row.get("file_name", ""),
                row.get("document_type", ""),
                row.get("operation_type", ""),
                row.get("document_key", ""),
                row.get("document_number", ""),
                row.get("document_series", ""),
                row.get("document_model", ""),
                row.get("document_date", ""),
                row.get("participant_name", ""),
                row.get("participant_tax_id", ""),
            ]
            for row in self.nfe_extract_rows
        ]
        nfe_rows = [row for row in rows if str(row[1]) == "NFe"]
        nfse_rows = [row for row in rows if str(row[1]) == "NFSe"]
        output_path = Path(output)
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                csv_rows = [
                    row + [row_origin_keys[index] if index < len(row_origin_keys) else ""]
                    for index, row in enumerate(rows)
                ]
                write_simple_csv_file(output_path, headers, csv_rows)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                sheets: list[tuple[str, list[str], list[list[object]], dict[str, object]]] = []
                if nfe_rows:
                    sheets.append(("NFe", headers, nfe_rows, {"include_total": False}))
                if nfse_rows:
                    sheets.append(("NFSe", headers, nfse_rows, {"include_total": False}))
                if not sheets:
                    sheets.append(("Chaves", headers, rows, {"include_total": False}))
                write_simple_excel_workbook(output_path, sheets)
            self.handle_export_success("Exportar chaves NFe", output_path, "Chaves NFe exportadas")
        except Exception as exc:
            self.handle_export_failure("Exportar chaves NFe", "chaves NFe", exc)

    def export_entry_filter(self) -> None:
        if not self.filtered_entry_rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar consulta filtrada",
            self.build_export_filename_with_company("consulta_entradas_filtrada.xlsx"),
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
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Consulta Entradas", headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar consulta", output_path, "Consulta exportada")
        except Exception as exc:
            self.handle_export_failure("Exportar consulta", "consulta", exc)

    def export_entry_exit_analysis(self) -> None:
        if not self.entry_exit_excel_rows:
            QMessageBox.warning(self, "Analise Entrada e Saida", "Nao ha analise para exportar.")
            return
        output, _selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar analise Entrada e Saida",
            "analise_entrada_saida.xlsx",
            "Arquivo Excel (*.xlsx)",
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        try:
            write_entry_exit_analysis_excel(output_path, self.entry_exit_excel_rows)
            self.handle_export_success("Analise Entrada e Saida", output_path, "Analise exportada")
        except Exception as exc:
            self.handle_export_failure("Analise Entrada e Saida", "analise", exc)

    def export_catalog_table_filtered(self, table: QTableWidget, context_name: str = "cadastros") -> None:
        headers = [
            table.horizontalHeaderItem(column_index).text() if table.horizontalHeaderItem(column_index) else f"Coluna {column_index + 1}"
            for column_index in range(table.columnCount())
        ]
        rows: list[list[object]] = []
        for row_index in range(table.rowCount()):
            if table.isRowHidden(row_index):
                continue
            rows.append(
                [
                    table.item(row_index, column_index).text() if table.item(row_index, column_index) else ""
                    for column_index in range(table.columnCount())
                ]
            )
        if not rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar consulta filtrada",
            f"{self.slugify_filename(context_name)}_filtrado.xlsx",
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        rows_to_write = rows
        output_path = Path(output)
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Consulta Cadastros", headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar consulta", output_path, "Consulta exportada")
        except Exception as exc:
            self.handle_export_failure("Exportar consulta", "consulta", exc)

    def export_product_page_filtered(self) -> None:
        headers = list(getattr(self, "product_page_full_headers", []))
        rows: list[list[object]] = []
        export_rows_by_id = getattr(self, "product_page_export_rows_by_id", {})
        for row_id in getattr(self, "product_page_visible_ids", []):
            export_row = export_rows_by_id.get(row_id)
            if export_row:
                rows.append(list(export_row))
        if not rows:
            QMessageBox.warning(self, "Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        selected_company_id = int(self.product_page_filter_company_combo.currentData() or 0) if hasattr(self, "product_page_filter_company_combo") else 0
        output, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Salvar produtos filtrados",
            self.build_export_filename_with_company("produtos_filtrado.xlsx", selected_company_id),
            "Arquivo Excel (*.xlsx);;Arquivo CSV (*.csv)",
        )
        if not output:
            return
        output_path = Path(output)
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                sheets: list[tuple[str, list[str], list[list[object]], dict[str, object]]] = []
                sheets.append(("BASE_COMPLETA", headers, rows, {"include_total": False}))

                fornecedor_index = headers.index("Fornecedor") if "Fornecedor" in headers else -1
                if fornecedor_index >= 0:
                    grouped_rows: dict[str, list[list[object]]] = {}
                    for row in rows:
                        supplier_name = str(row[fornecedor_index] if fornecedor_index < len(row) else "").strip() or "FORNECEDOR_SEM_NOME"
                        grouped_rows.setdefault(supplier_name, []).append(row)

                    used_sheet_names: set[str] = {"BASE_COMPLETA"}
                    for supplier_name, supplier_rows in grouped_rows.items():
                        sheet_name = supplier_name[:31] or "FORNECEDOR"
                        if sheet_name in used_sheet_names:
                            suffix = 2
                            base_name = sheet_name[:28]
                            while f"{base_name}_{suffix}" in used_sheet_names:
                                suffix += 1
                            sheet_name = f"{base_name}_{suffix}"
                        used_sheet_names.add(sheet_name)
                        sheets.append((sheet_name, headers, supplier_rows, {"include_total": False}))

                write_simple_excel_workbook(output_path, sheets)
            self.handle_export_success("Exportacao Produtos", output_path, "Exportacao concluida")
        except Exception as exc:
            self.handle_export_failure("Exportacao Produtos", "produtos", exc)

    def export_sale_filter(self) -> None:
        self.export_consultation_rows(
            self.filtered_sale_rows,
            self.build_export_filename_with_company("consulta_saidas_filtrada.xlsx"),
            "Consulta Saidas",
        )

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
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [(f"Consulta {operation_type}"[:31], headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar consulta", output_path, "Consulta PIS/COFINS exportada")
        except Exception as exc:
            self.handle_export_failure("Exportar consulta", "consulta PIS/COFINS", exc)

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
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Resumo XML", headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar XML", output_path, "Resumo XML exportado")
        except Exception as exc:
            self.handle_export_failure("Exportar XML", "resumo XML", exc)

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
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Creditos XML", headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar XML", output_path, "Creditos XML exportados")
        except Exception as exc:
            self.handle_export_failure("Exportar XML", "creditos XML", exc)

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
        try:
            if output_path.suffix.lower() == ".csv":
                csv_rows = [
                    row + [row_origin_keys[index] if index < len(row_origin_keys) else ""]
                    for index, row in enumerate(rows)
                ]
                write_simple_csv_file(output_path, headers, csv_rows)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [("Comparacao", headers, rows, {"include_total": False})])
            self.handle_export_success("Exportar comparacao", output_path, "Comparacao exportada")
        except Exception as exc:
            self.handle_export_failure("Exportar comparacao", "comparacao", exc)

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
        try:
            if selected_filter.startswith("Arquivo CSV") or output_path.suffix.lower() == ".csv":
                if output_path.suffix.lower() != ".csv":
                    output_path = output_path.with_suffix(".csv")
                write_simple_csv_file(output_path, headers, rows_to_write)
            else:
                if output_path.suffix.lower() != ".xlsx":
                    output_path = output_path.with_suffix(".xlsx")
                write_simple_excel_workbook(output_path, [(sheet_name, headers, rows_to_write, {"include_total": False})])
            self.handle_export_success("Exportar consulta", output_path, "Consulta exportada")
        except Exception as exc:
            self.handle_export_failure("Exportar consulta", "consulta", exc)

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
        runtime_rule_rows = [
            {
                "operation_type": operation_type,
                "cst_icms": row[0] if len(row) > 0 else "",
                "cfop": row[1] if len(row) > 1 else "",
                "icms_rate": self.decimal_value(row[2] if len(row) > 2 else Decimal("0")),
                "icms_value": self.decimal_value(row[5] if len(row) > 5 else Decimal("0")),
                "sale_value": self.decimal_value(row[8] if len(row) > 8 else Decimal("0")),
                "base_icms": self.decimal_value(row[9] if len(row) > 9 else Decimal("0")),
            }
            for row in export_rows
        ]
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, width, height)
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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
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
        self.ensure_full_header_visibility(table)

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
            self.handle_export_success("Exportar popup", output_path, "Popup exportado")
        except Exception as exc:
            self.handle_export_failure("Exportar popup", "popup", exc)

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
            if not self.popup_header_should_totalize(header):
                continue
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

    def popup_header_should_totalize(self, header: str) -> bool:
        normalized = self.normalize_search_text(header)
        # Aliquotas e percentuais nao devem ser somados na linha de total.
        if "aliq" in normalized or "aliquota" in normalized or "%" in str(header):
            return False
        return True

    def slugify_filename(self, value: str) -> str:
        normalized = self.normalize_search_text(value).replace(" ", "_")
        cleaned = "".join(char for char in normalized if char.isalnum() or char in {"_", "-"}).strip("_")
        return cleaned or "popup"

    def build_export_filename_with_company(self, base_name: str, preferred_company_id: int = 0) -> str:
        stem = Path(str(base_name or "").strip() or "exportacao").stem
        suffix = Path(str(base_name or "").strip() or "exportacao.xlsx").suffix or ".xlsx"
        try:
            companies = self.mysql_repo.list_companies(self.environment)
        except Exception:
            companies = []
        selected_row: dict[str, object] | None = None
        if preferred_company_id:
            selected_row = next((row for row in companies if int(row.get("id") or 0) == int(preferred_company_id)), None)
        if selected_row is None and len(companies) == 1:
            selected_row = companies[0]
        if not selected_row:
            return f"{self.slugify_filename(stem)}{suffix}"
        company_name = str(selected_row.get("nome", "")).strip()
        company_cnpj = self.digits_only(str(selected_row.get("cnpj", "")).strip())
        if not company_name and not company_cnpj:
            return f"{self.slugify_filename(stem)}{suffix}"
        company_part = self.slugify_filename(company_name) if company_name else "empresa"
        cnpj_part = company_cnpj or "sem_cnpj"
        return f"{self.slugify_filename(stem)}_{company_part}_{cnpj_part}{suffix}"

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
        row_limit = 5000
        non_empty_sections = []
        for name, headers, rows in sections:
            if not rows:
                continue
            display_rows = rows
            if len(rows) > row_limit:
                omitted = len(rows) - row_limit
                message_row = [""] * max(1, len(headers))
                message_row[-1] = f"{omitted} linha(s) omitida(s) para preservar desempenho da tela."
                display_rows = [*rows[:row_limit], message_row]
            non_empty_sections.append((name, headers, display_rows))
        if not non_empty_sections:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, width, height)
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
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.horizontalHeader().setStretchLastSection(True)
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    item.setForeground(QColor(COLORS["text"]))
                    if not (column_index < len(headers) and headers[column_index].lower() in {"descricao", "fornecedor", "chave", "motivo"}):
                        item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(row_index, column_index, item)
            self.apply_popup_column_widths(table, headers)
            self.enable_table_sorting(table)
            layout.addWidget(table, 1)
        if footer_text:
            metric_pairs: list[tuple[str, str]] = []
            for raw_part in [part.strip() for part in footer_text.split("    ") if part.strip()]:
                if ":" not in raw_part:
                    continue
                label, value = raw_part.split(":", 1)
                metric_pairs.append((label.strip(), value.strip()))
            if metric_pairs:
                cards_grid = QGridLayout()
                cards_grid.setHorizontalSpacing(8)
                for index, (label, value) in enumerate(metric_pairs):
                    card, value_label = self.create_metric_card_with_label(label, value)
                    value_label.setText(value)
                    cards_grid.addWidget(card, 0, index)
                layout.addLayout(cards_grid)
            else:
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
            self.handle_export_success("Exportar popup", output_path, "Popup exportado")
        except Exception as exc:
            self.handle_export_failure("Exportar popup", "popup", exc)

    def handle_export_success(self, title: str, output_path: Path, status_message: str) -> None:
        self.statusBar().showMessage(f"{status_message}: {output_path}")
        message_box = QMessageBox(self)
        message_box.setIcon(QMessageBox.Information)
        message_box.setWindowTitle(title)
        message_box.setText(f"Exportacao concluida com sucesso.\n\nDeseja abrir o arquivo?\n{output_path}")
        sim_button = message_box.addButton("Sim", QMessageBox.YesRole)
        message_box.addButton("Nao", QMessageBox.NoRole)
        message_box.exec()
        if message_box.clickedButton() == sim_button:
            os.startfile(str(output_path))

    def handle_export_failure(self, title: str, export_name: str, error: Exception) -> None:
        self.statusBar().showMessage(f"Falha na exportacao de {export_name}.")
        QMessageBox.critical(self, title, f"Falha ao exportar {export_name}.\n\n{error}")

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
        self.open_operation_summary_for_rows(self.filtered_entry_rows, "Entrada", "Resumo Entradas")

    def open_entry_abc_popup(self) -> None:
        _periods, headers, _display_rows, export_rows = build_product_monthly_linear_dataset(self.filtered_entry_rows, "Entrada")
        self.open_product_survey_popup_with_cards("Levantamento por Produto - Entrada", headers, export_rows)

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
        selected_periods = {label for label, check in self.period_checks.items() if check.isChecked()}
        sale_rows_for_apuracao = [
            row
            for row in self.sale_rows
            if not selected_periods or str(self.first_row_value(row, "period", "period_label")) in selected_periods
        ]
        exit_totals = {
            "icms_value": sum((self.decimal_value(row.get("icms_value")) for row in sale_rows_for_apuracao), Decimal("0")),
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
        self.open_apuracao_popup_with_cards("Apuracao do ICMS", rows, entry_totals["icms_value"], exit_totals["icms_value"])

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

    def handle_entry_table_double_click(self, item: QTableWidgetItem) -> None:
        row_index = item.row()
        column_index = item.column()
        if row_index < 0 or row_index >= len(self.filtered_entry_rows):
            return
        selected_row = self.filtered_entry_rows[row_index]
        if column_index == 4 and int(self.first_row_value(selected_row, "supplier_count", default=0) or 0) > 1:
            self.open_consultation_supplier_popup(selected_row)
            return
        self.open_row_detail_popup(selected_row, "Entrada")

    def handle_sale_table_double_click(self, item: QTableWidgetItem) -> None:
        row_index = item.row()
        column_index = item.column()
        if row_index < 0 or row_index >= len(self.filtered_sale_rows):
            return
        selected_row = self.filtered_sale_rows[row_index]
        if column_index == 4 and int(self.first_row_value(selected_row, "supplier_count", default=0) or 0) > 1:
            self.open_consultation_supplier_popup(selected_row)
            return
        self.open_row_detail_popup(selected_row, "Saida")

    def open_consultation_supplier_popup(self, summary_row: dict[str, object]) -> None:
        supplier_details = summary_row.get("supplier_details")
        if (not isinstance(supplier_details, list) or not supplier_details) and isinstance(summary_row.get("launch_details"), list):
            supplier_details = self.build_supplier_details_from_launch_details(summary_row)
        if not isinstance(supplier_details, list) or not supplier_details:
            QMessageBox.warning(self, "Fornecedores", "Nao ha fornecedores disponiveis para este produto.")
            return
        code = str(summary_row.get("code", "")).strip() or "produto"
        description = str(summary_row.get("description", "")).strip()
        headers = ["Fornecedor", "CPF/CNPJ", "Periodos", "Docs", "Lanc.", "Valor Operacao", "Base ICMS", "Valor ICMS"]
        rows: list[list[object]] = []
        for supplier in supplier_details:
            if not isinstance(supplier, dict):
                continue
            periods = supplier.get("periods")
            document_keys = supplier.get("document_keys")
            rows.append(
                [
                    supplier.get("name", ""),
                    supplier.get("tax_id", ""),
                    len(periods) if isinstance(periods, (set, list, tuple)) else 0,
                    len(document_keys) if isinstance(document_keys, (set, list, tuple)) else 0,
                    int(supplier.get("launch_count", 0) or 0),
                    self.decimal_value(supplier.get("sale_value")),
                    self.decimal_value(supplier.get("base_icms")),
                    self.decimal_value(supplier.get("icms_value")),
                ]
            )
        rows.sort(key=lambda row: str(row[0]).upper())
        footer = f"Fornecedores: {len(rows)}"
        if description:
            footer = f"Produto {code} - {description}\n\n{footer}"
        self.open_table_popup(f"Fornecedores - {code}", headers, rows, 1180, 640, footer)

    def build_supplier_details_from_launch_details(self, summary_row: dict[str, object]) -> list[dict[str, object]]:
        launch_details = summary_row.get("launch_details")
        if not isinstance(launch_details, list):
            return []
        supplier_buckets: dict[str, dict[str, object]] = {}
        for launch_detail in launch_details:
            if not isinstance(launch_detail, dict):
                continue
            supplier_name = str(launch_detail.get("participant_name", "")).strip()
            supplier_tax_id = str(launch_detail.get("participant_tax_id", "")).strip()
            supplier_key = supplier_name or supplier_tax_id
            if not supplier_key:
                continue
            bucket = supplier_buckets.setdefault(
                supplier_key,
                {
                    "name": supplier_name,
                    "tax_id": supplier_tax_id,
                    "periods": set(),
                    "document_keys": set(),
                    "launch_count": 0,
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            period_text = str(launch_detail.get("period", "")).strip()
            if period_text:
                bucket["periods"].add(period_text)
            document_identity = str(launch_detail.get("document_key", "") or launch_detail.get("document_number", "")).strip()
            if document_identity:
                bucket["document_keys"].add(document_identity)
            bucket["launch_count"] = int(bucket["launch_count"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(
                self.first_row_value(launch_detail, "sale_value", "total_operation_value", "operation_value")
            )
            bucket["base_icms"] = Decimal(bucket["base_icms"]) + self.decimal_value(launch_detail.get("base_icms"))
            bucket["icms_value"] = Decimal(bucket["icms_value"]) + self.decimal_value(launch_detail.get("icms_value"))
        return sorted(supplier_buckets.values(), key=lambda item: str(item.get("name", "")).upper())

    def open_row_detail_popup(self, summary_row: dict[str, object], operation_type: str) -> None:
        launch_details = summary_row.get("launch_details")
        details = launch_details if isinstance(launch_details, list) and launch_details else [summary_row]
        details = [detail for detail in details if isinstance(detail, dict)]
        if not details:
            QMessageBox.warning(self, "Detalhamento", "Nao ha detalhes disponiveis para esta linha.")
            return

        code = str(summary_row.get("code", "")).strip() or str(self.first_row_value(summary_row, "codigo", default="")).strip()
        period = str(self.first_row_value(summary_row, "period", "period_label")).strip()
        description = str(summary_row.get("description", "")).strip()
        title = f"Detalhamento - {code or 'produto'}{f' - {period}' if period else ''}"

        detail_payloads: list[dict[str, object]] = []
        total_quantity = Decimal("0")
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        document_keys: set[str] = set()
        for detail in details:
            operation_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(detail.get("base_icms"))
            icms_value = self.decimal_value(detail.get("icms_value"))
            quantity = self.decimal_value(detail.get("quantity"))
            total_quantity += quantity
            total_operation += operation_value
            total_base += base_icms
            total_icms += icms_value
            document_key = str(detail.get("document_key", "")).strip()
            if document_key:
                document_keys.add(document_key)
            detail_payloads.append(dict(detail))
        detail_payloads.sort(key=lambda row: (str(row.get("document_date", row.get("date", ""))), str(row.get("document_number", "")), str(row.get("document_key", "")), str(row.get("item_number", ""))))
        rows = []
        for detail in detail_payloads:
            operation_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(detail.get("base_icms"))
            icms_value = self.decimal_value(detail.get("icms_value"))
            rows.append(
                [
                    detail.get("document_date", detail.get("date", "")),
                    detail.get("document_number", ""),
                    detail.get("participant_name", ""),
                    detail.get("document_key", ""),
                    detail.get("item_number", ""),
                    detail.get("cest", ""),
                    detail.get("cfop", ""),
                    self.first_row_value(detail, "cst_icms", "cst"),
                    self.decimal_value(detail.get("icms_rate")),
                    compute_display_icms_rate(self.decimal_value(detail.get("icms_rate")), operation_value, base_icms, icms_value),
                ]
            )

        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, 1320, 760)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)
        subtitle = QLabel(
            f"Produto {code} - {description}" if description else f"Produto {code}"
        )
        subtitle.setObjectName("muted")
        layout.addWidget(subtitle)

        table = QTableWidget()
        headers = ["Data", "Documento", "Fornecedor", "Chave", "Item", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva"]
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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setMinimumSectionSize(70)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.UserRole, row_index)
                item.setForeground(QColor(COLORS["text"]))
                if column_index not in {2, 3}:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        self.apply_popup_column_widths(table, headers)
        self.enable_table_sorting(table)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_runtime_rule_context_menu_for_payloads(table, detail_payloads, operation_type, pos, include_document=True)
        )
        layout.addWidget(table, 1)

        ratio = (total_base * Decimal("100") / total_operation).quantize(Decimal("0.01")) if total_operation > 0 else Decimal("0.00")
        cards_grid = QGridLayout()
        cards_grid.setHorizontalSpacing(8)
        cards = (
            ("Lancamentos", str(len(rows))),
            ("Quantidade", self.format_number(total_quantity)),
            ("Valor Operacao", self.format_number(total_operation)),
            ("Base ICMS", self.format_number(total_base)),
            ("% Base/Oper", self.format_number(ratio)),
            ("Valor ICMS", self.format_number(total_icms)),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            cards_grid.addWidget(card, 0, index)
        layout.addLayout(cards_grid)

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, headers, rows, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, headers, rows, "csv")))
        actions.addStretch()
        footer = QLabel(f"Documentos: {len(document_keys)} | Operacao: {operation_type}")
        footer.setObjectName("muted")
        actions.addWidget(footer)
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def open_entry_docs_popup(self) -> None:
        self.open_documents_popup_for_details(self.get_filtered_launch_details(), "Entradas")

    def open_sale_operation_summary_popup(self) -> None:
        self.open_operation_summary_for_rows(self.filtered_sale_rows, "Saida", "Resumo Saidas")

    def open_confronto_entry(self) -> None:
        if not self.filtered_entry_rows:
            QMessageBox.warning(self, "Confronto", "Nao ha dados de entrada. Processe o SPED primeiro.")
            return
        from app.ui_qt.confronto_dialog import ConfrontoDialog
        from PySide6.QtWidgets import QDialog
        dlg = ConfrontoDialog(
            rows=self.filtered_entry_rows,
            operation_type="Entrada",
            repository=self.mysql_repo,
            environment=self.environment,
            parent=self,
        )
        if dlg.exec() == QDialog.Accepted and dlg.accepted_rules:
            self.generate_adjusted_sped_from_consultation("Entrada", override_rules=dlg.accepted_rules)

    def open_confronto_sale(self) -> None:
        if not self.filtered_sale_rows:
            QMessageBox.warning(self, "Confronto", "Nao ha dados de saida. Processe o SPED primeiro.")
            return
        from app.ui_qt.confronto_dialog import ConfrontoDialog
        from PySide6.QtWidgets import QDialog
        dlg = ConfrontoDialog(
            rows=self.filtered_sale_rows,
            operation_type="Saida",
            repository=self.mysql_repo,
            environment=self.environment,
            parent=self,
        )
        if dlg.exec() == QDialog.Accepted and dlg.accepted_rules:
            self.generate_adjusted_sped_from_consultation("Saida", override_rules=dlg.accepted_rules)

    def open_sale_abc_popup(self) -> None:
        _periods, headers, _display_rows, export_rows = build_product_monthly_linear_dataset(self.filtered_sale_rows, "Saida")
        self.open_product_survey_popup_with_cards("Levantamento por Produto - Saida", headers, export_rows)

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
        self.open_apuracao_popup_with_cards("Apuracao do ICMS - Saidas", rows, entry_totals["icms_value"], exit_totals["icms_value"])

    def open_apuracao_popup_with_cards(
        self,
        title: str,
        rows: list[list[object]] | list[tuple[object, ...]],
        entrada_credito: Decimal,
        saida_debito: Decimal,
    ) -> None:
        if not rows:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        export_rows = [list(row) for row in rows]
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, 1180, 700)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)

        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Descricao", "Valor R$"])
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
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for row_index, row in enumerate(export_rows):
            for column_index, value in enumerate(row[:2]):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if column_index == 1:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        self.enable_table_sorting(table)
        layout.addWidget(table, 1)

        resumo_grid = QGridLayout()
        resumo_grid.setHorizontalSpacing(8)
        icms_recolher = max(self.decimal_value(saida_debito) - self.decimal_value(entrada_credito), Decimal("0"))
        credito_restante = max(self.decimal_value(entrada_credito) - self.decimal_value(saida_debito), Decimal("0"))
        cards = (
            ("Entradas (credito)", self.format_number(self.decimal_value(entrada_credito))),
            ("Saidas (debito)", self.format_number(self.decimal_value(saida_debito))),
            ("ICMS a Recolher", self.format_number(icms_recolher)),
            ("Credito Restante", self.format_number(credito_restante)),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            resumo_grid.addWidget(card, 0, index)
        layout.addLayout(resumo_grid)

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, ["Descricao", "Valor R$"], export_rows, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, ["Descricao", "Valor R$"], export_rows, "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

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
        self.open_product_survey_popup_with_cards(f"Levantamento por Produto PIS/COFINS - {operation_type}", headers, export_rows)

    def open_product_survey_popup_with_cards(self, title: str, headers: list[str], rows: list[list[object]]) -> None:
        if not rows:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, 1460, 760)
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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.UserRole, row_index)
                item.setForeground(QColor(COLORS["text"]))
                if not (column_index < len(headers) and self.normalize_search_text(headers[column_index]) in {"descricao"}):
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        self.apply_popup_column_widths(table, headers)
        self.enable_table_sorting(table)
        layout.addWidget(table, 1)

        normalized_headers = [self.normalize_search_text(header) for header in headers]
        curve_index = next((index for index, header in enumerate(normalized_headers) if "curva" in header and "abc" in header), -1)
        code_index = next((index for index, header in enumerate(normalized_headers) if header == "codigo"), -1)
        operation_index = next((index for index, header in enumerate(normalized_headers) if "totaloperacao" in header), -1)
        base_index = next((index for index, header in enumerate(normalized_headers) if "totalbaseicm" in header), -1)
        icms_index = next((index for index, header in enumerate(normalized_headers) if "totalvlricm" in header), -1)
        products = {
            str(row[code_index]).strip()
            for row in rows
            if code_index >= 0 and code_index < len(row) and str(row[code_index]).strip()
        }
        total_operation = sum((self.decimal_value(row[operation_index]) for row in rows), Decimal("0")) if operation_index >= 0 else Decimal("0")
        curve_counts = {"A": 0, "B": 0, "C": 0}
        curve_totals = {"A": Decimal("0"), "B": Decimal("0"), "C": Decimal("0")}
        if curve_index >= 0:
            for row in rows:
                curve = str(row[curve_index]).strip().upper()
                if curve in curve_counts:
                    curve_counts[curve] += 1
                    if operation_index >= 0 and operation_index < len(row):
                        curve_totals[curve] += self.decimal_value(row[operation_index])
        a_share = (curve_totals["A"] * Decimal("100") / total_operation).quantize(Decimal("0.01")) if total_operation > 0 else Decimal("0.00")

        cards_grid = QGridLayout()
        cards = (
            ("Linhas", str(len(rows))),
            ("Produtos", str(len(products))),
            ("Produtos Curva A", str(curve_counts["A"])),
            ("Produtos Curva B", str(curve_counts["B"])),
            ("Produtos Curva C", str(curve_counts["C"])),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            cards_grid.addWidget(card, index // 5, index % 5)
        layout.addLayout(cards_grid)

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, headers, rows, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, headers, rows, "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

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
                    {"sale_value": Decimal("0"), "base_icms": Decimal("0"), "icms_value": Decimal("0"), "base_icms_st": Decimal("0"), "icms_st_value": Decimal("0"), "ipi_value": Decimal("0"), "document_keys": set(), "launch_count": 0, "details": []},
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
                bucket["details"].append(detail)
        rows = []
        detail_rows_by_index: dict[int, list[dict[str, object]]] = {}
        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for (cst, cfop, rate), values in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1], item[0][2])):
            sale_value = Decimal(values["sale_value"])
            base_icms = Decimal(values["base_icms"])
            icms_value = Decimal(values["icms_value"])
            rows.append((cst, cfop, rate, compute_display_icms_rate(rate, sale_value, base_icms, icms_value), Decimal(values["ipi_value"]), icms_value, Decimal(values["base_icms_st"]), Decimal(values["icms_st_value"]), sale_value, base_icms, (sale_value - base_icms).quantize(Decimal("0.01")), (sale_value - base_icms).quantize(Decimal("0.01")), len(values["document_keys"]), values["launch_count"]))
            detail_rows_by_index[len(rows) - 1] = [detail for detail in values.get("details", []) if isinstance(detail, dict)]
            total_sale += sale_value
            total_base += base_icms
            total_icms += icms_value
        headers = ["CST", "CFOP", "Aliq ICMS", "Aliq Efetiva", "Valor IPI", "Valor ICMS", "Base ICMS ST", "Valor ICMS ST", "Total Operacao", "Base ICMS", "Dif. Oper/Base", "Reducao BC", "Docs", "Lanc."]
        self.open_operation_summary_popup_with_cards(title, headers, rows, total_sale, total_base, total_icms, detail_rows_by_index, operation_type)

    def open_operation_summary_popup_with_cards(
        self,
        title: str,
        headers: list[str],
        rows: list[tuple[object, ...]] | list[list[object]],
        total_sale: Decimal,
        total_base: Decimal,
        total_icms: Decimal,
        detail_rows_by_index: dict[int, list[dict[str, object]]] | None = None,
        operation_type: str = "",
    ) -> None:
        if not rows:
            QMessageBox.warning(self, title, "Nao ha dados para os filtros atuais.")
            return
        export_rows = [list(row) for row in rows]
        runtime_rule_rows = [
            {
                "operation_type": operation_type,
                "cst_icms": row[0] if len(row) > 0 else "",
                "cfop": row[1] if len(row) > 1 else "",
                "icms_rate": self.decimal_value(row[2] if len(row) > 2 else Decimal("0")),
                "effective_rate": self.decimal_value(row[3] if len(row) > 3 else Decimal("0")),
                "ipi_value": self.decimal_value(row[4] if len(row) > 4 else Decimal("0")),
                "icms_value": self.decimal_value(row[5] if len(row) > 5 else Decimal("0")),
                "base_icms_st": self.decimal_value(row[6] if len(row) > 6 else Decimal("0")),
                "icms_st_value": self.decimal_value(row[7] if len(row) > 7 else Decimal("0")),
                "sale_value": self.decimal_value(row[8] if len(row) > 8 else Decimal("0")),
                "base_icms": self.decimal_value(row[9] if len(row) > 9 else Decimal("0")),
            }
            for row in export_rows
        ]
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, 1320, 760)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)

        subtitle = QLabel("Resumo por CST, CFOP e Aliquota")
        subtitle.setObjectName("muted")
        layout.addWidget(subtitle)
        filter_state_label = QLabel("Sem filtros ativos.")
        filter_state_label.setObjectName("muted")
        filter_state_label.setStyleSheet(f"color: {COLORS['muted']}; font-weight: 700;")
        layout.addWidget(filter_state_label)

        filters_grid = QGridLayout()
        filters_grid.setHorizontalSpacing(10)
        filters_grid.setVerticalSpacing(6)
        cst_filter_input = QLineEdit()
        cfop_filter_input = QLineEdit()
        aliq_filter_input = QLineEdit()
        search_filter_input = QLineEdit()
        cst_filter_input.setPlaceholderText("Ex.: 000, 020")
        cfop_filter_input.setPlaceholderText("Ex.: 1101, 1403")
        aliq_filter_input.setPlaceholderText("Ex.: 12, 18")
        search_filter_input.setPlaceholderText("CST, CFOP ou valores")
        filters_grid.addWidget(QLabel("Filtro CST"), 0, 0)
        filters_grid.addWidget(cst_filter_input, 1, 0)
        filters_grid.addWidget(QLabel("Filtro CFOP"), 0, 1)
        filters_grid.addWidget(cfop_filter_input, 1, 1)
        filters_grid.addWidget(QLabel("Filtro Aliquota"), 0, 2)
        filters_grid.addWidget(aliq_filter_input, 1, 2)
        filters_grid.addWidget(QLabel("Busca"), 0, 3)
        filters_grid.addWidget(search_filter_input, 1, 3)
        clear_filters_button = self.create_button("Limpar filtros", lambda: None)
        filters_grid.addWidget(clear_filters_button, 0, 4, 2, 1)
        filters_grid.setColumnStretch(0, 1)
        filters_grid.setColumnStretch(1, 1)
        filters_grid.setColumnStretch(2, 1)
        filters_grid.setColumnStretch(3, 2)
        layout.addLayout(filters_grid)

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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
        if detail_rows_by_index:
            table.itemDoubleClicked.connect(
                lambda item: self.handle_operation_summary_double_click(item, detail_rows_by_index, operation_type)
            )
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_operation_summary_runtime_rule_menu(table, runtime_rule_rows, detail_rows_by_index or {}, operation_type, pos)
        )
        self.apply_popup_column_widths(table, headers)
        self.enable_table_sorting(table)
        layout.addWidget(table, 1)

        cards_grid = QGridLayout()
        cards_grid.setHorizontalSpacing(8)
        card_labels: dict[str, QLabel] = {}
        cards = (("Linhas", "rows"), ("Total Operacao", "total_operation"), ("Base ICMS", "base_icms"), ("% Base/Oper", "base_ratio"), ("Valor ICMS", "icms_value"))
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, "0")
            card_labels[value] = value_label
            cards_grid.addWidget(card, 0, index)
        layout.addLayout(cards_grid)

        filtered_rows_state: dict[str, list[list[object]]] = {"rows": list(export_rows)}

        def parse_tokens(text: str) -> set[str]:
            return {part.strip() for part in text.split(",") if part.strip()}

        def apply_filters() -> None:
            cst_tokens = parse_tokens(cst_filter_input.text())
            cfop_tokens = parse_tokens(cfop_filter_input.text())
            aliq_tokens = parse_tokens(aliq_filter_input.text())
            search_text = self.normalize_search_text(search_filter_input.text())
            filtered_rows: list[list[object]] = []
            filtered_indices: list[int] = []
            for original_index, row in enumerate(export_rows):
                cst_value = str(row[0]).strip()
                cfop_value = str(row[1]).strip()
                aliq_value = str(row[2]).strip().replace(".", ",")
                if cst_tokens and cst_value not in cst_tokens:
                    continue
                if cfop_tokens and cfop_value not in cfop_tokens:
                    continue
                if aliq_tokens and all(token not in aliq_value for token in aliq_tokens):
                    continue
                if search_text:
                    searchable = self.normalize_search_text(" ".join(str(value) for value in row))
                    if search_text not in searchable:
                        continue
                filtered_rows.append(list(row))
                filtered_indices.append(original_index)

            table.setRowCount(len(filtered_rows))
            for row_index, row in enumerate(filtered_rows):
                for column_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    item.setData(Qt.UserRole, filtered_indices[row_index])
                    item.setForeground(QColor(COLORS["text"]))
                    item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(row_index, column_index, item)
            table.resizeColumnsToContents()

            if detail_rows_by_index:
                mapped_details: dict[int, list[dict[str, object]]] = {}
                for new_index, original_index in enumerate(filtered_indices):
                    mapped_details[new_index] = detail_rows_by_index.get(original_index, [])
                table.itemDoubleClicked.disconnect()
                table.itemDoubleClicked.connect(
                    lambda item: self.handle_operation_summary_double_click(item, mapped_details, operation_type)
                )

            filtered_rows_state["rows"] = filtered_rows
            active_filters: list[str] = []
            if cst_tokens:
                active_filters.append("CST")
            if cfop_tokens:
                active_filters.append("CFOP")
            if aliq_tokens:
                active_filters.append("Aliquota")
            if search_text:
                active_filters.append("Busca")
            filter_state_label.setText("Sem filtros ativos." if not active_filters else f"Filtros ativos: {', '.join(active_filters)}")

            total_operation_filtered = sum((self.decimal_value(row[8]) for row in filtered_rows), Decimal("0"))
            total_base_filtered = sum((self.decimal_value(row[9]) for row in filtered_rows), Decimal("0"))
            total_icms_filtered = sum((self.decimal_value(row[5]) for row in filtered_rows), Decimal("0"))
            ratio_filtered = (total_base_filtered * Decimal("100") / total_operation_filtered).quantize(Decimal("0.01")) if total_operation_filtered else Decimal("0.00")
            card_labels["rows"].setText(str(len(filtered_rows)))
            card_labels["total_operation"].setText(self.format_number(total_operation_filtered))
            card_labels["base_icms"].setText(self.format_number(total_base_filtered))
            card_labels["base_ratio"].setText(self.format_number(ratio_filtered))
            card_labels["icms_value"].setText(self.format_number(total_icms_filtered))

        cst_filter_input.textChanged.connect(lambda _text: apply_filters())
        cfop_filter_input.textChanged.connect(lambda _text: apply_filters())
        aliq_filter_input.textChanged.connect(lambda _text: apply_filters())
        search_filter_input.textChanged.connect(lambda _text: apply_filters())
        clear_filters_button.clicked.connect(lambda: (cst_filter_input.clear(), cfop_filter_input.clear(), aliq_filter_input.clear(), search_filter_input.clear()))
        apply_filters()

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, headers, filtered_rows_state["rows"], "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, headers, filtered_rows_state["rows"], "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def handle_operation_summary_double_click(
        self,
        item: QTableWidgetItem,
        detail_rows_by_index: dict[int, list[dict[str, object]]],
        operation_type: str,
    ) -> None:
        data_index = item.data(Qt.UserRole)
        try:
            row_index = int(data_index)
        except (TypeError, ValueError):
            row_index = item.row()
        details = detail_rows_by_index.get(row_index, [])
        if not details:
            return
        self.open_operation_summary_detail_popup(details, operation_type)

    def show_operation_summary_runtime_rule_menu(
        self,
        table: QTableWidget,
        runtime_rule_rows: list[dict[str, object]],
        detail_rows_by_index: dict[int, list[dict[str, object]]],
        operation_type: str,
        pos: object,
    ) -> None:
        item = table.itemAt(pos)
        if item is None:
            return
        data_index = item.data(Qt.UserRole)
        try:
            original_index = int(data_index)
        except (TypeError, ValueError):
            original_index = item.row()
        if original_index < 0 or original_index >= len(runtime_rule_rows):
            return
        table.selectRow(item.row())
        menu = QMenu(table)
        self.add_runtime_rule_actions_to_menu(menu, runtime_rule_rows[original_index], operation_type, include_document=False)
        if detail_rows_by_index:
            details = detail_rows_by_index.get(original_index, [])
            if details:
                menu.addSeparator()
                menu.addAction("Abrir Detalhamento", lambda current_details=details: self.open_operation_summary_detail_popup(current_details, operation_type))
        menu.exec(table.viewport().mapToGlobal(pos))

    def open_operation_summary_detail_popup(self, details: list[dict[str, object]], operation_type: str) -> None:
        rows: list[list[object]] = []
        detail_payloads: list[dict[str, object]] = []
        total_quantity = Decimal("0")
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for detail in details:
            operation_value = self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(detail.get("base_icms"))
            icms_value = self.decimal_value(detail.get("icms_value"))
            quantity = self.decimal_value(detail.get("quantity"))
            total_quantity += quantity
            total_operation += operation_value
            total_base += base_icms
            total_icms += icms_value
            detail_payloads.append(
                {
                    "document_date": detail.get("document_date", detail.get("date", "")),
                    "document_number": detail.get("document_number", ""),
                    "participant_name": detail.get("participant_name", ""),
                    "document_key": detail.get("document_key", ""),
                    "item_number": detail.get("item_number", ""),
                    "cest": detail.get("cest", ""),
                    "cfop": detail.get("cfop", ""),
                    "cst": self.first_row_value(detail, "cst_icms", "cst"),
                    "icms_rate": self.decimal_value(detail.get("icms_rate")),
                    "effective_rate": compute_display_icms_rate(
                        self.decimal_value(detail.get("icms_rate")),
                        operation_value,
                        base_icms,
                        icms_value,
                    ),
                    "code": detail.get("code", ""),
                    "description": detail.get("description", ""),
                    "ncm": detail.get("ncm", ""),
                    "document_series": detail.get("document_series", ""),
                    "document_model": detail.get("document_model", ""),
                    "quantity": quantity,
                    "operation_value": operation_value,
                    "base_icms": base_icms,
                    "icms_value": icms_value,
                }
            )
        detail_payloads.sort(
            key=lambda row: (
                str(row.get("document_date", "")),
                str(row.get("document_number", "")),
                str(row.get("document_key", "")),
                str(row.get("item_number", "")),
            )
        )
        for payload in detail_payloads:
            rows.append(
                [
                    payload.get("document_date", ""),
                    payload.get("document_number", ""),
                    payload.get("participant_name", ""),
                    payload.get("document_key", ""),
                    payload.get("item_number", ""),
                    payload.get("cest", ""),
                    payload.get("cfop", ""),
                    payload.get("cst", ""),
                    payload.get("icms_rate", Decimal("0")),
                    payload.get("effective_rate", Decimal("0")),
                    payload.get("code", ""),
                    payload.get("description", ""),
                    payload.get("ncm", ""),
                    payload.get("document_series", ""),
                    payload.get("document_model", ""),
                ]
            )

        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(f"Detalhamento - {operation_type}s")
        self.resize_dialog_to_screen(dialog, 1420, 760)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        title_label = QLabel(f"Detalhamento - {operation_type}s")
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)

        headers = ["Data", "Documento", "Fornecedor", "Chave", "Item", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Codigo", "Descricao", "NCM", "Serie", "Modelo"]
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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if column_index not in {2, 3, 11}:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        table.itemDoubleClicked.connect(lambda item: self.open_invoice_mirror_from_detail_row(item, detail_payloads, operation_type))
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_runtime_rule_context_menu_for_payloads(table, detail_payloads, operation_type, pos, include_document=True)
        )
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        for column_index in range(table.columnCount()):
            header.setSectionResizeMode(column_index, QHeaderView.ResizeToContents)
        self.enable_table_sorting(table)
        layout.addWidget(table, 1)

        ratio = (total_base * Decimal("100") / total_operation).quantize(Decimal("0.01")) if total_operation > 0 else Decimal("0.00")
        cards_grid = QGridLayout()
        cards = (
            ("Lancamentos", str(len(rows))),
            ("Quantidade", self.format_number(total_quantity)),
            ("Valor Operacao", self.format_number(total_operation)),
            ("Base ICMS", self.format_number(total_base)),
            ("% Base/Oper", self.format_number(ratio)),
            ("Valor ICMS", self.format_number(total_icms)),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            cards_grid.addWidget(card, 0, index)
        layout.addLayout(cards_grid)

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset("Detalhamento", headers, rows, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset("Detalhamento", headers, rows, "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def open_invoice_mirror_from_detail_row(
        self,
        item: QTableWidgetItem,
        detail_rows: list[dict[str, object]],
        operation_type: str,
    ) -> None:
        row_index = item.row()
        if not detail_rows:
            return
        selected = detail_rows[row_index] if 0 <= row_index < len(detail_rows) else detail_rows[0]
        selected_document = str(selected.get("document_number", ""))
        selected_key = str(selected.get("document_key", ""))
        selected_item = str(selected.get("item_number", "")).strip()
        selected_code = str(selected.get("code", "")).strip()
        selected_series = str(selected.get("document_series", "")).strip()

        source_pool = self.get_filtered_launch_details() if operation_type == "Entrada" else self.get_filtered_sale_launch_details()
        invoice_details = [
            row
            for row in source_pool
            if isinstance(row, dict)
            and (
                (selected_key and str(row.get("document_key", "")).strip() == selected_key)
                or (
                    selected_document
                    and str(row.get("document_number", "")).strip() == selected_document
                    and (not selected_series or str(row.get("document_series", "")).strip() == selected_series)
                )
            )
        ]
        if not invoice_details:
            invoice_details = [
                {
                    "item_number": selected.get("item_number", ""),
                    "code": selected.get("code", ""),
                    "description": selected.get("description", ""),
                    "ncm": selected.get("ncm", ""),
                    "cest": selected.get("cest", ""),
                    "cfop": selected.get("cfop", ""),
                    "cst_icms": selected.get("cst", ""),
                    "icms_rate": selected.get("icms_rate", Decimal("0")),
                    "document_number": selected_document,
                    "document_key": selected_key,
                    "document_date": selected.get("document_date", ""),
                    "document_series": selected.get("document_series", ""),
                    "document_model": selected.get("document_model", ""),
                    "participant_name": selected.get("participant_name", ""),
                    "participant_tax_id": "",
                    "quantity": selected.get("quantity", Decimal("0")),
                    "sale_value": selected.get("operation_value", Decimal("0")),
                    "base_icms": selected.get("base_icms", Decimal("0")),
                    "icms_value": selected.get("icms_value", Decimal("0")),
                }
            ]
        title = f"Espelho da Nota - {selected_document or selected_key}"
        headers = ["Item", "Codigo", "Descricao", "NCM", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva"]
        rows = []
        highlight_rows: set[int] = set()
        sorted_invoice_details = sorted(
            invoice_details,
            key=lambda current: (
                str(current.get("item_number", "")),
                str(current.get("code", "")),
            ),
        )
        for index, row in enumerate(sorted_invoice_details):
            operation_value = self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(row.get("base_icms"))
            icms_value = self.decimal_value(row.get("icms_value"))
            rows.append(
                [
                    row.get("item_number", ""),
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("ncm", ""),
                    row.get("cest", ""),
                    row.get("cfop", ""),
                    self.first_row_value(row, "cst_icms", "cst"),
                    self.decimal_value(row.get("icms_rate")),
                    compute_display_icms_rate(self.decimal_value(row.get("icms_rate")), operation_value, base_icms, icms_value),
                ]
            )
            current_item = str(row.get("item_number", "")).strip()
            current_code = str(row.get("code", "")).strip()
            if (selected_item and current_item == selected_item) or (selected_code and current_code == selected_code):
                highlight_rows.add(index)

        total_quantity = Decimal("0")
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for row in invoice_details:
            total_quantity += self.decimal_value(row.get("quantity"))
            total_operation += self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            total_base += self.decimal_value(row.get("base_icms"))
            total_icms += self.decimal_value(row.get("icms_value"))

        note_header = {
            "document_number": selected_document,
            "document_key": selected_key,
            "document_date": str(invoice_details[0].get("document_date", "")) if invoice_details else "",
            "document_series": str(invoice_details[0].get("document_series", "")) if invoice_details else "",
            "document_model": str(invoice_details[0].get("document_model", "")) if invoice_details else "",
            "participant_name": str(invoice_details[0].get("participant_name", "")) if invoice_details else "",
            "participant_tax_id": str(invoice_details[0].get("participant_tax_id", "")) if invoice_details else "",
            "participant_code": str(invoice_details[0].get("participant_code", "")) if invoice_details else "",
            "effective_rate": compute_display_icms_rate(
                self.decimal_value(invoice_details[0].get("icms_rate")) if invoice_details else Decimal("0"),
                total_operation,
                total_base,
                total_icms,
            ),
            "items_count": len(rows),
        }
        self.open_invoice_mirror_popup_with_cards(
            title,
            headers,
            rows,
            total_quantity,
            total_operation,
            total_base,
            total_icms,
            note_header,
            highlight_rows,
            sorted_invoice_details,
            operation_type,
        )

    def open_invoice_mirror_popup_with_cards(
        self,
        title: str,
        headers: list[str],
        rows: list[list[object]],
        total_quantity: Decimal,
        total_operation: Decimal,
        total_base: Decimal,
        total_icms: Decimal,
        note_header: dict[str, object],
        highlight_rows: set[int] | None = None,
        runtime_rule_rows: list[dict[str, object]] | None = None,
        operation_type: str = "",
    ) -> None:
        if not rows:
            QMessageBox.warning(self, title, "Nao ha dados para esta nota.")
            return
        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(title)
        self.resize_dialog_to_screen(dialog, 1220, 700)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_label = QLabel(title)
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)
        section_label = QLabel("Cabecalho da Nota")
        section_label.setObjectName("sectionTitle")
        layout.addWidget(section_label)

        header_panel = self.create_panel()
        header_layout = QGridLayout(header_panel)
        header_layout.setContentsMargins(10, 10, 10, 10)
        header_layout.setHorizontalSpacing(18)
        header_layout.setVerticalSpacing(6)
        header_layout.addWidget(QLabel(f"Documento: {note_header.get('document_number', '')}"), 0, 0)
        header_layout.addWidget(QLabel(f"Fornecedor: {note_header.get('participant_name', '')}"), 0, 1)
        header_layout.addWidget(QLabel(f"Chave: {note_header.get('document_key', '')}"), 1, 0)
        header_layout.addWidget(QLabel(f"Cod. Participante: {note_header.get('participant_code', '')}"), 1, 1)
        header_layout.addWidget(QLabel(f"Data: {note_header.get('document_date', '')}"), 2, 0)
        header_layout.addWidget(QLabel(f"CPF/CNPJ: {note_header.get('participant_tax_id', '')}"), 2, 1)
        header_layout.addWidget(QLabel(f"Serie: {note_header.get('document_series', '')}"), 3, 0)
        header_layout.addWidget(QLabel(f"Aliq ICMS Efetiva: {self.format_number(note_header.get('effective_rate', Decimal('0')))}"), 3, 1)
        header_layout.addWidget(QLabel(f"Modelo: {note_header.get('document_model', '')}"), 4, 0)
        header_layout.addWidget(QLabel(f"Itens: {note_header.get('items_count', 0)}"), 4, 1)
        header_layout.setColumnStretch(0, 1)
        header_layout.setColumnStretch(1, 1)
        layout.addWidget(header_panel)

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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.UserRole, row_index)
                item.setForeground(QColor(COLORS["text"]))
                if column_index not in {2}:
                    item.setTextAlignment(Qt.AlignCenter)
                if highlight_rows and row_index in highlight_rows:
                    item.setBackground(QColor("#d9ecff"))
                table.setItem(row_index, column_index, item)
        self.apply_popup_column_widths(table, headers)
        self.enable_table_sorting(table)
        if runtime_rule_rows:
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(
                lambda pos: self.show_runtime_rule_context_menu_for_payloads(table, runtime_rule_rows, operation_type, pos, include_document=True)
            )
        layout.addWidget(table, 1)

        ratio = (total_base * Decimal("100") / total_operation).quantize(Decimal("0.01")) if total_operation > 0 else Decimal("0.00")
        cards_grid = QGridLayout()
        cards = (
            ("Quantidade", self.format_number(total_quantity)),
            ("Valor Operacao", self.format_number(total_operation)),
            ("Base ICMS", self.format_number(total_base)),
            ("% Base/Oper", self.format_number(ratio)),
            ("Valor ICMS", self.format_number(total_icms)),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            cards_grid.addWidget(card, 0, index)
        layout.addLayout(cards_grid)

        actions = QHBoxLayout()
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(title, headers, rows, "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(title, headers, rows, "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def open_documents_popup_for_details(self, details: list[dict[str, object]], caption: str) -> None:
        grouped: dict[tuple[str, str, str], dict[str, object]] = {}
        grouped_launch_details: dict[tuple[str, str, str], list[dict[str, object]]] = {}
        for detail in details:
            key = (str(detail.get("period", "")), str(detail.get("document_number", "")), str(detail.get("document_key", "")))
            bucket = grouped.setdefault(
                key,
                {"period": detail.get("period", ""), "document_date": detail.get("document_date", detail.get("date", "")), "document_number": detail.get("document_number", ""), "document_series": detail.get("document_series", ""), "document_model": detail.get("document_model", ""), "participant_name": detail.get("participant_name", ""), "participant_tax_id": detail.get("participant_tax_id", ""), "document_key": detail.get("document_key", ""), "items": 0, "sale_value": Decimal("0"), "base_icms": Decimal("0"), "icms_value": Decimal("0")},
            )
            grouped_launch_details.setdefault(key, []).append(detail)
            bucket["items"] = int(bucket["items"]) + 1
            bucket["sale_value"] = Decimal(bucket["sale_value"]) + self.decimal_value(self.first_row_value(detail, "sale_value", "total_operation_value", "operation_value"))
            bucket["base_icms"] = Decimal(bucket["base_icms"]) + self.decimal_value(detail.get("base_icms"))
            bucket["icms_value"] = Decimal(bucket["icms_value"]) + self.decimal_value(detail.get("icms_value"))
        grouped_items = sorted(grouped.items(), key=lambda item: (str(item[1]["period"]), str(item[1]["document_date"]), str(item[1]["document_number"])))
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
            for _key, row in grouped_items
        ]
        detail_payloads_by_row_index: dict[int, list[dict[str, object]]] = {
            index: [dict(detail) for detail in grouped_launch_details.get(key, []) if isinstance(detail, dict)]
            for index, (key, _row) in enumerate(grouped_items)
        }

        dialog = QDialog(self)
        dialog.setObjectName("popupDialog")
        dialog.setWindowTitle(f"Espelho de Documentos Fiscais - {caption}")
        self.resize_dialog_to_screen(dialog, 1380, 720)
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_label = QLabel(f"Espelho de Documentos Fiscais - {caption}")
        title_label.setObjectName("popupTitle")
        layout.addWidget(title_label)

        headers = ["Periodo", "Data", "Documento", "Serie", "Modelo", "Participante", "CPF/CNPJ", "Chave", "Itens", "Valor Operacao", "Base ICMS", "Valor ICMS"]
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
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setStretchLastSection(True)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor(COLORS["text"]))
                if column_index not in {5, 7}:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_index, column_index, item)
        self.apply_popup_column_widths(table, headers)
        self.enable_table_sorting(table)
        sort_state = {"column": -1, "order": Qt.AscendingOrder}
        table.horizontalHeader().sectionClicked.connect(
            lambda column: self.sort_table_by_column(table, column, sort_state)
        )
        operation_type = "Entrada" if "Entrada" in caption else "Saida"
        table.itemDoubleClicked.connect(
            lambda item: self.open_invoice_mirror_from_documents_popup_selection(item.row(), table, grouped_launch_details, operation_type)
        )
        layout.addWidget(table, 1)

        total_docs = len(rows)
        total_items = sum(int(row[8] or 0) for row in rows)
        total_operation = sum((self.decimal_value(row[9]) for row in rows), Decimal("0"))
        total_base = sum((self.decimal_value(row[10]) for row in rows), Decimal("0"))
        total_icms = sum((self.decimal_value(row[11]) for row in rows), Decimal("0"))
        ratio = (total_base * Decimal("100") / total_operation).quantize(Decimal("0.01")) if total_operation > 0 else Decimal("0.00")

        cards_grid = QGridLayout()
        cards_grid.setHorizontalSpacing(8)
        cards = (
            ("Documentos", str(total_docs)),
            ("Itens", str(total_items)),
            ("Valor Operacao", self.format_number(total_operation)),
            ("Base ICMS", self.format_number(total_base)),
            ("% Base/Oper", self.format_number(ratio)),
            ("Valor ICMS", self.format_number(total_icms)),
        )
        for index, (label, value) in enumerate(cards):
            card, value_label = self.create_metric_card_with_label(label, value)
            value_label.setText(value)
            cards_grid.addWidget(card, 0, index)
        layout.addLayout(cards_grid)

        actions = QHBoxLayout()
        actions.addWidget(
            self.create_button(
                "Abrir Espelho",
                lambda: self.open_invoice_mirror_from_documents_popup_selected_row(table, grouped_launch_details, operation_type),
                primary=True,
            )
        )
        actions.addWidget(self.create_button("Exportar Excel", lambda: self.export_popup_dataset(f"Espelho de Documentos Fiscais - {caption}", headers, [list(row) for row in rows], "xlsx")))
        actions.addWidget(self.create_button("Exportar CSV", lambda: self.export_popup_dataset(f"Espelho de Documentos Fiscais - {caption}", headers, [list(row) for row in rows], "csv")))
        actions.addStretch()
        actions.addWidget(self.create_button("Fechar", dialog.close))
        layout.addLayout(actions)
        dialog.exec()

    def sort_table_by_column(self, table: QTableWidget, column: int, sort_state: dict[str, object]) -> None:
        previous_column = int(sort_state.get("column", -1))
        previous_order = sort_state.get("order", Qt.AscendingOrder)
        if column == previous_column:
            order = Qt.DescendingOrder if previous_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            order = Qt.AscendingOrder
        table.sortItems(column, order)
        sort_state["column"] = column
        sort_state["order"] = order

    def open_invoice_mirror_from_documents_popup_selection(
        self,
        row_index: int,
        table: QTableWidget,
        grouped_launch_details: dict[tuple[str, str, str], list[dict[str, object]]],
        operation_type: str,
    ) -> None:
        if row_index < 0 or row_index >= table.rowCount():
            return
        document_item = table.item(row_index, 2)
        key_item = table.item(row_index, 7)
        selected_document = str(document_item.text()).strip() if document_item else ""
        selected_key = str(key_item.text()).strip() if key_item else ""
        selected_details: list[dict[str, object]] = []
        for key_tuple, detail_rows in grouped_launch_details.items():
            current_document = str(key_tuple[1]).strip() if len(key_tuple) > 1 else ""
            current_key = str(key_tuple[2]).strip() if len(key_tuple) > 2 else ""
            if (selected_key and current_key == selected_key) or (selected_document and current_document == selected_document):
                selected_details = [detail for detail in detail_rows if isinstance(detail, dict)]
                break
        if not selected_details:
            QMessageBox.warning(self, "Espelho da Nota", "Nao foi possivel localizar os itens desta nota para abrir o espelho.")
            return
        self.open_invoice_mirror_from_note_details(selected_details, operation_type)

    def open_invoice_mirror_from_documents_popup_selected_row(
        self,
        table: QTableWidget,
        grouped_launch_details: dict[tuple[str, str, str], list[dict[str, object]]],
        operation_type: str,
    ) -> None:
        selected_rows = table.selectionModel().selectedRows() if table.selectionModel() else []
        if not selected_rows:
            QMessageBox.warning(self, "Espelho da Nota", "Selecione uma linha para abrir o espelho da nota.")
            return
        self.open_invoice_mirror_from_documents_popup_selection(selected_rows[0].row(), table, grouped_launch_details, operation_type)

    def open_invoice_mirror_from_note_details(self, invoice_details: list[dict[str, object]], operation_type: str) -> None:
        if not invoice_details:
            return
        selected = invoice_details[0]
        selected_document = str(selected.get("document_number", ""))
        selected_key = str(selected.get("document_key", ""))
        selected_item = str(selected.get("item_number", "")).strip()
        selected_code = str(selected.get("code", "")).strip()
        title = f"Espelho da Nota - {selected_document or selected_key}"
        headers = ["Item", "Codigo", "Descricao", "NCM", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva"]
        rows: list[list[object]] = []
        highlight_rows: set[int] = set()
        sorted_invoice_details = sorted(
            invoice_details,
            key=lambda current: (
                str(current.get("item_number", "")),
                str(current.get("code", "")),
            ),
        )
        for index, row in enumerate(sorted_invoice_details):
            operation_value = self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(row.get("base_icms"))
            icms_value = self.decimal_value(row.get("icms_value"))
            rows.append(
                [
                    row.get("item_number", ""),
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("ncm", ""),
                    row.get("cest", ""),
                    row.get("cfop", ""),
                    self.first_row_value(row, "cst_icms", "cst"),
                    self.decimal_value(row.get("icms_rate")),
                    compute_display_icms_rate(self.decimal_value(row.get("icms_rate")), operation_value, base_icms, icms_value),
                ]
            )
            current_item = str(row.get("item_number", "")).strip()
            current_code = str(row.get("code", "")).strip()
            if (selected_item and current_item == selected_item) or (selected_code and current_code == selected_code):
                highlight_rows.add(index)

        total_quantity = Decimal("0")
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for row in invoice_details:
            total_quantity += self.decimal_value(row.get("quantity"))
            total_operation += self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            total_base += self.decimal_value(row.get("base_icms"))
            total_icms += self.decimal_value(row.get("icms_value"))

        note_header = {
            "document_number": selected_document,
            "document_key": selected_key,
            "document_date": str(invoice_details[0].get("document_date", "")),
            "document_series": str(invoice_details[0].get("document_series", "")),
            "document_model": str(invoice_details[0].get("document_model", "")),
            "participant_name": str(invoice_details[0].get("participant_name", "")),
            "participant_tax_id": str(invoice_details[0].get("participant_tax_id", "")),
            "participant_code": str(invoice_details[0].get("participant_code", "")),
            "effective_rate": compute_display_icms_rate(
                self.decimal_value(invoice_details[0].get("icms_rate")),
                total_operation,
                total_base,
                total_icms,
            ),
            "items_count": len(rows),
        }
        self.open_invoice_mirror_popup_with_cards(
            title,
            headers,
            rows,
            total_quantity,
            total_operation,
            total_base,
            total_icms,
            note_header,
            highlight_rows,
            sorted_invoice_details,
            operation_type,
        )




