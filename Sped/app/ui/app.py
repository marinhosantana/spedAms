from __future__ import annotations

# Transitional import while rules are being split into focused service modules.
from app.services.legacy_rules import *  # noqa: F403

class SpedApp:
    def __init__(self, root: Tk) -> None:
        # Interface principal para selecionar arquivos, filtros e gerar as saidas.
        self.root = root
        self.app_config_path = self.get_application_base_dir() / "app_config.json"
        app_config = self.load_app_config()
        self.app_window_title_var = StringVar(value=app_config["window_title"])
        self.app_home_title_var = StringVar(value=app_config["home_title"])
        self.app_config_status_var = StringVar(value="Informe os nomes que devem aparecer na tela inicial.")
        self.root.title(self.app_window_title_var.get().strip() or APP_DEFAULT_CONFIG["window_title"])
        self.root.minsize(860, 680)
        self.set_dialog_screen_geometry(self.root, 1440, 900, 860, 680, margin_y=150)
        self.root.state("zoomed")
        self.audit_log_path = self.get_audit_log_path()
        self.audit_session_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.audit_session_closed = False
        self.sped_path_var = StringVar()
        self.multi_sped_paths_var = StringVar()
        self.consult_sped_paths_var = StringVar()
        self.consult_xml_paths_var = StringVar()
        self.sales_consult_sped_paths_var = StringVar()
        self.sales_consult_xml_paths_var = StringVar()
        self.compare_sped_icms_path_var = StringVar()
        self.compare_sped_contrib_path_var = StringVar()
        self.compare_xml_folder_var = StringVar()
        self.compare_sheet_path_var = StringVar()
        self.compare_mode_var = StringVar(value="icms")
        self.compare_operation_scope_var = StringVar(value="Ambos")
        self.db_host_var = StringVar()
        self.db_port_var = StringVar(value="3306")
        self.db_user_var = StringVar()
        self.db_password_var = StringVar()
        self.db_name_var = StringVar(value="sped_icms")
        self.db_status_var = StringVar(value="Configure o MySQL, teste a conexao e crie o banco inicial.")
        self.company_status_var = StringVar(value="Ativa")
        self.company_search_var = StringVar()
        self.company_status_filter_var = StringVar(value="Todos")
        self.company_selector_var = StringVar()
        self.company_selector_all_values: list[str] = []
        self.company_selector_index: dict[int, str] = {}
        self.company_id_var = StringVar()
        self.company_razao_var = StringVar()
        self.company_fantasia_var = StringVar()
        self.company_cnpj_var = StringVar()
        self.company_ie_var = StringVar()
        self.product_status_var = StringVar(value="Ativo")
        self.product_search_var = StringVar()
        self.product_status_filter_var = StringVar(value="Todos")
        self.product_id_var = StringVar()
        self.product_company_var = StringVar(value="Empresa nao selecionada")
        self.product_codigo_var = StringVar()
        self.product_codigo_origem_var = StringVar()
        self.product_descricao_var = StringVar()
        self.product_ncm_var = StringVar()
        self.product_unidade_var = StringVar(value="UN")
        self.product_tipo_var = StringVar(value="Revenda")
        self.product_cst_icms_entrada_var = StringVar()
        self.product_cst_icms_saida_var = StringVar()
        self.product_cst_pis_entrada_var = StringVar()
        self.product_cst_pis_saida_var = StringVar()
        self.product_cst_cofins_entrada_var = StringVar()
        self.product_cst_cofins_saida_var = StringVar()
        self.product_icms_entrada_var = StringVar()
        self.product_icms_saida_var = StringVar()
        self.product_pis_entrada_var = StringVar()
        self.product_pis_saida_var = StringVar()
        self.product_cofins_entrada_var = StringVar()
        self.product_cofins_saida_var = StringVar()
        self.excel_path_var = StringVar()
        self.xml_path_var = StringVar()
        self.contrib_path_var = StringVar()
        self.nfce_ncm_filter_var = StringVar()
        self.icms_rate_var = StringVar()
        self.cst_filter_var = StringVar()
        self.cfop_filter_var = StringVar()
        self.contrib_consult_sped_paths_var = StringVar()
        self.contrib_sales_consult_sped_paths_var = StringVar()
        self.contrib_consult_xml_paths_var = StringVar()
        self.contrib_sales_consult_xml_paths_var = StringVar()
        self.new_cst_var = StringVar()
        self.new_cfop_var = StringVar()
        self.zero_icms_cfop_var = StringVar()
        self.adjustment_mode_var = StringVar(value="Filtros")
        self.adjusted_sped_scope_var = StringVar(value="Ambos")
        self.runtime_rule_filter_var = StringVar()
        self.consult_cst_filter_var = StringVar()
        self.consult_cfop_filter_var = StringVar()
        self.consult_status_filter_var = StringVar(value="Todos")
        self.consult_search_var = StringVar()
        self.sales_consult_cst_filter_var = StringVar()
        self.sales_consult_cfop_filter_var = StringVar()
        self.sales_consult_status_filter_var = StringVar(value="Todos")
        self.sales_consult_search_var = StringVar()
        self.contrib_consult_cst_filter_var = StringVar()
        self.contrib_consult_cfop_filter_var = StringVar()
        self.contrib_consult_status_filter_var = StringVar(value="Todos")
        self.contrib_consult_search_var = StringVar()
        self.contrib_sales_consult_cst_filter_var = StringVar()
        self.contrib_sales_consult_cfop_filter_var = StringVar()
        self.contrib_sales_consult_status_filter_var = StringVar(value="Todos")
        self.contrib_sales_consult_search_var = StringVar()
        self.consult_total_items_var = StringVar(value="Linhas: 0")
        self.consult_total_products_var = StringVar(value="Produtos: 0")
        self.consult_total_sale_var = StringVar(value="Valor Operacao: 0,00")
        self.consult_total_base_var = StringVar(value="Base ICMS: 0,00")
        self.consult_total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        self.consult_total_icms_var = StringVar(value="Valor ICMS: 0,00")
        self.sales_consult_total_items_var = StringVar(value="Linhas: 0")
        self.sales_consult_total_products_var = StringVar(value="Produtos: 0")
        self.sales_consult_total_sale_var = StringVar(value="Valor Operacao: 0,00")
        self.sales_consult_total_base_var = StringVar(value="Base ICMS: 0,00")
        self.sales_consult_total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        self.sales_consult_total_icms_var = StringVar(value="Valor ICMS: 0,00")
        self.contrib_consult_total_items_var = StringVar(value="Linhas: 0")
        self.contrib_consult_total_products_var = StringVar(value="Produtos: 0")
        self.contrib_consult_total_sale_var = StringVar(value="Valor Operacao: 0,00")
        self.contrib_consult_total_pis_base_var = StringVar(value="Base PIS: 0,00")
        self.contrib_consult_total_cofins_base_var = StringVar(value="Base COFINS: 0,00")
        self.contrib_consult_total_pis_var = StringVar(value="Valor PIS: 0,00")
        self.contrib_consult_total_cofins_var = StringVar(value="Valor COFINS: 0,00")
        self.contrib_sales_consult_total_items_var = StringVar(value="Linhas: 0")
        self.contrib_sales_consult_total_products_var = StringVar(value="Produtos: 0")
        self.contrib_sales_consult_total_sale_var = StringVar(value="Valor Operacao: 0,00")
        self.contrib_sales_consult_total_pis_base_var = StringVar(value="Base PIS: 0,00")
        self.contrib_sales_consult_total_cofins_base_var = StringVar(value="Base COFINS: 0,00")
        self.contrib_sales_consult_total_pis_var = StringVar(value="Valor PIS: 0,00")
        self.contrib_sales_consult_total_cofins_var = StringVar(value="Valor COFINS: 0,00")
        self.compare_stats_var = StringVar(value="Nenhuma comparacao executada.")
        self.compare_status_var = StringVar(value="Selecione os SPEDs e a pasta de XML ou uma planilha.")
        self.compare_progress_var = tk.DoubleVar(value=0)
        self.xml_sped_path_var = StringVar()
        self.xml_source_path_var = StringVar()
        self.xml_company_tax_id_var = StringVar(value="CNPJ/CPF empresa: nao carregado")
        self.xml_operation_scope_var = StringVar(value="Todos")
        self.xml_status_var = StringVar(value="Selecione uma pasta ou arquivo XML para processar.")
        self.xml_progress_var = tk.DoubleVar(value=0)
        self.xml_total_files_var = StringVar(value="XMLs: 0")
        self.xml_total_operation_files_var = StringVar(value="XMLs da operacao: 0")
        self.xml_total_cfops_var = StringVar(value="CFOPs: 0")
        self.xml_total_operation_var = StringVar(value="Total Operacao: 0,00")
        self.xml_total_base_icms_var = StringVar(value="Base ICMS: 0,00")
        self.xml_total_icms_var = StringVar(value="Valor ICMS: 0,00")
        self.xml_credit_status_var = StringVar(value="Processe os XMLs para listar os creditos destacados nas entradas.")
        self.xml_credit_total_notes_var = StringVar(value="Notas entrada: 0")
        self.xml_credit_total_items_var = StringVar(value="Itens: 0")
        self.xml_credit_total_operation_var = StringVar(value="Total Operacao: 0,00")
        self.xml_credit_total_base_icms_var = StringVar(value="Base ICMS: 0,00")
        self.xml_credit_total_icms_var = StringVar(value="Credito ICMS: 0,00")
        self.xml_credit_total_ipi_var = StringVar(value="Credito IPI: 0,00")
        self.entry_exit_sped_paths_var = StringVar(value="")
        self.entry_exit_status_var = StringVar(value="Selecione os SPEDs fiscais para montar a analise de entrada e saida.")
        self.entry_exit_total_saida_var = StringVar(value="Saidas ICMS: 0,00")
        self.entry_exit_total_entrada_var = StringVar(value="Entradas ICMS: 0,00")
        self.entry_exit_recolher_var = StringVar(value="A recolher: 0,00")
        self.entry_exit_rows: list[dict[str, object]] = []
        self.entry_exit_excel_rows: list[list[object]] = []
        self.entry_exit_totals: dict[str, Decimal] = {}
        self.status_var = StringVar(value="Selecione o SPED, a planilha, ou ambos para gerar o Excel.")
        self.runtime_rule_history_path = self.get_runtime_rule_history_path()
        self.runtime_rule_history: list[dict[str, object]] = []
        self.consult_period_labels: list[str] = []
        self.consult_comparison_rows: list[dict[str, object]] = []
        self.consult_filtered_rows: list[dict[str, object]] = []
        self.consult_summary_rows: list[dict[str, object]] = []
        self.sales_consult_period_labels: list[str] = []
        self.sales_consult_comparison_rows: list[dict[str, object]] = []
        self.sales_consult_filtered_rows: list[dict[str, object]] = []
        self.sales_consult_summary_rows: list[dict[str, object]] = []
        self.contrib_consult_period_labels: list[str] = []
        self.contrib_consult_comparison_rows: list[dict[str, object]] = []
        self.contrib_consult_filtered_rows: list[dict[str, object]] = []
        self.contrib_sales_consult_period_labels: list[str] = []
        self.contrib_sales_consult_comparison_rows: list[dict[str, object]] = []
        self.contrib_sales_consult_filtered_rows: list[dict[str, object]] = []
        self.xml_summary_all_rows: list[dict[str, object]] = []
        self.xml_summary_rows: list[dict[str, object]] = []
        self.xml_summary_stats: dict[str, object] = {}
        self.xml_ignored_log_rows: list[dict[str, object]] = []
        self.xml_document_log_rows: list[dict[str, object]] = []
        self.xml_cancelled_display_rows: list[dict[str, object]] = []
        self.xml_credit_invoice_rows: list[dict[str, object]] = []
        self.xml_credit_cfop_rows: list[dict[str, object]] = []
        self.xml_credit_stats: dict[str, object] = {}
        self.consult_period_path_map: dict[str, Path] = {}
        self.sales_consult_period_path_map: dict[str, Path] = {}
        self.contrib_consult_period_path_map: dict[str, Path] = {}
        self.contrib_sales_consult_period_path_map: dict[str, Path] = {}
        self.product_rows_by_id: dict[int, dict[str, object]] = {}
        self.compare_last_results: list[tuple[str, str, object]] = []
        self.compare_last_result_origin = ""
        self.compare_last_operation_scope = ""
        self.compare_last_mode = ""
        self.compare_last_stats: dict[str, object] = {}
        self.compare_diagnostic_status_var = StringVar(value="Execute uma comparacao em SPED x XML/Planilha e clique em Atualizar diagnostico.")
        self.compare_diagnostic_conclusion_var = StringVar(value="Nenhum diagnostico carregado.")
        self.compare_diagnostic_footer_var = StringVar(value="")
        self.runtime_rule_refresh_after_id: str | None = None
        self.consult_tree_sort_column = "code"
        self.consult_tree_sort_reverse = False
        self.consult_summary_sort_column = "code"
        self.consult_summary_sort_reverse = False
        self.sales_consult_tree_sort_column = "code"
        self.sales_consult_tree_sort_reverse = False
        self.sales_consult_summary_sort_column = "code"
        self.sales_consult_summary_sort_reverse = False
        self.contrib_consult_tree_sort_column = "code"
        self.contrib_consult_tree_sort_reverse = False
        self.contrib_sales_consult_tree_sort_column = "code"
        self.contrib_sales_consult_tree_sort_reverse = False
        self.consult_selected_product_code = ""
        self.sales_consult_selected_product_code = ""
        self.contrib_consult_selected_product_code = ""
        self.contrib_sales_consult_selected_product_code = ""
        self.selected_company_id: int | None = None
        self.selected_product_id: int | None = None
        self._consultation_filters_ready = False
        self._sales_consultation_filters_ready = False
        self._contrib_consultation_filters_ready = False
        self._contrib_sales_consultation_filters_ready = False
        self.mysql_repo = MysqlCadastroRepository(
            self.get_application_base_dir() / "mysql_config.json",
            self.get_application_base_dir() / "mysql_schema.sql",
        )
        self._configure_styles()
        self._build_layout()
        self.apply_app_config()
        self.load_mysql_config()
        self.setup_consultation_filter_traces()
        self.setup_sales_consultation_filter_traces()
        self.setup_contrib_consultation_filter_traces()
        self.setup_contrib_sales_consultation_filter_traces()
        self._consultation_filters_ready = True
        self._sales_consultation_filters_ready = True
        self._contrib_consultation_filters_ready = True
        self._contrib_sales_consultation_filters_ready = True
        self.load_runtime_rule_history()
        self.refresh_runtime_rule_history_list()
        self.schedule_company_tree_refresh()
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        self.write_audit_log("INICIO_SESSAO", f"Sistema iniciado. Sessao={self.audit_session_id}")

    def get_application_base_dir(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def get_audit_log_path(self) -> Path:
        log_dir = self.get_application_base_dir() / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        return log_dir / "auditoria_sped.log"

    def get_runtime_rule_history_path(self) -> Path:
        if getattr(sys, "frozen", False):
            local_appdata = Path(os.environ.get("LOCALAPPDATA", self.get_application_base_dir()))
            data_dir = local_appdata / "Revisor de SPED - DZ Consultoria"
            data_dir.mkdir(parents=True, exist_ok=True)
            return data_dir / "runtime_rule_history.json"
        return self.get_application_base_dir() / "runtime_rule_history.json"

    def schedule_company_tree_refresh(self) -> None:
        self.db_status_var.set("Carregando empresas...")
        threading.Thread(target=self._load_company_tree_data_background, daemon=True).start()

    def _load_company_tree_data_background(self) -> None:
        try:
            companies = self.mysql_repo.list_companies() if self.mysql_repo.mysql_available() else []
            self.root.after(0, lambda companies=companies: self._apply_company_tree_refresh(companies, ""))
        except Exception as exc:
            self.root.after(0, lambda error_message=str(exc): self._apply_company_tree_refresh([], error_message))

    def _apply_company_tree_refresh(self, companies: list[dict[str, object]], error_message: str) -> None:
        if error_message:
            self.db_status_var.set(f"MySQL indisponivel para listar empresas: {error_message}")
            self.log_message(f"Falha ao carregar empresas na inicializacao: {error_message}")
            return
        self.refresh_company_tree(companies)

    def _configure_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        bg_app = "#eaf6ff"
        bg_card = "#f7fbff"
        border = "#b8d9ef"
        text = "#184c73"
        accent = "#7cc6f2"
        accent_hover = "#68b9e8"
        accent_dark = "#4a9ed1"

        self.root.configure(bg=bg_app)

        style.configure("TFrame", background=bg_app)
        style.configure("Card.TFrame", background=bg_card)
        style.configure("TLabel", background=bg_app, foreground=text)
        style.configure("Title.TLabel", background=bg_app, foreground=accent_dark)
        style.configure("Strong.TLabel", background=bg_app, foreground="#000000", font=("Segoe UI", 10, "bold"))
        style.configure(
            "TLabelframe",
            background=bg_card,
            bordercolor=border,
            relief="solid",
            borderwidth=1,
        )
        style.configure("TLabelframe.Label", background=bg_app, foreground=accent_dark, font=("Segoe UI", 10, "bold"))
        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            foreground=text,
            bordercolor=border,
            lightcolor=border,
            darkcolor=border,
            padding=5,
        )
        style.configure(
            "Primary.TButton",
            background=accent,
            foreground="#08324d",
            bordercolor=accent,
            lightcolor=accent,
            darkcolor=accent,
            relief="flat",
            padding=(10, 6),
            focusthickness=0,
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", accent_hover), ("pressed", accent_dark)],
            foreground=[("active", "#08324d")],
        )
        style.configure(
            "Secondary.TButton",
            background="#d9eefc",
            foreground=text,
            bordercolor=border,
            lightcolor=border,
            darkcolor=border,
            relief="flat",
            padding=(9, 6),
            focusthickness=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#c9e6fa"), ("pressed", "#b4daf5")],
            foreground=[("active", text)],
        )
        style.configure("TNotebook", background=bg_app, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background="#d9eefc",
            foreground=text,
            padding=(10, 7),
            font=("Segoe UI", 8, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#ffffff"), ("active", "#c9e6fa")],
            foreground=[("selected", accent_dark)],
        )
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground=text,
            rowheight=24,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure("Treeview.Heading", background="#d8eefc", foreground=text, relief="flat", font=("Segoe UI", 9, "bold"))

    def _build_layout(self) -> None:
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(
            frame,
            textvariable=self.app_home_title_var,
            font=("Segoe UI", 15, "bold"),
            foreground="#08324d",
            justify="center",
        ).pack(pady=(0, 6))

        notebook = ttk.Notebook(frame)
        notebook.pack(fill=BOTH, expand=True)

        fiscal_root_tab = ttk.Frame(notebook, padding=6, style="Card.TFrame")
        contrib_root_tab = ttk.Frame(notebook, padding=6, style="Card.TFrame")
        xml_root_tab = ttk.Frame(notebook, padding=6, style="Card.TFrame")
        entry_exit_root_tab = ttk.Frame(notebook, padding=6, style="Card.TFrame")
        settings_root_tab = ttk.Frame(notebook, padding=6, style="Card.TFrame")
        notebook.add(fiscal_root_tab, text="Fiscal")
        notebook.add(contrib_root_tab, text="Contribuicoes")
        notebook.add(xml_root_tab, text="XML")
        notebook.add(entry_exit_root_tab, text="Analise Entrada e Saida")
        notebook.add(settings_root_tab, text="Configuracoes")

        fiscal_root_tab.columnconfigure(0, weight=1)
        fiscal_root_tab.rowconfigure(0, weight=1)
        contrib_root_tab.columnconfigure(0, weight=1)
        contrib_root_tab.rowconfigure(0, weight=1)
        entry_exit_root_tab.columnconfigure(0, weight=1)
        entry_exit_root_tab.rowconfigure(1, weight=1)
        settings_root_tab.columnconfigure(0, weight=1)
        settings_root_tab.rowconfigure(0, weight=1)

        fiscal_notebook = ttk.Notebook(fiscal_root_tab)
        fiscal_notebook.grid(row=0, column=0, sticky="nsew")
        contrib_notebook = ttk.Notebook(contrib_root_tab)
        contrib_notebook.grid(row=0, column=0, sticky="nsew")

        main_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        filters_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        runtime_rules_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        consult_process_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        consult_view_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        sales_consult_view_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        compare_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        diagnostic_tab = ttk.Frame(fiscal_notebook, padding=6, style="Card.TFrame")
        fiscal_notebook.add(consult_process_tab, text="Processar Consultas")
        fiscal_notebook.add(consult_view_tab, text="Consulta Entradas ICMS")
        fiscal_notebook.add(sales_consult_view_tab, text="Consulta Saidas ICMS")
        fiscal_notebook.add(compare_tab, text="SPED x XML/Planilha")
        fiscal_notebook.add(diagnostic_tab, text="Diagnostico")
        fiscal_notebook.add(runtime_rules_tab, text="Regras Dinamicas")
        settings_notebook = ttk.Notebook(settings_root_tab)
        settings_notebook.grid(row=0, column=0, sticky="nsew")
        app_config_tab = ttk.Frame(settings_notebook, padding=6, style="Card.TFrame")
        mysql_connection_tab = ttk.Frame(settings_notebook, padding=6, style="Card.TFrame")
        cadastro_tab = ttk.Frame(settings_notebook, padding=6, style="Card.TFrame")
        settings_notebook.add(app_config_tab, text="Aplicacao")
        settings_notebook.add(mysql_connection_tab, text="Conexao MySQL")
        settings_notebook.add(cadastro_tab, text="Cadastro")

        contrib_process_tab = ttk.Frame(contrib_notebook, padding=6, style="Card.TFrame")
        contrib_consult_view_tab = ttk.Frame(contrib_notebook, padding=6, style="Card.TFrame")
        contrib_sales_consult_view_tab = ttk.Frame(contrib_notebook, padding=6, style="Card.TFrame")
        contrib_notebook.add(contrib_process_tab, text="Processar Consultas")
        contrib_notebook.add(contrib_consult_view_tab, text="Consulta Entradas PIS/COFINS")
        contrib_notebook.add(contrib_sales_consult_view_tab, text="Consulta Saidas PIS/COFINS")

        main_tab.columnconfigure(0, weight=1)
        main_tab.columnconfigure(1, weight=1)
        filters_tab.columnconfigure(0, weight=3)
        filters_tab.columnconfigure(1, weight=2)
        runtime_rules_tab.columnconfigure(0, weight=1)
        runtime_rules_tab.columnconfigure(1, weight=1)
        runtime_rules_tab.rowconfigure(0, weight=1)
        consult_process_tab.columnconfigure(0, weight=1)
        consult_process_tab.rowconfigure(0, weight=1)
        contrib_process_tab.columnconfigure(0, weight=1)
        contrib_process_tab.rowconfigure(0, weight=1)
        consult_view_tab.columnconfigure(0, weight=1)
        consult_view_tab.rowconfigure(1, weight=1)
        sales_consult_view_tab.columnconfigure(0, weight=1)
        sales_consult_view_tab.rowconfigure(1, weight=1)
        contrib_consult_view_tab.columnconfigure(0, weight=1)
        contrib_consult_view_tab.rowconfigure(1, weight=1)
        contrib_sales_consult_view_tab.columnconfigure(0, weight=1)
        contrib_sales_consult_view_tab.rowconfigure(1, weight=1)
        compare_tab.columnconfigure(0, weight=1)
        compare_tab.rowconfigure(1, weight=1)
        diagnostic_tab.columnconfigure(0, weight=1)
        diagnostic_tab.rowconfigure(2, weight=1)
        mysql_connection_tab.columnconfigure(0, weight=1)
        app_config_tab.columnconfigure(0, weight=1)
        cadastro_tab.columnconfigure(0, weight=1)
        cadastro_tab.rowconfigure(0, weight=1)

        source_box = ttk.LabelFrame(main_tab, text="Arquivos de Entrada", padding=10)
        source_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))

        sped_frame = ttk.Frame(source_box)
        sped_frame.pack(fill="x")
        ttk.Label(sped_frame, text="Arquivo SPED (opcional):").pack(anchor="w")
        sped_row = ttk.Frame(sped_frame)
        sped_row.pack(fill="x", pady=(4, 8))
        ttk.Entry(sped_row, textvariable=self.sped_path_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            sped_row,
            text="Selecionar SPED",
            style="Secondary.TButton",
            command=lambda: self.select_file(self.sped_path_var, "Selecione o arquivo SPED Fiscal", [("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")]),
        ).pack(side=RIGHT, padx=(8, 0))

        excel_frame = ttk.Frame(source_box)
        excel_frame.pack(fill="x")
        ttk.Label(excel_frame, text="Planilha de produtos (opcional):").pack(anchor="w")
        excel_row = ttk.Frame(excel_frame)
        excel_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(excel_row, textvariable=self.excel_path_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            excel_row,
            text="Selecionar Planilha",
            style="Secondary.TButton",
            command=lambda: self.select_file(self.excel_path_var, "Selecione a planilha Excel", [("Arquivo Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]),
        ).pack(side=RIGHT, padx=(8, 0))

        xml_frame = ttk.Frame(source_box)
        xml_frame.pack(fill="x", pady=(8, 0))
        ttk.Label(xml_frame, text="XML NF-e/NFC-e (opcional):").pack(anchor="w")
        xml_row = ttk.Frame(xml_frame)
        xml_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(xml_row, textvariable=self.xml_path_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            xml_row,
            text="Selecionar XMLs/Pasta",
            style="Secondary.TButton",
            command=lambda: self.select_files_or_directory(self.xml_path_var, "Selecione um ou mais XMLs da NF-e/NFC-e ou uma pasta de XMLs", [("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]),
        ).pack(side=RIGHT, padx=(8, 0))
        ttk.Label(xml_frame, text="Filtro NCM para exportacao NFC-e:").pack(anchor="w", pady=(10, 0))
        ttk.Entry(xml_frame, textvariable=self.nfce_ncm_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(
            xml_frame,
            text="Informe um ou mais NCMs separados por virgula, ponto e virgula ou linha. Use * para prefixo, ex.: 2203*.",
            wraplength=760,
        ).pack(anchor="w", pady=(4, 0))

        contrib_frame = ttk.Frame(source_box)
        contrib_frame.pack(fill="x", pady=(8, 0))
        ttk.Label(contrib_frame, text="Arquivo SPED Contribuicoes (opcional):").pack(anchor="w")
        contrib_row = ttk.Frame(contrib_frame)
        contrib_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(contrib_row, textvariable=self.contrib_path_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            contrib_row,
            text="Selecionar Contribuicoes",
            style="Secondary.TButton",
            command=lambda: self.select_file(self.contrib_path_var, "Selecione o arquivo SPED Contribuicoes", [("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")]),
        ).pack(side=RIGHT, padx=(8, 0))
        source_actions = ttk.Frame(source_box)
        source_actions.pack(fill="x", pady=(12, 0))
        ttk.Button(
            source_actions,
            text="Limpar Tela",
            style="Secondary.TButton",
            command=self.clear_form,
        ).pack(side=RIGHT)

        settings_box = ttk.LabelFrame(main_tab, text="Filtros e Recalculo", padding=10)
        settings_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))

        mode_frame = ttk.Frame(settings_box)
        mode_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(mode_frame, text="Modo de ajuste:").pack(anchor="w")
        ttk.Combobox(
            mode_frame,
            textvariable=self.adjustment_mode_var,
            values=(
                "Filtros",
                "Regra Padrao Filial",
                "Regra Padrao Matriz",
                "Regra Nova Matriz",
                "Nascimento e Curtarelli",
                "BELLA CITTA 0001-02 - ENTRADAS-Saidas",
                "NASCIMENTO 0001-88 - ENTRADAS-Saidas",
                "CASA DE PAES - ENTRADAS",
                "NIG",
            ),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Label(
            mode_frame,
            text="Use Filtros para ajustes manuais ou selecione uma regra padrao para aplicar automaticamente as combinacoes CST/CFOP da filial ou matriz.",
            wraplength=760,
        ).pack(anchor="w")

        recalc_grid = ttk.Frame(settings_box)
        recalc_grid.pack(fill="x")
        recalc_grid.columnconfigure(0, weight=1)
        recalc_grid.columnconfigure(1, weight=1)

        rate_frame = ttk.Frame(recalc_grid)
        rate_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 8))
        ttk.Label(rate_frame, text="Aliquota ICMS para recalculo (%):").pack(anchor="w")
        ttk.Entry(rate_frame, textvariable=self.icms_rate_var).pack(fill="x", pady=(4, 0))
        ttk.Label(
            rate_frame,
            text="Ex.: 4, 7 ou 12. Vazio mantém a aliquota atual.",
            wraplength=360,
        ).pack(anchor="w", pady=(4, 0))

        replacement_left = ttk.Frame(recalc_grid)
        replacement_left.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 8))
        ttk.Label(replacement_left, text="Novo CST para itens filtrados:").pack(anchor="w")
        ttk.Entry(replacement_left, textvariable=self.new_cst_var).pack(fill="x", pady=(4, 6))
        ttk.Label(
            replacement_left,
            text="Ex.: 000, 020 ou 040. Vazio nao troca o CST.",
            wraplength=360,
        ).pack(anchor="w")
        ttk.Label(replacement_left, text="Novo CFOP para itens filtrados:").pack(anchor="w", pady=(10, 0))
        ttk.Entry(replacement_left, textvariable=self.new_cfop_var).pack(fill="x", pady=(4, 6))
        ttk.Label(
            replacement_left,
            text="Ex.: 5102, 5405 ou 6108. Vazio nao troca o CFOP.",
            wraplength=360,
        ).pack(anchor="w")

        replacement_bottom = ttk.Frame(settings_box)
        replacement_bottom.pack(fill="x", pady=(4, 0))
        ttk.Label(replacement_bottom, text="CFOPs para zerar ICMS/Base/Valor:").pack(anchor="w")
        ttk.Entry(replacement_bottom, textvariable=self.zero_icms_cfop_var).pack(fill="x", pady=(4, 6))
        ttk.Label(
            replacement_bottom,
            text="Informe um ou mais CFOPs separados por virgula, ponto e virgula ou quebra de linha. Ex.: 5405, 5403.",
            wraplength=760,
        ).pack(anchor="w")

        sped_scope_frame = ttk.Frame(settings_box)
        sped_scope_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(sped_scope_frame, text="Gerar SPED Ajustado para:").pack(anchor="w")
        ttk.Combobox(
            sped_scope_frame,
            textvariable=self.adjusted_sped_scope_var,
            values=("Ambos", "Saida", "Entrada"),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Label(
            sped_scope_frame,
            text="Escolha se o ajuste do arquivo .txt deve considerar entradas, saidas ou ambos.",
            wraplength=760,
        ).pack(anchor="w")

        settings_box.grid_remove()
        source_box.grid_configure(columnspan=2, padx=(0, 0))

        filter_box = ttk.LabelFrame(filters_tab, text="Filtro por Descricao", padding=10)
        filter_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))

        filter_header = ttk.Frame(filter_box)
        filter_header.pack(fill="x")
        ttk.Label(filter_header, text="Descricoes para filtrar (uma por linha):").pack(side=LEFT, anchor="w")
        ttk.Button(filter_header, text="Carregar Filtros", style="Secondary.TButton", command=self.load_filter_descriptions).pack(side=RIGHT)
        self.filter_text = Text(
            filter_box,
            height=5,
            bg="#ffffff",
            fg="#184c73",
            insertbackground="#184c73",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#b8d9ef",
            highlightcolor="#7cc6f2",
        )
        self.filter_text.pack(fill="x", pady=(4, 12))

        extra_filters_box = ttk.LabelFrame(filters_tab, text="Filtros de CST e CFOP", padding=10)
        extra_filters_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))
        ttk.Label(extra_filters_box, text="Filtro CST (separe por virgula, ponto e virgula ou linha):").pack(anchor="w")
        ttk.Entry(extra_filters_box, textvariable=self.cst_filter_var).pack(fill="x", pady=(4, 8))
        ttk.Label(extra_filters_box, text="Ex.: 000, 020, 060. Vazio nao filtra CST.").pack(anchor="w")
        ttk.Label(extra_filters_box, text="Filtro CFOP (separe por virgula, ponto e virgula ou linha):").pack(anchor="w")
        ttk.Entry(extra_filters_box, textvariable=self.cfop_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(extra_filters_box, text="Ex.: 5102, 5405, 6108. Vazio nao filtra CFOP.").pack(anchor="w", pady=(4, 0))

        runtime_box = ttk.LabelFrame(runtime_rules_tab, text="Regras informadas em tempo de execucao", padding=10)
        runtime_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(
            runtime_box,
            text=(
                "Informe uma regra por linha no formato chave=valor; chave=valor. "
                "Essas regras sao aplicadas antes do perfil fixo e nao alteram IPI."
            ),
            wraplength=760,
        ).pack(anchor="w")
        ttk.Label(
            runtime_box,
            text=(
                "Filtros: tipo, documento, cnpj, cst, cfop, aliquota, codigo, codigos.\n"
                "Acoes: novo_cst, novo_cfop, nova_aliquota, zerar_icms, usar_valor_operacao_como_base,\n"
                "definir_base_icms, definir_valor_icms, percentual_reducao_base, recalcular_base_reducao,\n"
                "recalcular_valor_icms."
            ),
            wraplength=760,
        ).pack(anchor="w", pady=(6, 8))
        ttk.Label(
            runtime_box,
            text=(
                "Exemplos:\n"
                "tipo=entrada; cst=000; cfop=1102; aliquota=18; novo_cfop=1101\n"
                "tipo=entrada; cst=000; cfop=1250; aliquota=0; nova_aliquota=18; usar_valor_operacao_como_base=sim\n"
                "tipo=entrada; cst=000; cfop=1556; aliquota=18; novo_cst=090; zerar_icms=sim\n"
                "tipo=entrada; cst=000; cfop=1102; aliquota=12; percentual_reducao_base=20; recalcular_base_reducao=sim"
            ),
            wraplength=760,
        ).pack(anchor="w", pady=(0, 8))
        runtime_actions = ttk.Frame(runtime_box)
        runtime_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            runtime_actions,
            text="Adicionar Regra Assistida",
            style="Secondary.TButton",
            command=self.open_runtime_rule_builder,
        ).pack(side=LEFT)
        ttk.Button(
            runtime_actions,
            text="Limpar Regras Dinamicas",
            style="Secondary.TButton",
            command=lambda: self.runtime_rules_text.delete("1.0", END),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            runtime_actions,
            text="Exportar Regras Word",
            style="Secondary.TButton",
            command=self.open_rule_export_dialog,
        ).pack(side=LEFT, padx=(8, 0))
        self.runtime_rules_text = Text(
            runtime_box,
            height=10,
            bg="#ffffff",
            fg="#184c73",
            insertbackground="#184c73",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#b8d9ef",
            highlightcolor="#7cc6f2",
        )
        self.runtime_rules_text.pack(fill="both", expand=True)

        history_box = ttk.LabelFrame(runtime_rules_tab, text="Historico de Regras Utilizadas", padding=10)
        history_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        history_filter_row = ttk.Frame(history_box)
        history_filter_row.pack(fill="x", pady=(0, 8))
        ttk.Label(history_filter_row, text="Filtrar por qualquer parte da regra:").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(history_filter_row, textvariable=self.runtime_rule_filter_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            history_filter_row,
            text="Filtrar",
            style="Secondary.TButton",
            command=self.refresh_runtime_rule_history_list,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            history_filter_row,
            text="Limpar Filtro",
            style="Secondary.TButton",
            command=self.clear_runtime_rule_filter,
        ).pack(side=LEFT, padx=(8, 0))
        self.runtime_rule_history_tree = ttk.Treeview(
            history_box,
            columns=("regra",),
            show="headings",
            height=8,
        )
        self.runtime_rule_history_tree.heading("regra", text="Regra")
        self.runtime_rule_history_tree.column("regra", width=560, anchor="center")
        self.runtime_rule_history_tree.pack(fill=BOTH, expand=True)
        self.runtime_rule_history_tree.bind("<<TreeviewSelect>>", lambda _event: self.update_runtime_rule_history_preview())
        self.runtime_rule_history_tree.bind("<Double-1>", lambda _event: self.insert_selected_runtime_rule())
        preview_label = ttk.Label(history_box, text="Regra completa selecionada:")
        preview_label.pack(anchor="w", pady=(8, 4))
        preview_frame = ttk.Frame(history_box)
        preview_frame.pack(fill="both", expand=True)
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical")
        self.runtime_rule_history_preview = Text(
            preview_frame,
            height=4,
            bg="#ffffff",
            fg="#184c73",
            insertbackground="#184c73",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#b8d9ef",
            highlightcolor="#7cc6f2",
            wrap="word",
            yscrollcommand=preview_scrollbar.set,
        )
        preview_scrollbar.config(command=self.runtime_rule_history_preview.yview)
        self.runtime_rule_history_preview.pack(side=LEFT, fill=BOTH, expand=True)
        preview_scrollbar.pack(side=RIGHT, fill="y")
        history_actions = ttk.Frame(history_box)
        history_actions.pack(fill="x", pady=(8, 0))
        ttk.Button(
            history_actions,
            text="Inserir Selecionada",
            style="Secondary.TButton",
            command=self.insert_selected_runtime_rule,
        ).pack(side=LEFT)
        ttk.Button(
            history_actions,
            text="Excluir do Historico",
            style="Secondary.TButton",
            command=self.delete_selected_runtime_rule,
        ).pack(side=LEFT, padx=(8, 0))

        consult_process_canvas = Canvas(consult_process_tab, background="#f7fbff", highlightthickness=0)
        consult_process_scrollbar = ttk.Scrollbar(consult_process_tab, orient="vertical", command=consult_process_canvas.yview)
        consult_process_scrollable = ttk.Frame(consult_process_canvas, style="Card.TFrame", padding=(0, 0, 10, 0))
        consult_process_scrollable.columnconfigure(0, weight=1)
        consult_process_canvas.configure(yscrollcommand=consult_process_scrollbar.set)
        consult_process_canvas.grid(row=0, column=0, sticky="nsew")
        consult_process_scrollbar.grid(row=0, column=1, sticky="ns")
        consult_process_window = consult_process_canvas.create_window((0, 0), window=consult_process_scrollable, anchor="nw")

        def update_consult_process_scroll_region(_event: object) -> None:
            consult_process_canvas.configure(scrollregion=consult_process_canvas.bbox("all"))

        def resize_consult_process_content(event: object) -> None:
            consult_process_canvas.itemconfigure(consult_process_window, width=getattr(event, "width", 0))

        consult_process_scrollable.bind("<Configure>", update_consult_process_scroll_region)
        consult_process_canvas.bind("<Configure>", resize_consult_process_content)

        combined_consult_actions_box = ttk.LabelFrame(consult_process_scrollable, text="Acoes de Processamento", padding=10)
        combined_consult_actions_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(
            combined_consult_actions_box,
            text="Use os botoes abaixo para processar somente entradas, somente saidas ou atualizar as duas consultas em sequencia.",
            wraplength=820,
        ).pack(anchor="w")
        combined_consult_actions_row = ttk.Frame(combined_consult_actions_box)
        combined_consult_actions_row.pack(fill="x", pady=(10, 0))
        ttk.Button(
            combined_consult_actions_row,
            text="Processar Entradas",
            style="Primary.TButton",
            command=self.process_consultation_speds,
        ).pack(side=LEFT)
        ttk.Button(
            combined_consult_actions_row,
            text="Processar Saidas",
            style="Primary.TButton",
            command=self.process_sales_consultation_speds,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            combined_consult_actions_row,
            text="Processar Ambos",
            style="Secondary.TButton",
            command=self.process_all_consultations,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            combined_consult_actions_row,
            text="Reprocessar SPED",
            style="Secondary.TButton",
            command=self.generate_adjusted_sped,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            combined_consult_actions_row,
            text="Limpar Tela",
            style="Secondary.TButton",
            command=self.clear_form,
        ).pack(side=LEFT, padx=(8, 0))

        consult_process_box = ttk.LabelFrame(consult_process_scrollable, text="Consulta Entradas", padding=10)
        consult_process_box.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        ttk.Label(
            consult_process_box,
            text=(
                "Carregue de 1 a 12 SPEDs para consulta em tela. Depois use a aba Consulta Entradas "
                "para comparar os periodos por CST, CFOP e produto."
            ),
            wraplength=820,
        ).pack(anchor="w")
        consult_process_row = ttk.Frame(consult_process_box)
        consult_process_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(consult_process_row, textvariable=self.consult_sped_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            consult_process_row,
            text="Selecionar SPEDs",
            style="Secondary.TButton",
            command=self.select_consult_sped_files,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_process_row,
            text="Processar",
            style="Primary.TButton",
            command=self.process_consultation_speds,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_process_row,
            text="Limpar Lista",
            style="Secondary.TButton",
            command=lambda: self.consult_sped_paths_var.set(""),
        ).pack(side=LEFT, padx=(8, 0))
        consult_xml_row = ttk.Frame(consult_process_box)
        consult_xml_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(consult_xml_row, textvariable=self.consult_xml_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            consult_xml_row,
            text="Importar XMLs 55/65",
            style="Secondary.TButton",
            command=self.select_fiscal_consult_xml_sources,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_xml_row,
            text="Limpar XMLs",
            style="Secondary.TButton",
            command=lambda: [self.consult_xml_paths_var.set(""), self.sales_consult_xml_paths_var.set("")],
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            consult_process_box,
            text="Use o botao Processar para carregar os SPEDs na aba Consulta Entradas.",
            style="Title.TLabel",
        ).pack(anchor="w", pady=(10, 0))
        ttk.Label(
            consult_process_box,
            text=(
                "Quando houver documento 55/65 sem C170 no SPED, os itens serao buscados nos XMLs importados "
                "pela chave informada no C100."
            ),
            wraplength=820,
        ).pack(anchor="w", pady=(8, 0))

        consult_filter_box = ttk.LabelFrame(consult_view_tab, text="Filtros de Consulta", padding=10)
        consult_filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        consult_filter_box.columnconfigure(0, weight=2)
        consult_filter_box.columnconfigure(1, weight=1)
        consult_filter_box.columnconfigure(2, weight=1)
        consult_filter_box.columnconfigure(3, weight=1)

        periods_box = ttk.Frame(consult_filter_box)
        periods_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ttk.Label(periods_box, text="Periodos para comparar:").pack(anchor="w")
        self.consult_periods_checks_box = ttk.Frame(periods_box)
        self.consult_periods_checks_box.pack(fill="x", pady=(4, 0))
        self.consult_period_check_vars: dict[str, BooleanVar] = {}

        cst_box = ttk.Frame(consult_filter_box)
        cst_box.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(cst_box, text="Filtro CST:").pack(anchor="w")
        ttk.Entry(cst_box, textvariable=self.consult_cst_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(cst_box, text="Ex.: 000, 020, 060").pack(anchor="w", pady=(4, 0))

        cfop_box = ttk.Frame(consult_filter_box)
        cfop_box.grid(row=0, column=2, sticky="ew", padx=(0, 10))
        ttk.Label(cfop_box, text="Filtro CFOP:").pack(anchor="w")
        ttk.Entry(cfop_box, textvariable=self.consult_cfop_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(cfop_box, text="Ex.: 1101, 1401, 1556").pack(anchor="w", pady=(4, 0))

        extra_filter_box = ttk.Frame(consult_filter_box)
        extra_filter_box.grid(row=0, column=3, sticky="ew")
        ttk.Label(extra_filter_box, text="Status / Busca:").pack(anchor="w")
        ttk.Combobox(
            extra_filter_box,
            textvariable=self.consult_status_filter_var,
            values=("Todos", "Ok", "Sem credito", "Sem entrada", "Com divergencia"),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Entry(extra_filter_box, textvariable=self.consult_search_var).pack(fill="x")
        ttk.Label(extra_filter_box, text="Codigo ou descricao").pack(anchor="w", pady=(4, 0))

        consult_filter_actions = ttk.Frame(consult_filter_box)
        consult_filter_actions.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        consult_filter_actions_top = ttk.Frame(consult_filter_actions)
        consult_filter_actions_top.pack(fill="x", anchor="w")
        ttk.Button(
            consult_filter_actions_top,
            text="Aplicar Filtros",
            style="Primary.TButton",
            command=self.refresh_consultation_tree,
        ).pack(side=LEFT)
        ttk.Button(
            consult_filter_actions_top,
            text="Limpar Filtros",
            style="Secondary.TButton",
            command=self.clear_consultation_filters,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Exportar Filtro Atual",
            style="Secondary.TButton",
            command=self.export_current_consultation_filter,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Entradas",
            style="Secondary.TButton",
            command=lambda: self.open_operation_summary_popup("Entrada"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Comp. Diag. Credito",
            style="Secondary.TButton",
            command=self.open_credit_diagnostic_period_comparison_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Diag. Credito",
            style="Secondary.TButton",
            command=self.open_credit_diagnostic_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Curva ABC",
            style="Secondary.TButton",
            command=lambda: self.open_product_monthly_summary_popup("Entrada", "consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Reducao BC",
            style="Secondary.TButton",
            command=self.open_reduction_items_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Apuracao",
            style="Secondary.TButton",
            command=self.open_apuracao_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            consult_filter_actions_top,
            text="Espelho Docs",
            style="Secondary.TButton",
            command=lambda: self.open_fiscal_documents_popup("consult", "Entrada"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            consult_filter_actions,
            text="Dica: clique em Docs ou Lancamentos para abrir o detalhamento.",
        ).pack(anchor="w", pady=(6, 0))

        consult_results_box = ttk.Frame(consult_view_tab)
        consult_results_box.grid(row=1, column=0, sticky="nsew")
        consult_results_box.columnconfigure(0, weight=1)
        consult_results_box.rowconfigure(0, weight=1)

        consult_grid_box = ttk.LabelFrame(consult_results_box, text="Comparacao por Produto e Periodo", padding=10)
        consult_grid_box.grid(row=0, column=0, sticky="nsew")
        consult_grid_box.rowconfigure(0, weight=1)
        consult_grid_box.columnconfigure(0, weight=1)
        consult_grid_box.columnconfigure(1, weight=0)
        consult_columns = (
            "period",
            "curve_abc",
            "code",
            "description",
            "supplier",
            "ncm",
            "cest",
            "cfop",
            "cst",
            "icms_rate",
            "effective_icms_rate",
            "documents",
            "launches",
            "sale_value",
            "base_icms",
            "icms_value",
            "status",
        )
        self.consult_tree = ttk.Treeview(
            consult_grid_box,
            columns=consult_columns,
            show="headings",
            height=14,
            selectmode="extended",
        )
        consult_headings = {
            "period": "Periodo",
            "curve_abc": "Curva ABC",
            "code": "Codigo",
            "description": "Descricao",
            "supplier": "Fornecedor",
            "ncm": "NCM",
            "cest": "CEST",
            "cfop": "CFOP",
            "cst": "CST",
            "icms_rate": "Aliq ICMS",
            "effective_icms_rate": "Aliq Efetiva",
            "documents": "Docs",
            "launches": "Lancamentos",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
            "status": "Status",
        }
        consult_widths = {
            "period": 90,
            "curve_abc": 80,
            "code": 100,
            "description": 260,
            "supplier": 220,
            "ncm": 100,
            "cest": 90,
            "cfop": 90,
            "cst": 90,
            "icms_rate": 90,
            "effective_icms_rate": 95,
            "documents": 60,
            "launches": 90,
            "sale_value": 110,
            "base_icms": 110,
            "icms_value": 110,
            "status": 220,
        }
        for column_id in consult_columns:
            self.consult_tree.heading(
                column_id,
                text=consult_headings[column_id],
                command=lambda current=column_id: self.sort_consultation_tree_by(current),
            )
            self.consult_tree.column(column_id, width=consult_widths[column_id], anchor="center")
        self.consult_tree.grid(row=0, column=0, sticky="nsew")
        self.consult_tree.tag_configure("ok", background="#f4fff4")
        self.consult_tree.tag_configure("warning", background="#fff8e1")
        self.consult_tree.tag_configure("divergent", background="#ffe8e8")
        self.consult_tree.tag_configure("reduced", background="#e8f2ff")
        self.consult_tree.tag_configure("multi_supplier", background="#fff1d6")
        self.consult_tree.tag_configure("dynamic_rule", background="#eadcff")
        self.consult_tree.bind("<ButtonRelease-1>", self.handle_consult_tree_click)
        self.consult_tree.bind("<Double-1>", self.handle_consult_tree_double_click)
        self.consult_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.consult_tree))
        self.consult_tree.bind(
            "<Button-3>",
            lambda event: self.show_consult_runtime_rule_menu(event, self.consult_tree, self.consult_filtered_rows),
        )
        consult_scrollbar = ttk.Scrollbar(consult_grid_box, orient="vertical", command=self.consult_tree.yview)
        consult_scrollbar.grid(row=0, column=1, sticky="ns")
        consult_x_scrollbar = ttk.Scrollbar(consult_grid_box, orient="horizontal", command=self.consult_tree.xview)
        consult_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.consult_tree.configure(yscrollcommand=consult_scrollbar.set, xscrollcommand=consult_x_scrollbar.set)

        consult_footer = ttk.Frame(consult_view_tab)
        consult_footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(consult_footer, textvariable=self.consult_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(consult_footer, textvariable=self.consult_total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(consult_footer, textvariable=self.consult_total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(consult_footer, textvariable=self.consult_total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(consult_footer, textvariable=self.consult_total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(consult_footer, textvariable=self.consult_total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        sales_consult_process_box = ttk.LabelFrame(consult_process_scrollable, text="Consulta Saidas", padding=10)
        sales_consult_process_box.grid(row=2, column=0, sticky="nsew")
        ttk.Label(
            sales_consult_process_box,
            text=(
                "Carregue de 1 a 12 SPEDs para consulta em tela. Para documentos modelos 55/65 sem C170, "
                "importe tambem os XMLs para recompor os itens pela chave do C100."
            ),
            wraplength=820,
        ).pack(anchor="w")
        sales_consult_sped_row = ttk.Frame(sales_consult_process_box)
        sales_consult_sped_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(sales_consult_sped_row, textvariable=self.sales_consult_sped_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            sales_consult_sped_row,
            text="Selecionar SPEDs",
            style="Secondary.TButton",
            command=self.select_sales_consult_sped_files,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_sped_row,
            text="Processar",
            style="Primary.TButton",
            command=self.process_sales_consultation_speds,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_sped_row,
            text="Limpar Lista",
            style="Secondary.TButton",
            command=lambda: self.sales_consult_sped_paths_var.set(""),
        ).pack(side=LEFT, padx=(8, 0))
        sales_consult_xml_row = ttk.Frame(sales_consult_process_box)
        sales_consult_xml_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(sales_consult_xml_row, textvariable=self.sales_consult_xml_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            sales_consult_xml_row,
            text="Importar XMLs 55/65",
            style="Secondary.TButton",
            command=self.select_fiscal_consult_xml_sources,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_xml_row,
            text="Limpar XMLs",
            style="Secondary.TButton",
            command=lambda: [self.sales_consult_xml_paths_var.set(""), self.consult_xml_paths_var.set("")],
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            sales_consult_process_box,
            text="Use o botao Processar para carregar os SPEDs na aba Consulta Saidas.",
            style="Title.TLabel",
        ).pack(anchor="w", pady=(10, 0))
        ttk.Label(
            sales_consult_process_box,
            text=(
                "Quando houver documento 55/65 sem C170 no SPED, os itens serao buscados nos XMLs importados "
                "pela chave informada no C100."
            ),
            wraplength=820,
        ).pack(anchor="w", pady=(8, 0))

        contrib_process_canvas = Canvas(contrib_process_tab, background="#f7fbff", highlightthickness=0)
        contrib_process_scrollbar = ttk.Scrollbar(contrib_process_tab, orient="vertical", command=contrib_process_canvas.yview)
        contrib_process_scrollable = ttk.Frame(contrib_process_canvas, style="Card.TFrame", padding=(0, 0, 10, 0))
        contrib_process_scrollable.columnconfigure(0, weight=1)
        contrib_process_canvas.configure(yscrollcommand=contrib_process_scrollbar.set)
        contrib_process_canvas.grid(row=0, column=0, sticky="nsew")
        contrib_process_scrollbar.grid(row=0, column=1, sticky="ns")
        contrib_process_window = contrib_process_canvas.create_window((0, 0), window=contrib_process_scrollable, anchor="nw")

        def update_contrib_process_scroll_region(_event: object) -> None:
            contrib_process_canvas.configure(scrollregion=contrib_process_canvas.bbox("all"))

        def resize_contrib_process_content(event: object) -> None:
            contrib_process_canvas.itemconfigure(contrib_process_window, width=getattr(event, "width", 0))

        contrib_process_scrollable.bind("<Configure>", update_contrib_process_scroll_region)
        contrib_process_canvas.bind("<Configure>", resize_contrib_process_content)

        contrib_process_box = ttk.LabelFrame(contrib_process_scrollable, text="Processamento PIS/COFINS", padding=10)
        contrib_process_box.grid(row=3, column=0, sticky="nsew", pady=(10, 10))
        ttk.Label(
            contrib_process_box,
            text=(
                "Carregue de 1 a 12 arquivos do SPED Contribuicoes uma unica vez. O sistema usa a mesma lista "
                "para Entradas e Saidas e processa as duas consultas no mesmo fluxo."
            ),
            wraplength=820,
        ).pack(anchor="w")
        contrib_sped_row = ttk.Frame(contrib_process_box)
        contrib_sped_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(contrib_sped_row, textvariable=self.contrib_consult_sped_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            contrib_sped_row,
            text="Selecionar SPEDs",
            style="Secondary.TButton",
            command=self.select_contrib_consult_sped_files,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sped_row,
            text="Processar Entradas e Saidas",
            style="Primary.TButton",
            command=self.process_all_contrib_consultations,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sped_row,
            text="Limpar Lista",
            style="Secondary.TButton",
            command=lambda: [self.contrib_consult_sped_paths_var.set(""), self.contrib_sales_consult_sped_paths_var.set("")],
        ).pack(side=LEFT, padx=(8, 0))
        contrib_xml_row = ttk.Frame(contrib_process_box)
        contrib_xml_row.pack(fill="x", pady=(10, 0))
        ttk.Entry(contrib_xml_row, textvariable=self.contrib_consult_xml_paths_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(
            contrib_xml_row,
            text="Importar XMLs 55/65",
            style="Secondary.TButton",
            command=self.select_contrib_xml_sources,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_xml_row,
            text="Limpar XMLs",
            style="Secondary.TButton",
            command=lambda: [self.contrib_consult_xml_paths_var.set(""), self.contrib_sales_consult_xml_paths_var.set("")],
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_xml_row,
            text="Limpar Tela",
            style="Secondary.TButton",
            command=self.clear_form,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            contrib_process_box,
            text="Quando houver documento 55/65 sem C170 no SPED Contribuicoes, a recomposicao por XML sera aplicada automaticamente nas consultas de Entradas e Saidas.",
            wraplength=820,
        ).pack(anchor="w", pady=(8, 0))
        ttk.Label(
            contrib_process_box,
            text="As duas abas de consulta continuam separadas. O processamento agora so evita carga duplicada.",
            wraplength=820,
        ).pack(anchor="w", pady=(6, 0))

        sales_consult_filter_box = ttk.LabelFrame(sales_consult_view_tab, text="Filtros de Consulta", padding=10)
        sales_consult_filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        sales_consult_filter_box.columnconfigure(0, weight=2)
        sales_consult_filter_box.columnconfigure(1, weight=1)
        sales_consult_filter_box.columnconfigure(2, weight=1)
        sales_consult_filter_box.columnconfigure(3, weight=1)

        sales_periods_box = ttk.Frame(sales_consult_filter_box)
        sales_periods_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ttk.Label(sales_periods_box, text="Periodos para comparar:").pack(anchor="w")
        self.sales_consult_periods_checks_box = ttk.Frame(sales_periods_box)
        self.sales_consult_periods_checks_box.pack(fill="x", pady=(4, 0))
        self.sales_consult_period_check_vars: dict[str, BooleanVar] = {}

        sales_cst_box = ttk.Frame(sales_consult_filter_box)
        sales_cst_box.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(sales_cst_box, text="Filtro CST:").pack(anchor="w")
        ttk.Entry(sales_cst_box, textvariable=self.sales_consult_cst_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(sales_cst_box, text="Ex.: 000, 020, 060").pack(anchor="w", pady=(4, 0))

        sales_cfop_box = ttk.Frame(sales_consult_filter_box)
        sales_cfop_box.grid(row=0, column=2, sticky="ew", padx=(0, 10))
        ttk.Label(sales_cfop_box, text="Filtro CFOP:").pack(anchor="w")
        ttk.Entry(sales_cfop_box, textvariable=self.sales_consult_cfop_filter_var).pack(fill="x", pady=(4, 0))
        ttk.Label(sales_cfop_box, text="Ex.: 5101, 5405, 6108").pack(anchor="w", pady=(4, 0))

        sales_extra_filter_box = ttk.Frame(sales_consult_filter_box)
        sales_extra_filter_box.grid(row=0, column=3, sticky="ew")
        ttk.Label(sales_extra_filter_box, text="Status / Busca:").pack(anchor="w")
        ttk.Combobox(
            sales_extra_filter_box,
            textvariable=self.sales_consult_status_filter_var,
            values=("Todos", "Ok", "Sem debito", "Sem saida", "Com divergencia"),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Entry(sales_extra_filter_box, textvariable=self.sales_consult_search_var).pack(fill="x")
        ttk.Label(sales_extra_filter_box, text="Codigo ou descricao").pack(anchor="w", pady=(4, 0))

        sales_consult_filter_actions = ttk.Frame(sales_consult_filter_box)
        sales_consult_filter_actions.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        sales_consult_filter_actions_top = ttk.Frame(sales_consult_filter_actions)
        sales_consult_filter_actions_top.pack(fill="x", anchor="w")
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Aplicar Filtros",
            style="Primary.TButton",
            command=self.refresh_sales_consultation_tree,
        ).pack(side=LEFT)
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Limpar Filtros",
            style="Secondary.TButton",
            command=self.clear_sales_consultation_filters,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Exportar Filtro Atual",
            style="Secondary.TButton",
            command=self.export_current_sales_consultation_filter,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Saidas",
            style="Secondary.TButton",
            command=lambda: self.open_operation_summary_popup("Saida", "sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Comp. Diag. Debito",
            style="Secondary.TButton",
            command=lambda: self.open_credit_diagnostic_period_comparison_popup("sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Diag. Debito",
            style="Secondary.TButton",
            command=lambda: self.open_credit_diagnostic_popup("sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Curva ABC",
            style="Secondary.TButton",
            command=lambda: self.open_product_monthly_summary_popup("Saida", "sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Reducao BC",
            style="Secondary.TButton",
            command=lambda: self.open_reduction_items_popup("sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Apuracao",
            style="Secondary.TButton",
            command=lambda: self.open_apuracao_popup("sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            sales_consult_filter_actions_top,
            text="Espelho Docs",
            style="Secondary.TButton",
            command=lambda: self.open_fiscal_documents_popup("sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            sales_consult_filter_actions,
            text="Dica: clique em Docs ou Lancamentos para abrir o detalhamento.",
        ).pack(anchor="w", pady=(6, 0))

        sales_consult_results_box = ttk.Frame(sales_consult_view_tab)
        sales_consult_results_box.grid(row=1, column=0, sticky="nsew")
        sales_consult_results_box.columnconfigure(0, weight=1)
        sales_consult_results_box.rowconfigure(0, weight=1)

        sales_consult_grid_box = ttk.LabelFrame(sales_consult_results_box, text="Comparacao por Produto e Periodo", padding=10)
        sales_consult_grid_box.grid(row=0, column=0, sticky="nsew")
        sales_consult_grid_box.rowconfigure(0, weight=1)
        sales_consult_grid_box.columnconfigure(0, weight=1)
        sales_consult_grid_box.columnconfigure(1, weight=0)
        sales_consult_columns = (
            "period",
            "curve_abc",
            "code",
            "description",
            "supplier",
            "ncm",
            "cest",
            "cfop",
            "cst",
            "icms_rate",
            "effective_icms_rate",
            "documents",
            "launches",
            "sale_value",
            "base_icms",
            "icms_value",
            "status",
        )
        self.sales_consult_tree = ttk.Treeview(
            sales_consult_grid_box,
            columns=sales_consult_columns,
            show="headings",
            height=14,
            selectmode="extended",
        )
        for column_id in sales_consult_columns:
            self.sales_consult_tree.heading(
                column_id,
                text=consult_headings[column_id],
                command=lambda current=column_id: self.sort_sales_consultation_tree_by(current),
            )
            self.sales_consult_tree.column(column_id, width=consult_widths[column_id], anchor="center")
        self.sales_consult_tree.grid(row=0, column=0, sticky="nsew")
        self.sales_consult_tree.tag_configure("ok", background="#f4fff4")
        self.sales_consult_tree.tag_configure("warning", background="#fff8e1")
        self.sales_consult_tree.tag_configure("divergent", background="#ffe8e8")
        self.sales_consult_tree.tag_configure("reduced", background="#e8f2ff")
        self.sales_consult_tree.tag_configure("multi_supplier", background="#fff1d6")
        self.sales_consult_tree.tag_configure("dynamic_rule", background="#eadcff")
        self.sales_consult_tree.bind("<ButtonRelease-1>", self.handle_sales_consult_tree_click)
        self.sales_consult_tree.bind("<Double-1>", self.handle_sales_consult_tree_double_click)
        self.sales_consult_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.sales_consult_tree))
        self.sales_consult_tree.bind(
            "<Button-3>",
            lambda event: self.show_consult_runtime_rule_menu(event, self.sales_consult_tree, self.sales_consult_filtered_rows),
        )
        sales_consult_scrollbar = ttk.Scrollbar(sales_consult_grid_box, orient="vertical", command=self.sales_consult_tree.yview)
        sales_consult_scrollbar.grid(row=0, column=1, sticky="ns")
        sales_consult_x_scrollbar = ttk.Scrollbar(sales_consult_grid_box, orient="horizontal", command=self.sales_consult_tree.xview)
        sales_consult_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.sales_consult_tree.configure(yscrollcommand=sales_consult_scrollbar.set, xscrollcommand=sales_consult_x_scrollbar.set)

        sales_consult_footer = ttk.Frame(sales_consult_view_tab)
        sales_consult_footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(sales_consult_footer, textvariable=self.sales_consult_total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        contrib_consult_filter_box = ttk.LabelFrame(contrib_consult_view_tab, text="Filtros de Consulta", padding=10)
        contrib_consult_filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        contrib_consult_filter_box.columnconfigure(0, weight=2)
        contrib_consult_filter_box.columnconfigure(1, weight=1)
        contrib_consult_filter_box.columnconfigure(2, weight=1)
        contrib_consult_filter_box.columnconfigure(3, weight=1)
        contrib_periods_box = ttk.Frame(contrib_consult_filter_box)
        contrib_periods_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ttk.Label(contrib_periods_box, text="Periodos para comparar:").pack(anchor="w")
        self.contrib_consult_periods_checks_box = ttk.Frame(contrib_periods_box)
        self.contrib_consult_periods_checks_box.pack(fill="x", pady=(4, 0))
        self.contrib_consult_period_check_vars: dict[str, BooleanVar] = {}
        contrib_cst_box = ttk.Frame(contrib_consult_filter_box)
        contrib_cst_box.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(contrib_cst_box, text="Filtro CST PIS/COFINS:").pack(anchor="w")
        ttk.Entry(contrib_cst_box, textvariable=self.contrib_consult_cst_filter_var).pack(fill="x", pady=(4, 0))
        contrib_cfop_box = ttk.Frame(contrib_consult_filter_box)
        contrib_cfop_box.grid(row=0, column=2, sticky="ew", padx=(0, 10))
        ttk.Label(contrib_cfop_box, text="Filtro CFOP:").pack(anchor="w")
        ttk.Entry(contrib_cfop_box, textvariable=self.contrib_consult_cfop_filter_var).pack(fill="x", pady=(4, 0))
        contrib_extra_box = ttk.Frame(contrib_consult_filter_box)
        contrib_extra_box.grid(row=0, column=3, sticky="ew")
        ttk.Label(contrib_extra_box, text="Status / Busca:").pack(anchor="w")
        ttk.Combobox(
            contrib_extra_box,
            textvariable=self.contrib_consult_status_filter_var,
            values=("Todos", "Ok", "Sem credito", "Sem entrada", "Com divergencia"),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Entry(contrib_extra_box, textvariable=self.contrib_consult_search_var).pack(fill="x")
        ttk.Label(contrib_extra_box, text="Codigo ou descricao").pack(anchor="w", pady=(4, 0))
        contrib_consult_actions = ttk.Frame(contrib_consult_filter_box)
        contrib_consult_actions.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        contrib_consult_actions_top = ttk.Frame(contrib_consult_actions)
        contrib_consult_actions_top.pack(fill="x", anchor="w")
        ttk.Button(contrib_consult_actions_top, text="Aplicar Filtros", style="Primary.TButton", command=self.refresh_contrib_consultation_tree).pack(side=LEFT)
        ttk.Button(contrib_consult_actions_top, text="Limpar Filtros", style="Secondary.TButton", command=self.clear_contrib_consultation_filters).pack(side=LEFT, padx=(8, 0))
        ttk.Button(contrib_consult_actions_top, text="Exportar Filtro Atual", style="Secondary.TButton", command=self.export_current_contrib_consultation_filter).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Entradas",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_operation_summary_popup("Entrada"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Comp. Diag. Credito",
            style="Secondary.TButton",
            command=self.open_contrib_diagnostic_period_comparison_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Diag. Credito",
            style="Secondary.TButton",
            command=self.open_contrib_diagnostic_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Curva ABC",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_product_monthly_summary_popup("Entrada", "contrib_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Apuracao",
            style="Secondary.TButton",
            command=self.open_contrib_apuracao_popup,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_consult_actions_top,
            text="Espelho Docs",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_fiscal_documents_popup("contrib_consult", "Entrada"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(contrib_consult_filter_box, text="Dica: clique em Docs ou Lancamentos para abrir o detalhamento.").grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        contrib_consult_results_box = ttk.Frame(contrib_consult_view_tab)
        contrib_consult_results_box.grid(row=1, column=0, sticky="nsew")
        contrib_consult_results_box.columnconfigure(0, weight=1)
        contrib_consult_results_box.rowconfigure(0, weight=1)
        contrib_consult_grid_box = ttk.LabelFrame(contrib_consult_results_box, text="Comparacao por Produto e Periodo", padding=10)
        contrib_consult_grid_box.grid(row=0, column=0, sticky="nsew")
        contrib_consult_grid_box.rowconfigure(0, weight=1)
        contrib_consult_grid_box.columnconfigure(0, weight=1)
        contrib_consult_columns = ("period", "curve_abc", "code", "description", "supplier", "ncm", "cest", "cfop", "cst_pis", "cst_cofins", "pis_rate", "cofins_rate", "documents", "launches", "sale_value", "base_pis", "base_cofins", "pis_value", "cofins_value", "status")
        self.contrib_consult_tree = ttk.Treeview(contrib_consult_grid_box, columns=contrib_consult_columns, show="headings", height=14, selectmode="extended")
        contrib_headings = {"period": "Periodo", "curve_abc": "Curva ABC", "code": "Codigo", "description": "Descricao", "supplier": "Participante", "ncm": "NCM", "cest": "CEST", "cfop": "CFOP", "cst_pis": "CST PIS", "cst_cofins": "CST COFINS", "pis_rate": "Aliq PIS", "cofins_rate": "Aliq COFINS", "documents": "Docs", "launches": "Lancamentos", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "base_cofins": "Base COFINS", "pis_value": "Valor PIS", "cofins_value": "Valor COFINS", "status": "Status"}
        contrib_widths = {"period": 90, "curve_abc": 80, "code": 100, "description": 240, "supplier": 220, "ncm": 100, "cest": 90, "cfop": 90, "cst_pis": 80, "cst_cofins": 95, "pis_rate": 85, "cofins_rate": 95, "documents": 60, "launches": 90, "sale_value": 110, "base_pis": 100, "base_cofins": 110, "pis_value": 95, "cofins_value": 110, "status": 220}
        for column_id in contrib_consult_columns:
            self.contrib_consult_tree.heading(column_id, text=contrib_headings[column_id], command=lambda current=column_id: self.sort_contrib_consultation_tree_by(current))
            self.contrib_consult_tree.column(column_id, width=contrib_widths[column_id], anchor="center")
        self.contrib_consult_tree.grid(row=0, column=0, sticky="nsew")
        self.contrib_consult_tree.tag_configure("ok", background="#f4fff4")
        self.contrib_consult_tree.tag_configure("warning", background="#fff8e1")
        self.contrib_consult_tree.tag_configure("divergent", background="#ffe8e8")
        self.contrib_consult_tree.bind("<ButtonRelease-1>", self.handle_contrib_consult_tree_click)
        self.contrib_consult_tree.bind("<Double-1>", self.handle_contrib_consult_tree_double_click)
        self.contrib_consult_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.contrib_consult_tree))
        self.contrib_consult_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.contrib_consult_tree))
        contrib_consult_scrollbar = ttk.Scrollbar(contrib_consult_grid_box, orient="vertical", command=self.contrib_consult_tree.yview)
        contrib_consult_scrollbar.grid(row=0, column=1, sticky="ns")
        contrib_consult_x_scrollbar = ttk.Scrollbar(contrib_consult_grid_box, orient="horizontal", command=self.contrib_consult_tree.xview)
        contrib_consult_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.contrib_consult_tree.configure(yscrollcommand=contrib_consult_scrollbar.set, xscrollcommand=contrib_consult_x_scrollbar.set)
        contrib_consult_footer = ttk.Frame(contrib_consult_view_tab)
        contrib_consult_footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_pis_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_cofins_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_pis_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_consult_footer, textvariable=self.contrib_consult_total_cofins_var, style="Strong.TLabel").pack(side=LEFT)

        contrib_sales_filter_box = ttk.LabelFrame(contrib_sales_consult_view_tab, text="Filtros de Consulta", padding=10)
        contrib_sales_filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        contrib_sales_filter_box.columnconfigure(0, weight=2)
        contrib_sales_filter_box.columnconfigure(1, weight=1)
        contrib_sales_filter_box.columnconfigure(2, weight=1)
        contrib_sales_filter_box.columnconfigure(3, weight=1)
        contrib_sales_periods_box = ttk.Frame(contrib_sales_filter_box)
        contrib_sales_periods_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ttk.Label(contrib_sales_periods_box, text="Periodos para comparar:").pack(anchor="w")
        self.contrib_sales_consult_periods_checks_box = ttk.Frame(contrib_sales_periods_box)
        self.contrib_sales_consult_periods_checks_box.pack(fill="x", pady=(4, 0))
        self.contrib_sales_consult_period_check_vars: dict[str, BooleanVar] = {}
        contrib_sales_cst_box = ttk.Frame(contrib_sales_filter_box)
        contrib_sales_cst_box.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(contrib_sales_cst_box, text="Filtro CST PIS/COFINS:").pack(anchor="w")
        ttk.Entry(contrib_sales_cst_box, textvariable=self.contrib_sales_consult_cst_filter_var).pack(fill="x", pady=(4, 0))
        contrib_sales_cfop_box = ttk.Frame(contrib_sales_filter_box)
        contrib_sales_cfop_box.grid(row=0, column=2, sticky="ew", padx=(0, 10))
        ttk.Label(contrib_sales_cfop_box, text="Filtro CFOP:").pack(anchor="w")
        ttk.Entry(contrib_sales_cfop_box, textvariable=self.contrib_sales_consult_cfop_filter_var).pack(fill="x", pady=(4, 0))
        contrib_sales_extra_box = ttk.Frame(contrib_sales_filter_box)
        contrib_sales_extra_box.grid(row=0, column=3, sticky="ew")
        ttk.Label(contrib_sales_extra_box, text="Status / Busca:").pack(anchor="w")
        ttk.Combobox(
            contrib_sales_extra_box,
            textvariable=self.contrib_sales_consult_status_filter_var,
            values=("Todos", "Ok", "Sem debito", "Sem saida", "Com divergencia"),
            state="readonly",
        ).pack(fill="x", pady=(4, 6))
        ttk.Entry(contrib_sales_extra_box, textvariable=self.contrib_sales_consult_search_var).pack(fill="x")
        ttk.Label(contrib_sales_extra_box, text="Codigo ou descricao").pack(anchor="w", pady=(4, 0))
        contrib_sales_actions = ttk.Frame(contrib_sales_filter_box)
        contrib_sales_actions.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        contrib_sales_actions_top = ttk.Frame(contrib_sales_actions)
        contrib_sales_actions_top.pack(fill="x", anchor="w")
        ttk.Button(contrib_sales_actions_top, text="Aplicar Filtros", style="Primary.TButton", command=self.refresh_contrib_sales_consultation_tree).pack(side=LEFT)
        ttk.Button(contrib_sales_actions_top, text="Limpar Filtros", style="Secondary.TButton", command=self.clear_contrib_sales_consultation_filters).pack(side=LEFT, padx=(8, 0))
        ttk.Button(contrib_sales_actions_top, text="Exportar Filtro Atual", style="Secondary.TButton", command=self.export_current_contrib_sales_consultation_filter).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Saidas",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_operation_summary_popup("Saida", "contrib_sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Comp. Diag. Debito",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_diagnostic_period_comparison_popup("contrib_sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Diag. Debito",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_diagnostic_popup("contrib_sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Curva ABC",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_product_monthly_summary_popup("Saida", "contrib_sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Apuracao",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_apuracao_popup("contrib_sales_consult"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            contrib_sales_actions_top,
            text="Espelho Docs",
            style="Secondary.TButton",
            command=lambda: self.open_contrib_fiscal_documents_popup("contrib_sales_consult", "Saida"),
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(contrib_sales_filter_box, text="Dica: clique em Docs ou Lancamentos para abrir o detalhamento.").grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        contrib_sales_results_box = ttk.Frame(contrib_sales_consult_view_tab)
        contrib_sales_results_box.grid(row=1, column=0, sticky="nsew")
        contrib_sales_results_box.columnconfigure(0, weight=1)
        contrib_sales_results_box.rowconfigure(0, weight=1)
        contrib_sales_grid_box = ttk.LabelFrame(contrib_sales_results_box, text="Comparacao por Produto e Periodo", padding=10)
        contrib_sales_grid_box.grid(row=0, column=0, sticky="nsew")
        contrib_sales_grid_box.rowconfigure(0, weight=1)
        contrib_sales_grid_box.columnconfigure(0, weight=1)
        self.contrib_sales_consult_tree = ttk.Treeview(contrib_sales_grid_box, columns=contrib_consult_columns, show="headings", height=14, selectmode="extended")
        for column_id in contrib_consult_columns:
            self.contrib_sales_consult_tree.heading(column_id, text=contrib_headings[column_id], command=lambda current=column_id: self.sort_contrib_sales_consultation_tree_by(current))
            self.contrib_sales_consult_tree.column(column_id, width=contrib_widths[column_id], anchor="center")
        self.contrib_sales_consult_tree.grid(row=0, column=0, sticky="nsew")
        self.contrib_sales_consult_tree.tag_configure("ok", background="#f4fff4")
        self.contrib_sales_consult_tree.tag_configure("warning", background="#fff8e1")
        self.contrib_sales_consult_tree.tag_configure("divergent", background="#ffe8e8")
        self.contrib_sales_consult_tree.bind("<ButtonRelease-1>", self.handle_contrib_sales_consult_tree_click)
        self.contrib_sales_consult_tree.bind("<Double-1>", self.handle_contrib_sales_consult_tree_double_click)
        self.contrib_sales_consult_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.contrib_sales_consult_tree))
        self.contrib_sales_consult_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.contrib_sales_consult_tree))
        contrib_sales_scrollbar = ttk.Scrollbar(contrib_sales_grid_box, orient="vertical", command=self.contrib_sales_consult_tree.yview)
        contrib_sales_scrollbar.grid(row=0, column=1, sticky="ns")
        contrib_sales_x_scrollbar = ttk.Scrollbar(contrib_sales_grid_box, orient="horizontal", command=self.contrib_sales_consult_tree.xview)
        contrib_sales_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.contrib_sales_consult_tree.configure(yscrollcommand=contrib_sales_scrollbar.set, xscrollcommand=contrib_sales_x_scrollbar.set)
        contrib_sales_footer = ttk.Frame(contrib_sales_consult_view_tab)
        contrib_sales_footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_pis_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_cofins_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_pis_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(contrib_sales_footer, textvariable=self.contrib_sales_consult_total_cofins_var, style="Strong.TLabel").pack(side=LEFT)

        self.build_compare_invoices_tab(compare_tab)
        self.build_fiscal_diagnostic_tab(diagnostic_tab)
        self.build_xml_summary_tab(xml_root_tab)
        self.build_entry_exit_analysis_tab(entry_exit_root_tab)
        self.build_app_config_tab(app_config_tab)
        self.build_mysql_connection_tab(mysql_connection_tab)
        self.build_mysql_cadastro_tab(cadastro_tab)

        self.log = ttk.Treeview(frame, columns=("mensagem",), show="headings", height=1)
        self.log.heading("mensagem", text="Andamento")
        self.log.column("mensagem", width=1, anchor="center")

    def build_xml_summary_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        xml_notebook = ttk.Notebook(parent)
        xml_notebook.grid(row=0, column=0, sticky="nsew")
        summary_tab = ttk.Frame(xml_notebook, padding=6, style="Card.TFrame")
        credit_tab = ttk.Frame(xml_notebook, padding=6, style="Card.TFrame")
        xml_notebook.add(summary_tab, text="Resumo CFOP")
        xml_notebook.add(credit_tab, text="Creditos Entradas")
        parent = summary_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        setup_box = ttk.LabelFrame(parent, text="Leitura de XML", padding=10)
        setup_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        setup_box.columnconfigure(1, weight=1)
        setup_box.columnconfigure(4, weight=1)

        ttk.Label(
            setup_box,
            text="Resumo dos XMLs por CFOP",
            style="Title.TLabel",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))
        ttk.Label(setup_box, text="SPED empresa:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(setup_box, textvariable=self.xml_sped_path_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(setup_box, text="Selecionar", style="Secondary.TButton", command=self.select_xml_summary_sped_file).grid(row=1, column=2, sticky="w", pady=4)
        ttk.Label(setup_box, textvariable=self.xml_company_tax_id_var).grid(row=1, column=3, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(setup_box, text="Pasta/Arquivo XML:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(setup_box, textvariable=self.xml_source_path_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(setup_box, text="Pasta", style="Secondary.TButton", command=self.select_xml_summary_folder).grid(row=2, column=2, sticky="w", pady=4)
        ttk.Button(setup_box, text="Arquivo", style="Secondary.TButton", command=self.select_xml_summary_file).grid(row=2, column=3, sticky="w", padx=(8, 0), pady=4)

        operation_box = ttk.Frame(setup_box)
        operation_box.grid(row=3, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Label(operation_box, text="Operacao:").pack(side=LEFT)
        ttk.Radiobutton(
            operation_box,
            text="Todos",
            value="Todos",
            variable=self.xml_operation_scope_var,
            command=self.refresh_xml_summary_tree,
        ).pack(side=LEFT, padx=(8, 12))
        ttk.Radiobutton(
            operation_box,
            text="Entradas",
            value="Entrada",
            variable=self.xml_operation_scope_var,
            command=self.refresh_xml_summary_tree,
        ).pack(side=LEFT, padx=(0, 12))
        ttk.Radiobutton(
            operation_box,
            text="Saidas",
            value="Saida",
            variable=self.xml_operation_scope_var,
            command=self.refresh_xml_summary_tree,
        ).pack(side=LEFT)

        actions_box = ttk.Frame(setup_box)
        actions_box.grid(row=4, column=0, columnspan=4, sticky="w", pady=(14, 0))
        self.xml_process_btn = ttk.Button(actions_box, text="Processar XMLs", style="Primary.TButton", command=self.run_xml_summary)
        self.xml_process_btn.pack(side=LEFT)
        ttk.Button(actions_box, text="Exportar Excel", style="Secondary.TButton", command=lambda: self.export_xml_summary("xlsx")).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions_box, text="Exportar CSV", style="Secondary.TButton", command=lambda: self.export_xml_summary("csv")).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions_box, text="Exportar Chaves Lidas", style="Secondary.TButton", command=self.export_xml_document_keys).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions_box, text="Exportar Log Ignorados", style="Secondary.TButton", command=self.export_xml_ignored_log).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions_box, text="Limpar Tela", style="Secondary.TButton", command=self.clear_xml_screen).pack(side=LEFT, padx=(8, 0))

        ttk.Label(setup_box, textvariable=self.xml_status_var, wraplength=1280).grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))
        self.xml_progress_bar = ttk.Progressbar(
            setup_box,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.xml_progress_var,
        )
        self.xml_progress_bar.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(6, 0))

        results_box = ttk.LabelFrame(parent, text="Resumo por CFOP", padding=10)
        results_box.grid(row=1, column=0, sticky="nsew")
        results_box.columnconfigure(0, weight=1)
        results_box.rowconfigure(0, weight=1)

        columns = (
            "operation_type",
            "cfop",
            "documents",
            "items",
            "operation_value",
            "discount_value",
            "base_icms",
            "icms_value",
            "base_icms_st",
            "icms_st_value",
            "ipi_value",
            "base_pis",
            "pis_value",
            "base_cofins",
            "cofins_value",
        )
        headings = {
            "operation_type": "Tipo",
            "cfop": "CFOP",
            "documents": "Docs",
            "items": "Itens",
            "operation_value": "Total Operacao",
            "discount_value": "Desconto",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "ipi_value": "Valor IPI",
            "base_pis": "Base PIS",
            "pis_value": "Valor PIS",
            "base_cofins": "Base COFINS",
            "cofins_value": "Valor COFINS",
        }
        widths = {
            "operation_type": 90,
            "cfop": 80,
            "documents": 70,
            "items": 70,
            "operation_value": 130,
            "discount_value": 110,
            "base_icms": 120,
            "icms_value": 110,
            "base_icms_st": 120,
            "icms_st_value": 120,
            "ipi_value": 110,
            "base_pis": 110,
            "pis_value": 110,
            "base_cofins": 120,
            "cofins_value": 120,
        }

        self.xml_summary_tree = ttk.Treeview(results_box, columns=columns, show="headings", height=16, selectmode="extended")
        for column_id in columns:
            self.xml_summary_tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(self.xml_summary_tree, columns, headings)
        self.xml_summary_tree.grid(row=0, column=0, sticky="nsew")
        self.xml_summary_tree.tag_configure("cancelled", background="#ffd6d6", foreground="#8a1f1f")
        self.xml_summary_tree.bind("<Double-1>", self.handle_xml_summary_double_click)
        self.xml_summary_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.xml_summary_tree))
        self.xml_summary_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.xml_summary_tree))

        scroll_y = ttk.Scrollbar(results_box, orient="vertical", command=self.xml_summary_tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(results_box, orient="horizontal", command=self.xml_summary_tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.xml_summary_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        footer = ttk.Frame(parent)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, textvariable=self.xml_total_files_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_total_operation_files_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_total_cfops_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_total_operation_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_total_base_icms_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_total_icms_var, style="Strong.TLabel").pack(side=LEFT)
        self.build_xml_entry_credit_tab(credit_tab)

    def build_xml_entry_credit_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header,
            text="Creditos destacados em XMLs de entrada",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            textvariable=self.xml_credit_status_var,
            wraplength=920,
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Button(
            header,
            text="Processar Creditos",
            style="Primary.TButton",
            command=self.run_xml_entry_credit_summary,
        ).grid(row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Button(
            header,
            text="Exportar Excel",
            style="Secondary.TButton",
            command=lambda: self.export_xml_entry_credits("xlsx"),
        ).grid(row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Button(
            header,
            text="Exportar CSV",
            style="Secondary.TButton",
            command=lambda: self.export_xml_entry_credits("csv"),
        ).grid(row=0, column=3, sticky="e", padx=(8, 0))
        ttk.Button(
            header,
            text="Limpar Tela",
            style="Secondary.TButton",
            command=self.clear_xml_screen,
        ).grid(row=0, column=4, sticky="e", padx=(8, 0))

        results_notebook = ttk.Notebook(parent)
        results_notebook.grid(row=1, column=0, sticky="nsew")
        notes_tab = ttk.Frame(results_notebook, padding=8, style="Card.TFrame")
        cfop_tab = ttk.Frame(results_notebook, padding=8, style="Card.TFrame")
        results_notebook.add(notes_tab, text="Notas")
        results_notebook.add(cfop_tab, text="Agrupamento CFOP")

        notes_tab.columnconfigure(0, weight=1)
        notes_tab.rowconfigure(0, weight=1)
        note_columns = (
            "number",
            "series",
            "issue_date",
            "issuer_cnpj",
            "issuer_name",
            "cfops",
            "items",
            "operation_value",
            "base_icms",
            "icms_value",
            "base_icms_st",
            "icms_st_value",
            "ipi_value",
            "key",
        )
        note_headings = {
            "number": "Numero",
            "series": "Serie",
            "issue_date": "Data",
            "issuer_cnpj": "CNPJ Emitente",
            "issuer_name": "Fornecedor",
            "cfops": "CFOPs",
            "items": "Itens",
            "operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Credito ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "ipi_value": "Credito IPI",
            "key": "Chave",
        }
        note_widths = {
            "number": 90,
            "series": 70,
            "issue_date": 135,
            "issuer_cnpj": 135,
            "issuer_name": 260,
            "cfops": 130,
            "items": 70,
            "operation_value": 125,
            "base_icms": 115,
            "icms_value": 115,
            "base_icms_st": 115,
            "icms_st_value": 115,
            "ipi_value": 105,
            "key": 310,
        }
        self.xml_credit_notes_tree = ttk.Treeview(notes_tab, columns=note_columns, show="headings", height=15, selectmode="extended")
        for column_id in note_columns:
            self.xml_credit_notes_tree.column(column_id, width=note_widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(self.xml_credit_notes_tree, note_columns, note_headings)
        self.xml_credit_notes_tree.grid(row=0, column=0, sticky="nsew")
        self.xml_credit_notes_tree.bind("<Double-1>", self.handle_xml_credit_note_double_click)
        self.xml_credit_notes_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.xml_credit_notes_tree))
        self.xml_credit_notes_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.xml_credit_notes_tree))
        notes_y = ttk.Scrollbar(notes_tab, orient="vertical", command=self.xml_credit_notes_tree.yview)
        notes_y.grid(row=0, column=1, sticky="ns")
        notes_x = ttk.Scrollbar(notes_tab, orient="horizontal", command=self.xml_credit_notes_tree.xview)
        notes_x.grid(row=1, column=0, sticky="ew")
        self.xml_credit_notes_tree.configure(yscrollcommand=notes_y.set, xscrollcommand=notes_x.set)

        cfop_tab.columnconfigure(0, weight=1)
        cfop_tab.rowconfigure(0, weight=1)
        cfop_columns = ("cfop", "documents", "items", "operation_value", "discount_value", "base_icms", "icms_value", "base_icms_st", "icms_st_value", "ipi_value")
        cfop_headings = {
            "cfop": "CFOP",
            "documents": "Docs",
            "items": "Itens",
            "operation_value": "Total Operacao",
            "discount_value": "Desconto",
            "base_icms": "Base ICMS",
            "icms_value": "Credito ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "ipi_value": "Credito IPI",
        }
        cfop_widths = {
            "cfop": 90,
            "documents": 80,
            "items": 80,
            "operation_value": 140,
            "discount_value": 120,
            "base_icms": 125,
            "icms_value": 125,
            "base_icms_st": 125,
            "icms_st_value": 125,
            "ipi_value": 115,
        }
        self.xml_credit_cfop_tree = ttk.Treeview(cfop_tab, columns=cfop_columns, show="headings", height=15, selectmode="extended")
        for column_id in cfop_columns:
            self.xml_credit_cfop_tree.column(column_id, width=cfop_widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(self.xml_credit_cfop_tree, cfop_columns, cfop_headings)
        self.xml_credit_cfop_tree.grid(row=0, column=0, sticky="nsew")
        self.xml_credit_cfop_tree.bind("<Double-1>", self.handle_xml_credit_cfop_double_click)
        self.xml_credit_cfop_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.xml_credit_cfop_tree))
        self.xml_credit_cfop_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.xml_credit_cfop_tree))
        cfop_y = ttk.Scrollbar(cfop_tab, orient="vertical", command=self.xml_credit_cfop_tree.yview)
        cfop_y.grid(row=0, column=1, sticky="ns")
        cfop_x = ttk.Scrollbar(cfop_tab, orient="horizontal", command=self.xml_credit_cfop_tree.xview)
        cfop_x.grid(row=1, column=0, sticky="ew")
        self.xml_credit_cfop_tree.configure(yscrollcommand=cfop_y.set, xscrollcommand=cfop_x.set)

        footer = ttk.Frame(parent)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, textvariable=self.xml_credit_total_notes_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_credit_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_credit_total_operation_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_credit_total_base_icms_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_credit_total_icms_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=self.xml_credit_total_ipi_var, style="Strong.TLabel").pack(side=LEFT)

    def run_xml_entry_credit_summary(self) -> None:
        source_text = self.xml_source_path_var.get().strip()
        if not source_text:
            messagebox.showerror("XML", "Selecione uma pasta ou arquivo XML na aba Resumo CFOP.")
            return
        source_path = Path(source_text)
        if not source_path.exists():
            messagebox.showerror("XML", "O caminho selecionado nao existe.")
            return
        if source_path.is_file() and source_path.suffix.lower() != ".xml":
            messagebox.showerror("XML", "Selecione um arquivo .xml ou uma pasta com XMLs.")
            return
        company_tax_id = ""
        sped_text = self.xml_sped_path_var.get().strip()
        if sped_text:
            sped_path = Path(sped_text)
            if not sped_path.exists():
                messagebox.showerror("XML", "O SPED informado para identificar a empresa nao existe.")
                return
            company_tax_id = extract_company_tax_id_from_sped(sped_path)
            if not company_tax_id:
                messagebox.showerror("XML", "Nao foi possivel localizar CNPJ/CPF da empresa no registro 0000 do SPED.")
                return
            self.xml_company_tax_id_var.set(f"CNPJ/CPF empresa: {company_tax_id}")
        self.xml_credit_status_var.set("Iniciando leitura dos creditos de entrada.")
        threading.Thread(
            target=self.process_xml_entry_credit_background,
            args=(source_path, company_tax_id),
            daemon=True,
        ).start()

    def process_xml_entry_credit_background(self, source_path: Path, company_tax_id: str) -> None:
        try:
            invoice_rows, cfop_rows, stats = build_xml_entry_credit_rows(
                source_path,
                company_tax_id,
                self.schedule_xml_entry_credit_progress_update,
            )
            self.root.after(
                0,
                lambda invoice_rows=invoice_rows, cfop_rows=cfop_rows, stats=stats: self.render_xml_entry_credit_rows(
                    invoice_rows,
                    cfop_rows,
                    stats,
                ),
            )
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda message=error_message: self.handle_xml_entry_credit_error(message))

    def schedule_xml_entry_credit_progress_update(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.root.after(0, lambda percent=percent, message=message: self.xml_credit_status_var.set(f"{message} ({percent:.0f}%)"))

    def render_xml_entry_credit_rows(
        self,
        invoice_rows: list[dict[str, object]],
        cfop_rows: list[dict[str, object]],
        stats: dict[str, object],
    ) -> None:
        self.xml_credit_invoice_rows = invoice_rows
        self.xml_credit_cfop_rows = cfop_rows
        self.xml_credit_stats = stats
        self.refresh_xml_entry_credit_trees()
        skipped = int(stats.get("xml_skipped_by_company", 0) or 0)
        cancelled = int(stats.get("xml_cancelled", 0) or 0)
        cancelled_value = Decimal(stats.get("xml_cancelled_value_total", Decimal("0")))
        skipped_text = f" | fora da empresa: {skipped}" if skipped else ""
        cancelled_text = f" | cancelados excluidos: {cancelled} | valor cancelado: {format_decimal_sped(cancelled_value)}" if cancelled else ""
        self.xml_credit_status_var.set(
            "Leitura concluida. Valores exibidos sao destacados/potenciais para analise fiscal "
            f"de entradas; empresa usada: {stats.get('company_tax_id', '') or 'sem filtro'} | "
            f"XMLs invalidos: {stats.get('xml_ignored', 0)}{skipped_text}{cancelled_text}."
        )

    def refresh_xml_entry_credit_trees(self) -> None:
        if not hasattr(self, "xml_credit_notes_tree") or not hasattr(self, "xml_credit_cfop_tree"):
            return
        self.xml_credit_notes_tree.delete(*self.xml_credit_notes_tree.get_children())
        self.xml_credit_cfop_tree.delete(*self.xml_credit_cfop_tree.get_children())
        total_items = 0
        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        total_ipi = Decimal("0")

        for row_index, row in enumerate(self.xml_credit_invoice_rows):
            total_items += int(row.get("items", 0) or 0)
            total_operation += Decimal(row.get("operation_value", Decimal("0")))
            total_base += Decimal(row.get("base_icms", Decimal("0")))
            total_icms += Decimal(row.get("icms_value", Decimal("0")))
            total_ipi += Decimal(row.get("ipi_value", Decimal("0")))
            self.xml_credit_notes_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row.get("number", ""),
                    row.get("series", ""),
                    row.get("issue_date", ""),
                    row.get("issuer_cnpj", ""),
                    row.get("issuer_name", ""),
                    row.get("cfops", ""),
                    row.get("items", 0),
                    format_decimal_sped(Decimal(row.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms_st", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_st_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("ipi_value", Decimal("0")))),
                    row.get("key", ""),
                ),
            )

        for row_index, row in enumerate(self.xml_credit_cfop_rows):
            self.xml_credit_cfop_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row.get("cfop", ""),
                    row.get("document_count", 0),
                    row.get("items", 0),
                    format_decimal_sped(Decimal(row.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("discount_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms_st", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_st_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("ipi_value", Decimal("0")))),
                ),
            )

        self.xml_credit_total_notes_var.set(f"Notas entrada: {len(self.xml_credit_invoice_rows)}")
        self.xml_credit_total_items_var.set(f"Itens: {total_items}")
        self.xml_credit_total_operation_var.set(f"Total Operacao: {format_decimal_sped(total_operation)}")
        self.xml_credit_total_base_icms_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
        self.xml_credit_total_icms_var.set(f"Credito ICMS: {format_decimal_sped(total_icms)}")
        self.xml_credit_total_ipi_var.set(f"Credito IPI: {format_decimal_sped(total_ipi)}")

    def handle_xml_entry_credit_error(self, message: str) -> None:
        self.xml_credit_status_var.set("Erro durante a leitura dos creditos de entrada.")
        messagebox.showerror("XML", message)

    def handle_xml_credit_note_double_click(self, event: object) -> None:
        row_id = self.xml_credit_notes_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.xml_credit_invoice_rows):
            return
        self.open_xml_entry_credit_items_popup(self.xml_credit_invoice_rows[row_index], "Nota")

    def handle_xml_credit_cfop_double_click(self, event: object) -> None:
        row_id = self.xml_credit_cfop_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.xml_credit_cfop_rows):
            return
        self.open_xml_entry_credit_items_popup(self.xml_credit_cfop_rows[row_index], "CFOP")

    def open_xml_entry_credit_items_popup(self, source_row: dict[str, object], source_type: str) -> None:
        details = list(source_row.get("details", []))
        if not details:
            messagebox.showinfo("XML", "Nao ha itens para detalhar.")
            return
        title_piece = source_row.get("number", "") if source_type == "Nota" else source_row.get("cfop", "")
        dialog = Toplevel(self.root)
        dialog.title(f"Itens de Credito XML - {source_type} {title_piece}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1320, 620, 1080, 520, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        ttk.Label(container, text=f"Itens de entrada | {source_type} {title_piece}", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=1, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = (
            "number",
            "series",
            "issue_date",
            "item_no",
            "code",
            "description",
            "ncm",
            "cfop",
            "cst_icms",
            "quantity",
            "operation_value",
            "base_icms",
            "icms_rate",
            "icms_value",
            "base_icms_st",
            "icms_st_value",
            "ipi_value",
            "key",
        )
        headings = {
            "number": "Nota",
            "series": "Serie",
            "issue_date": "Data",
            "item_no": "Item",
            "code": "Codigo",
            "description": "Descricao",
            "ncm": "NCM",
            "cfop": "CFOP",
            "cst_icms": "CST",
            "quantity": "Qtd",
            "operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "icms_rate": "Aliq ICMS",
            "icms_value": "Credito ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "ipi_value": "Credito IPI",
            "key": "Chave",
        }
        widths = {
            "number": 85,
            "series": 60,
            "issue_date": 130,
            "item_no": 55,
            "code": 100,
            "description": 280,
            "ncm": 95,
            "cfop": 75,
            "cst_icms": 70,
            "quantity": 90,
            "operation_value": 120,
            "base_icms": 110,
            "icms_rate": 90,
            "icms_value": 110,
            "base_icms_st": 110,
            "icms_st_value": 110,
            "ipi_value": 100,
            "key": 300,
        }
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=15, selectmode="extended")
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        total_ipi = Decimal("0")
        for index, detail in enumerate(details):
            total_operation += Decimal(detail.get("operation_value", Decimal("0")))
            total_base += Decimal(detail.get("base_icms", Decimal("0")))
            total_icms += Decimal(detail.get("icms_value", Decimal("0")))
            total_ipi += Decimal(detail.get("ipi_value", Decimal("0")))
            tree.insert(
                "",
                END,
                iid=str(index),
                values=(
                    detail.get("number", ""),
                    detail.get("series", ""),
                    detail.get("issue_date", ""),
                    detail.get("item_no", ""),
                    detail.get("code", ""),
                    detail.get("description", ""),
                    detail.get("ncm", ""),
                    detail.get("cfop", ""),
                    detail.get("cst_icms", ""),
                    format_decimal_sped(Decimal(detail.get("quantity", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("icms_rate", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("base_icms_st", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("icms_st_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("ipi_value", Decimal("0")))),
                    detail.get("key", ""),
                ),
            )

        footer = ttk.Frame(container)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Itens: {len(details)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Total Operacao: {format_decimal_sped(total_operation)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Credito ICMS: {format_decimal_sped(total_icms)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Credito IPI: {format_decimal_sped(total_ipi)}", style="Strong.TLabel").pack(side=LEFT)

    def export_xml_entry_credits(self, output_type: str) -> None:
        if not self.xml_credit_invoice_rows:
            messagebox.showinfo("XML", "Nao ha creditos de entrada para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Exportar creditos XML de entrada",
            defaultextension=suffix,
            initialfile=f"creditos_xml_entradas{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return
        note_headers = ["Numero", "Serie", "Data", "CNPJ Emitente", "Fornecedor", "CFOPs", "Itens", "Total Operacao", "Base ICMS", "Credito ICMS", "Base ICMS ST", "Valor ICMS ST", "Credito IPI", "Chave", "Arquivo"]
        note_rows = [
            [
                row.get("number", ""),
                row.get("series", ""),
                row.get("issue_date", ""),
                row.get("issuer_cnpj", ""),
                row.get("issuer_name", ""),
                row.get("cfops", ""),
                row.get("items", 0),
                Decimal(row.get("operation_value", Decimal("0"))),
                Decimal(row.get("base_icms", Decimal("0"))),
                Decimal(row.get("icms_value", Decimal("0"))),
                Decimal(row.get("base_icms_st", Decimal("0"))),
                Decimal(row.get("icms_st_value", Decimal("0"))),
                Decimal(row.get("ipi_value", Decimal("0"))),
                row.get("key", ""),
                row.get("file_path", ""),
            ]
            for row in self.xml_credit_invoice_rows
        ]
        cfop_headers = ["CFOP", "Docs", "Itens", "Total Operacao", "Desconto", "Base ICMS", "Credito ICMS", "Base ICMS ST", "Valor ICMS ST", "Credito IPI"]
        cfop_rows = [
            [
                row.get("cfop", ""),
                row.get("document_count", 0),
                row.get("items", 0),
                Decimal(row.get("operation_value", Decimal("0"))),
                Decimal(row.get("discount_value", Decimal("0"))),
                Decimal(row.get("base_icms", Decimal("0"))),
                Decimal(row.get("icms_value", Decimal("0"))),
                Decimal(row.get("base_icms_st", Decimal("0"))),
                Decimal(row.get("icms_st_value", Decimal("0"))),
                Decimal(row.get("ipi_value", Decimal("0"))),
            ]
            for row in self.xml_credit_cfop_rows
        ]
        item_headers = ["Nota", "Serie", "Data", "Item", "Codigo", "Descricao", "NCM", "CFOP", "CST", "Qtd", "Total Operacao", "Base ICMS", "Aliq ICMS", "Credito ICMS", "Base ICMS ST", "Valor ICMS ST", "Credito IPI", "Chave", "Arquivo"]
        item_rows = [
            [
                detail.get("number", ""),
                detail.get("series", ""),
                detail.get("issue_date", ""),
                detail.get("item_no", ""),
                detail.get("code", ""),
                detail.get("description", ""),
                detail.get("ncm", ""),
                detail.get("cfop", ""),
                detail.get("cst_icms", ""),
                Decimal(detail.get("quantity", Decimal("0"))),
                Decimal(detail.get("operation_value", Decimal("0"))),
                Decimal(detail.get("base_icms", Decimal("0"))),
                Decimal(detail.get("icms_rate", Decimal("0"))),
                Decimal(detail.get("icms_value", Decimal("0"))),
                Decimal(detail.get("base_icms_st", Decimal("0"))),
                Decimal(detail.get("icms_st_value", Decimal("0"))),
                Decimal(detail.get("ipi_value", Decimal("0"))),
                detail.get("key", ""),
                detail.get("file_path", ""),
            ]
            for row in self.xml_credit_invoice_rows
            for detail in row.get("details", [])
            if isinstance(detail, dict)
        ]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, note_headers, note_rows)
            else:
                write_simple_excel_workbook(
                    output_path,
                    [
                        ("Creditos por Nota", note_headers, note_rows),
                        ("Agrupamento CFOP", cfop_headers, cfop_rows),
                        ("Itens", item_headers, item_rows),
                    ],
                )
            messagebox.showinfo("XML", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XML", str(exc))

    def select_xml_summary_sped_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o SPED da empresa",
            filetypes=[("Arquivo SPED", "*.txt"), ("Todos os arquivos", "*.*")],
        )
        if not selected:
            return
        self.xml_sped_path_var.set(selected)
        tax_id = extract_company_tax_id_from_sped(Path(selected))
        if tax_id:
            self.xml_company_tax_id_var.set(f"CNPJ/CPF empresa: {tax_id}")
            self.xml_status_var.set(f"SPED carregado. Empresa identificada: {tax_id}.")
        else:
            self.xml_company_tax_id_var.set("CNPJ/CPF empresa: nao encontrado no 0000")
            self.xml_status_var.set("SPED carregado, mas CNPJ/CPF nao foi encontrado no registro 0000.")

    def select_xml_summary_folder(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta com XMLs")
        if selected:
            self.xml_source_path_var.set(selected)

    def select_xml_summary_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o XML",
            filetypes=[("Arquivo XML", "*.xml"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.xml_source_path_var.set(selected)

    def update_xml_summary_progress(self, percent: int | float, message: str = "") -> None:
        normalized_percent = max(0, min(100, float(percent)))
        self.xml_progress_var.set(normalized_percent)
        if message:
            self.xml_status_var.set(f"{message} ({normalized_percent:.0f}%)")
        if hasattr(self, "xml_progress_bar"):
            self.xml_progress_bar.update_idletasks()

    def schedule_xml_summary_progress_update(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.root.after(0, lambda percent=percent, message=message: self.update_xml_summary_progress(percent, message))

    def run_xml_summary(self) -> None:
        source_text = self.xml_source_path_var.get().strip()
        if not source_text:
            messagebox.showerror("XML", "Selecione uma pasta ou arquivo XML.")
            return
        source_path = Path(source_text)
        if not source_path.exists():
            messagebox.showerror("XML", "O caminho selecionado nao existe.")
            return
        if source_path.is_file() and source_path.suffix.lower() != ".xml":
            messagebox.showerror("XML", "Selecione um arquivo .xml ou uma pasta com XMLs.")
            return
        company_tax_id = ""
        sped_text = self.xml_sped_path_var.get().strip()
        if sped_text:
            sped_path = Path(sped_text)
            if not sped_path.exists():
                messagebox.showerror("XML", "O SPED informado para identificar a empresa nao existe.")
                return
            company_tax_id = extract_company_tax_id_from_sped(sped_path)
            if not company_tax_id:
                messagebox.showerror("XML", "Nao foi possivel localizar CNPJ/CPF da empresa no registro 0000 do SPED.")
                return
            self.xml_company_tax_id_var.set(f"CNPJ/CPF empresa: {company_tax_id}")
        self.xml_process_btn.configure(state="disabled")
        self.update_xml_summary_progress(0, "Iniciando leitura dos XMLs.")
        threading.Thread(
            target=self.process_xml_summary_background,
            args=(source_path, company_tax_id),
            daemon=True,
        ).start()

    def process_xml_summary_background(self, source_path: Path, company_tax_id: str) -> None:
        try:
            rows, stats = build_xml_cfop_summary_rows(
                source_path,
                company_tax_id,
                "Todos",
                self.schedule_xml_summary_progress_update,
            )
            credit_invoice_rows, credit_cfop_rows, credit_stats = build_xml_entry_credit_rows(
                source_path,
                company_tax_id,
                self.schedule_xml_entry_credit_progress_update,
            )
            self.root.after(
                0,
                lambda rows=rows, stats=stats, credit_invoice_rows=credit_invoice_rows, credit_cfop_rows=credit_cfop_rows, credit_stats=credit_stats: (
                    self.render_xml_summary_rows(rows, stats),
                    self.render_xml_entry_credit_rows(credit_invoice_rows, credit_cfop_rows, credit_stats),
                ),
            )
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda message=error_message: self.handle_xml_summary_error(message))

    def render_xml_summary_rows(self, rows: list[dict[str, object]], stats: dict[str, object]) -> None:
        self.xml_process_btn.configure(state="normal")
        self.xml_summary_all_rows = rows
        self.xml_summary_stats = stats
        self.xml_ignored_log_rows = list(stats.get("ignored_log_rows", []))
        self.xml_document_log_rows = list(stats.get("document_log_rows", []))
        self.xml_cancelled_display_rows = list(stats.get("cancelled_display_rows", []))
        self.refresh_xml_summary_tree()
        skipped = int(stats.get("xml_skipped_by_company", 0) or 0)
        cancelled = int(stats.get("xml_cancelled", 0) or 0)
        cancelled_value = Decimal(stats.get("xml_cancelled_value_total", Decimal("0")))
        skipped_text = f" | fora da empresa: {skipped}" if skipped else ""
        cancelled_text = f" | cancelados: {cancelled} | valor cancelado: {format_decimal_sped(cancelled_value)}" if cancelled else ""
        self.update_xml_summary_progress(
            100,
            f"Leitura concluida. Empresa usada: {stats.get('company_tax_id', '') or 'sem filtro'} | XMLs invalidos: {stats.get('xml_ignored', 0)}{skipped_text}{cancelled_text}.",
        )

    def refresh_xml_summary_tree(self) -> None:
        if not hasattr(self, "xml_summary_tree"):
            return
        rows = filter_xml_summary_rows_by_scope(self.xml_summary_all_rows, self.xml_operation_scope_var.get())
        cancelled_rows = list(self.xml_cancelled_display_rows)
        self.xml_summary_rows = rows
        self.xml_summary_tree._xml_visible_rows = rows + cancelled_rows
        self.xml_summary_tree.delete(*self.xml_summary_tree.get_children())

        total_operation = Decimal("0")
        total_base_icms = Decimal("0")
        total_icms = Decimal("0")
        selected_document_keys: set[str] = set()
        for row_index, row in enumerate(rows):
            total_operation += Decimal(row.get("operation_value", Decimal("0")))
            total_base_icms += Decimal(row.get("base_icms", Decimal("0")))
            total_icms += Decimal(row.get("icms_value", Decimal("0")))
            for detail in row.get("details", []):
                if isinstance(detail, dict):
                    document_key = str(detail.get("key", "")).strip()
                    if document_key:
                        selected_document_keys.add(document_key)
            self.xml_summary_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row.get("operation_type", ""),
                    row.get("cfop", ""),
                    row.get("document_count", 0),
                    row.get("items", 0),
                    format_decimal_sped(Decimal(row.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("discount_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_icms_st", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("icms_st_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("ipi_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("cofins_value", Decimal("0")))),
                ),
            )

        for cancelled_row in cancelled_rows:
            self.xml_summary_tree.insert(
                "",
                END,
                iid=str(len(self.xml_summary_tree.get_children())),
                values=(
                    cancelled_row.get("operation_type", "Cancelado"),
                    cancelled_row.get("cfop", ""),
                    cancelled_row.get("document_count", 0),
                    cancelled_row.get("items", 0),
                    format_decimal_sped(Decimal(cancelled_row.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("discount_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("base_icms_st", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("icms_st_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("ipi_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("base_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("base_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(cancelled_row.get("cofins_value", Decimal("0")))),
                ),
                tags=("cancelled",),
            )

        self.xml_total_files_var.set(f"XMLs: {self.xml_summary_stats.get('xml_total', 0)}")
        selected_scope = normalize_compare_operation_scope(self.xml_operation_scope_var.get())
        scope_label = "Todos" if selected_scope == "Ambos" else selected_scope
        self.xml_total_operation_files_var.set(f"XMLs {scope_label}: {len(selected_document_keys)}")
        self.xml_total_cfops_var.set(f"CFOPs: {len(rows)}")
        self.xml_total_operation_var.set(f"Total Operacao: {format_decimal_sped(total_operation)}")
        self.xml_total_base_icms_var.set(f"Base ICMS: {format_decimal_sped(total_base_icms)}")
        self.xml_total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

    def handle_xml_summary_error(self, message: str) -> None:
        self.xml_process_btn.configure(state="normal")
        self.update_xml_summary_progress(0, "Erro durante a leitura dos XMLs.")
        messagebox.showerror("XML", message)

    def handle_xml_summary_double_click(self, event: object) -> None:
        row_id = self.xml_summary_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        visible_rows = getattr(self.xml_summary_tree, "_xml_visible_rows", self.xml_summary_rows)
        if row_index < 0 or row_index >= len(visible_rows):
            return
        self.open_xml_summary_details_popup(visible_rows[row_index])

    def open_xml_summary_details_popup(self, summary_row: dict[str, object]) -> None:
        details = list(summary_row.get("details", []))
        if not details:
            if str(summary_row.get("operation_type", "")).strip() == "Cancelado":
                messagebox.showinfo(
                    "XML cancelado",
                    (
                        f"Nota cancelada: {summary_row.get('number', '')}\n"
                        f"Chave: {summary_row.get('key', '')}\n"
                        f"Motivo: {summary_row.get('cancel_reason', '')}"
                    ),
                )
                return
            messagebox.showinfo("XML", "Nao ha detalhes para este agrupamento.")
            return

        operation_type = str(summary_row.get("operation_type", "")).strip()
        cfop = str(summary_row.get("cfop", "")).strip()
        dialog = Toplevel(self.root)
        dialog.title(f"XMLs do Agrupamento - {operation_type} / {cfop}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1320, 620, 1080, 520, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text=f"XMLs do agrupamento | Tipo {operation_type} | CFOP {cfop}",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=1, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = (
            "operation_type",
            "cfop",
            "key",
            "number",
            "series",
            "issue_date",
            "issuer_cnpj",
            "issuer_name",
            "recipient_cnpj",
            "recipient_name",
            "item_no",
            "code",
            "description",
            "operation_value",
            "base_icms",
            "icms_value",
            "ipi_value",
            "pis_value",
            "cofins_value",
            "file_path",
        )
        headings = {
            "operation_type": "Tipo",
            "cfop": "CFOP",
            "key": "Chave",
            "number": "Numero",
            "series": "Serie",
            "issue_date": "Data Emissao",
            "issuer_cnpj": "CNPJ Emitente",
            "issuer_name": "Emitente",
            "recipient_cnpj": "CNPJ/CPF Dest.",
            "recipient_name": "Destinatario",
            "item_no": "Item",
            "code": "Codigo",
            "description": "Descricao",
            "operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
            "ipi_value": "Valor IPI",
            "pis_value": "Valor PIS",
            "cofins_value": "Valor COFINS",
            "file_path": "Arquivo",
        }
        widths = {
            "operation_type": 90,
            "cfop": 80,
            "key": 310,
            "number": 90,
            "series": 70,
            "issue_date": 140,
            "issuer_cnpj": 130,
            "issuer_name": 220,
            "recipient_cnpj": 140,
            "recipient_name": 240,
            "item_no": 60,
            "code": 100,
            "description": 280,
            "operation_value": 120,
            "base_icms": 110,
            "icms_value": 110,
            "ipi_value": 100,
            "pis_value": 100,
            "cofins_value": 110,
            "file_path": 360,
        }
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=15, selectmode="extended")
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Double-1>", lambda event: self.handle_xml_summary_detail_double_click(event, tree, details))
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))

        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        total_operation = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for index, detail in enumerate(details):
            total_operation += Decimal(detail.get("operation_value", Decimal("0")))
            total_base += Decimal(detail.get("base_icms", Decimal("0")))
            total_icms += Decimal(detail.get("icms_value", Decimal("0")))
            tree.insert(
                "",
                END,
                iid=str(index),
                values=(
                    detail.get("operation_type", ""),
                    detail.get("cfop", ""),
                    detail.get("key", ""),
                    detail.get("number", ""),
                    detail.get("series", ""),
                    detail.get("issue_date", ""),
                    detail.get("issuer_cnpj", ""),
                    detail.get("issuer_name", ""),
                    detail.get("recipient_cnpj", ""),
                    detail.get("recipient_name", ""),
                    detail.get("item_no", ""),
                    detail.get("code", ""),
                    detail.get("description", ""),
                    format_decimal_sped(Decimal(detail.get("operation_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("icms_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("ipi_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(detail.get("cofins_value", Decimal("0")))),
                    detail.get("file_path", ""),
                ),
            )

        footer = ttk.Frame(container)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Linhas: {len(details)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Total Operacao: {format_decimal_sped(total_operation)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(total_icms)}", style="Strong.TLabel").pack(side=LEFT)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_xml_summary_details(details, operation_type, cfop, "xlsx"),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_xml_summary_details(details, operation_type, cfop, "csv"),
        ).pack(side=LEFT, padx=(8, 0))

    def handle_xml_summary_detail_double_click(
        self,
        event: object,
        tree: ttk.Treeview,
        details: list[dict[str, object]],
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(details):
            return

        source_path = str(details[row_index].get("file_path", "")).strip()
        if not source_path:
            messagebox.showwarning("Espelho XML", "Esta linha nao possui arquivo XML de origem.")
            return
        xml_path = Path(source_path)
        if not xml_path.exists():
            messagebox.showwarning("Espelho XML", "XML de origem nao encontrado no disco.")
            return
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            messagebox.showwarning("Espelho XML", "Nao foi possivel ler o XML selecionado.")
            return
        if not invoice.items:
            messagebox.showwarning("Espelho XML", "XML sem itens para montar o espelho da nota.")
            return
        self.open_compare_xml_invoice_mirror(invoice)

    def export_xml_summary_details(
        self,
        details: list[dict[str, object]],
        operation_type: str,
        cfop: str,
        output_type: str,
    ) -> None:
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Exportar popup XML",
            defaultextension=suffix,
            initialfile=f"xmls_{operation_type.lower()}_{cfop}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return
        headers = [
            "Tipo",
            "CFOP",
            "Chave",
            "Numero",
            "Serie",
            "Data Emissao",
            "CNPJ Emitente",
            "Emitente",
            "CNPJ/CPF Destinatario",
            "Destinatario",
            "Item",
            "Codigo",
            "Descricao",
            "Total Operacao",
            "Desconto",
            "Base ICMS",
            "Valor ICMS",
            "Base ICMS ST",
            "Valor ICMS ST",
            "Valor IPI",
            "Base PIS",
            "Valor PIS",
            "Base COFINS",
            "Valor COFINS",
            "Arquivo",
        ]
        rows = [
            [
                detail.get("operation_type", ""),
                detail.get("cfop", ""),
                detail.get("key", ""),
                detail.get("number", ""),
                detail.get("series", ""),
                detail.get("issue_date", ""),
                detail.get("issuer_cnpj", ""),
                detail.get("issuer_name", ""),
                detail.get("recipient_cnpj", ""),
                detail.get("recipient_name", ""),
                detail.get("item_no", ""),
                detail.get("code", ""),
                detail.get("description", ""),
                Decimal(detail.get("operation_value", Decimal("0"))),
                Decimal(detail.get("discount_value", Decimal("0"))),
                Decimal(detail.get("base_icms", Decimal("0"))),
                Decimal(detail.get("icms_value", Decimal("0"))),
                Decimal(detail.get("base_icms_st", Decimal("0"))),
                Decimal(detail.get("icms_st_value", Decimal("0"))),
                Decimal(detail.get("ipi_value", Decimal("0"))),
                Decimal(detail.get("base_pis", Decimal("0"))),
                Decimal(detail.get("pis_value", Decimal("0"))),
                Decimal(detail.get("base_cofins", Decimal("0"))),
                Decimal(detail.get("cofins_value", Decimal("0"))),
                detail.get("file_path", ""),
            ]
            for detail in details
        ]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                write_simple_excel_workbook(output_path, [("XMLs Agrupamento", headers, rows)])
            messagebox.showinfo("XML", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XML", str(exc))

    def export_xml_summary(self, output_type: str) -> None:
        visible_rows = list(getattr(getattr(self, "xml_summary_tree", None), "_xml_visible_rows", []))
        rows_to_export = [
            row for row in visible_rows
            if str(row.get("operation_type", "")).strip().lower() != "cancelado"
        ]
        if not rows_to_export:
            rows_to_export = filter_xml_summary_rows_by_scope(self.xml_summary_all_rows, self.xml_operation_scope_var.get())
        if not rows_to_export:
            messagebox.showinfo("XML", "Nao ha dados de XML para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Exportar resumo XML",
            defaultextension=suffix,
            initialfile=f"resumo_xml_cfop{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return
        headers = [
            "Tipo",
            "CFOP",
            "Docs",
            "Itens",
            "Total Operacao",
            "Desconto",
            "Base ICMS",
            "Valor ICMS",
            "Base ICMS ST",
            "Valor ICMS ST",
            "Valor IPI",
            "Base PIS",
            "Valor PIS",
            "Base COFINS",
            "Valor COFINS",
        ]
        rows = [
            [
                row.get("operation_type", ""),
                row.get("cfop", ""),
                row.get("document_count", 0),
                row.get("items", 0),
                Decimal(row.get("operation_value", Decimal("0"))),
                Decimal(row.get("discount_value", Decimal("0"))),
                Decimal(row.get("base_icms", Decimal("0"))),
                Decimal(row.get("icms_value", Decimal("0"))),
                Decimal(row.get("base_icms_st", Decimal("0"))),
                Decimal(row.get("icms_st_value", Decimal("0"))),
                Decimal(row.get("ipi_value", Decimal("0"))),
                Decimal(row.get("base_pis", Decimal("0"))),
                Decimal(row.get("pis_value", Decimal("0"))),
                Decimal(row.get("base_cofins", Decimal("0"))),
                Decimal(row.get("cofins_value", Decimal("0"))),
            ]
            for row in rows_to_export
        ]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                write_simple_excel_workbook(output_path, [("Resumo XML CFOP", headers, rows)])
            messagebox.showinfo("XML", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XML", str(exc))

    def export_xml_document_keys(self) -> None:
        selected_scope = normalize_compare_operation_scope(self.xml_operation_scope_var.get())
        rows_to_export = [
            row
            for row in self.xml_document_log_rows
            if selected_scope == "Ambos" or normalize_compare_operation_scope(row.get("operation_type", "")) == selected_scope
        ]
        if not rows_to_export:
            messagebox.showinfo("XML", "Nao ha chaves lidas para exportar.")
            return
        selected = filedialog.asksaveasfilename(
            title="Exportar chaves lidas dos XMLs",
            defaultextension=".xlsx",
            initialfile=f"xml_chaves_lidas_{selected_scope.lower() if selected_scope != 'Ambos' else 'todos'}.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not selected:
            return

        headers = [
            "Operacao",
            "Chave",
            "Numero",
            "Serie",
            "Data Emissao",
            "CNPJ Emitente",
            "Emitente",
            "CNPJ/CPF Destinatario",
            "Destinatario",
            "Valor",
            "Itens",
            "Arquivo",
        ]
        rows = [
            [
                row.get("operation_type", ""),
                row.get("key", ""),
                row.get("number", ""),
                row.get("series", ""),
                row.get("issue_date", ""),
                row.get("issuer_cnpj", ""),
                row.get("issuer_name", ""),
                row.get("recipient_cnpj", ""),
                row.get("recipient_name", ""),
                Decimal(row.get("total_value", Decimal("0"))),
                int(row.get("item_count", 0) or 0),
                row.get("file_path", ""),
            ]
            for row in rows_to_export
        ]
        total_value = sum((Decimal(row.get("total_value", Decimal("0"))) for row in rows_to_export), Decimal("0"))
        total_items = sum((int(row.get("item_count", 0) or 0) for row in rows_to_export), 0)
        total_row = ["Total", "", "", "", "", "", "", "", "", total_value, total_items, ""]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows + [total_row])
            else:
                write_simple_excel_workbook(output_path, [("Chaves Lidas", headers, rows)])
            messagebox.showinfo("XML", f"Chaves lidas exportadas com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XML", str(exc))

    def export_xml_ignored_log(self) -> None:
        if not self.xml_ignored_log_rows:
            messagebox.showinfo("XML", "Nao ha XMLs ignorados para exportar.")
            return
        selected = filedialog.asksaveasfilename(
            title="Exportar log de XMLs ignorados",
            defaultextension=".xlsx",
            initialfile="log_xmls_ignorados.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not selected:
            return

        headers = [
            "Motivo",
            "CNPJ/CPF Empresa Usado",
            "Chave",
            "Numero",
            "Serie",
            "Data Emissao",
            "CNPJ Emitente",
            "Emitente",
            "CNPJ/CPF Destinatario",
            "Destinatario",
            "Valor",
            "Arquivo",
        ]
        rows = [
            [
                row.get("reason", ""),
                row.get("company_tax_id", ""),
                row.get("key", ""),
                row.get("number", ""),
                row.get("series", ""),
                row.get("issue_date", ""),
                row.get("issuer_cnpj", ""),
                row.get("issuer_name", ""),
                row.get("recipient_cnpj", ""),
                row.get("recipient_name", ""),
                Decimal(row.get("total_value", Decimal("0"))),
                row.get("file_path", ""),
            ]
            for row in self.xml_ignored_log_rows
        ]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                write_simple_excel_workbook(output_path, [("XMLs Ignorados", headers, rows)])
            messagebox.showinfo("XML", f"Log gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XML", str(exc))

    def build_fiscal_diagnostic_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header,
            text="Diagnostico da Comparacao Fiscal",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            header,
            text="Atualizar diagnostico",
            style="Primary.TButton",
            command=lambda: self.refresh_fiscal_diagnostic_tab(show_message=True),
        ).grid(row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Button(
            header,
            text="Copiar conclusao",
            style="Secondary.TButton",
            command=lambda: self.copy_text_to_clipboard(
                self.compare_diagnostic_conclusion_var.get(),
                "Conclusao copiada para a area de transferencia.",
            ),
        ).grid(row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Button(
            header,
            text="Exportar investigacao",
            style="Secondary.TButton",
            command=lambda: self.export_compare_investigation_dataset(),
        ).grid(row=0, column=3, sticky="e", padx=(8, 0))

        conclusion_box = ttk.LabelFrame(parent, text="Conclusao para o usuario", padding=8)
        conclusion_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        conclusion_box.columnconfigure(0, weight=1)
        ttk.Label(
            conclusion_box,
            textvariable=self.compare_diagnostic_conclusion_var,
            wraplength=1280,
            justify=LEFT,
        ).grid(row=0, column=0, sticky="ew")
        ttk.Label(
            conclusion_box,
            textvariable=self.compare_diagnostic_status_var,
            style="Strong.TLabel",
            wraplength=1280,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        body = ttk.Frame(parent)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        self.compare_diagnostic_motive_tree = self.create_diagnostic_tree(
            body,
            "Motivos da Diferenca",
            ["Motivo", "Valor", "Qtd", "Explicacao", "Onde conferir"],
            [300, 110, 60, 430, 260],
            row=0,
            column=0,
        )
        self.compare_diagnostic_check_tree = self.create_diagnostic_tree(
            body,
            "Verificacao da Conciliacao",
            ["Item", "Valor/Status", "Leitura"],
            [230, 140, 560],
            row=0,
            column=1,
        )
        self.compare_diagnostic_composition_tree = self.create_diagnostic_tree(
            body,
            "Composicao Resumida",
            ["Componente", "Valor", "Leitura"],
            [260, 110, 620],
            row=1,
            column=0,
            columnspan=2,
        )

        footer = ttk.Frame(parent)
        footer.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(footer, textvariable=self.compare_diagnostic_footer_var, style="Strong.TLabel", wraplength=1280).pack(side=LEFT)

    def create_diagnostic_tree(
        self,
        parent: ttk.Frame,
        title: str,
        headers: list[str],
        widths: list[int],
        row: int,
        column: int,
        columnspan: int = 1,
    ) -> ttk.Treeview:
        box = ttk.LabelFrame(parent, text=title, padding=6)
        box.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=(0 if column == 0 else 6, 0), pady=(0, 8))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        columns = [f"c{index}" for index in range(len(headers))]
        tree = ttk.Treeview(box, columns=columns, show="headings", height=8, selectmode="extended")
        tree._diagnostic_headers = headers
        for index, column_id in enumerate(columns):
            tree.heading(column_id, text=headers[index])
            tree.column(column_id, width=widths[index] if index < len(widths) else 120, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        tree.bind("<Control-c>", lambda event, current_tree=tree: self.copy_treeview_selection(current_tree))
        tree.bind("<Button-3>", lambda event, current_tree=tree: self.show_treeview_copy_menu(event, current_tree))
        return tree

    def set_diagnostic_tree_rows(self, tree: ttk.Treeview, rows: list[list[object]]) -> None:
        tree.delete(*tree.get_children())
        for row in rows:
            tree.insert("", END, values=[format_decimal_sped(value) if isinstance(value, Decimal) else value for value in row])

    def refresh_fiscal_diagnostic_tab(self, show_message: bool = False) -> None:
        try:
            dataset = self.build_compare_investigation_dataset()
        except Exception as exc:
            self.compare_diagnostic_status_var.set(str(exc))
            self.compare_diagnostic_conclusion_var.set("Nenhum diagnostico disponivel.")
            self.compare_diagnostic_footer_var.set("")
            for tree_name in ("compare_diagnostic_motive_tree", "compare_diagnostic_check_tree", "compare_diagnostic_composition_tree"):
                tree = getattr(self, tree_name, None)
                if tree is not None:
                    tree.delete(*tree.get_children())
            if show_message:
                messagebox.showinfo("Diagnostico", str(exc))
            return

        self.compare_diagnostic_conclusion_var.set(str(dataset.get("simple_explanation_text", "")))
        self.compare_diagnostic_status_var.set(
            f"Status: {dataset.get('reconciliation_status', '')} | "
            f"Diferenca pendente: {format_decimal_sped(Decimal(dataset.get('pending_difference_total', Decimal('0.00'))))}"
        )
        self.compare_diagnostic_footer_var.set(
            f"Total comparacao: {format_decimal_sped(Decimal(dataset['comparison_total']))}   |   "
            f"Total fiscal: {format_decimal_sped(Decimal(dataset['fiscal_total']))}   |   "
            f"Diferenca: {format_decimal_sped(Decimal(dataset['total_difference']))}   |   "
            f"Diferenca explicada: {format_decimal_sped(Decimal(dataset.get('explained_difference_total', Decimal('0.00'))))}   |   "
            f"Diferenca pendente: {format_decimal_sped(Decimal(dataset.get('pending_difference_total', Decimal('0.00'))))}"
        )
        self.set_diagnostic_tree_rows(self.compare_diagnostic_motive_tree, dataset["simple_diagnostic_rows"])
        self.set_diagnostic_tree_rows(self.compare_diagnostic_check_tree, dataset["reconciliation_check_rows"])
        self.set_diagnostic_tree_rows(self.compare_diagnostic_composition_tree, dataset["simple_composition_rows"])
        if show_message:
            self.compare_status_var.set("Diagnostico fiscal atualizado.")

    def build_compare_invoices_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        setup_box = ttk.LabelFrame(parent, text="Importacao e Acao", padding=6)
        setup_box.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        setup_box.columnconfigure(1, weight=1)
        setup_box.columnconfigure(4, weight=1)

        ttk.Label(
            setup_box,
            text="Conferencia e lancamento de NF-e/NFC-e/NFS-e em SPED ICMS + Contribuicoes",
            style="Title.TLabel",
        ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 4))

        compare_fields = (
            (1, 0, "SPED ICMS:", self.compare_sped_icms_path_var, self.select_compare_sped_icms_file),
            (1, 3, "SPED Contrib.:", self.compare_sped_contrib_path_var, self.select_compare_sped_contrib_file),
            (2, 0, "Pasta XML:", self.compare_xml_folder_var, self.select_compare_xml_folder),
            (2, 3, "Planilha:", self.compare_sheet_path_var, self.select_compare_sheet_file),
        )
        for row_index, column_index, label, variable, command in compare_fields:
            ttk.Label(setup_box, text=label).grid(row=row_index, column=column_index, sticky="w", pady=2)
            ttk.Entry(setup_box, textvariable=variable).grid(row=row_index, column=column_index + 1, sticky="ew", padx=(6, 6), pady=2)
            ttk.Button(setup_box, text="Selecionar", style="Secondary.TButton", command=command).grid(row=row_index, column=column_index + 2, sticky="w", pady=2)

        controls_box = ttk.Frame(setup_box)
        controls_box.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(4, 0))
        controls_box.columnconfigure(2, weight=1)

        mode_box = ttk.Frame(controls_box)
        mode_box.grid(row=0, column=0, sticky="w")
        ttk.Label(mode_box, text="Comparar com:").pack(side=LEFT)
        ttk.Radiobutton(mode_box, text="SPED ICMS", value="icms", variable=self.compare_mode_var).pack(side=LEFT, padx=(8, 16))
        ttk.Radiobutton(mode_box, text="SPED Contribuicoes", value="contrib", variable=self.compare_mode_var).pack(side=LEFT)

        operation_box = ttk.Frame(controls_box)
        operation_box.grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Label(operation_box, text="Operacao no SPED:").pack(side=LEFT)
        ttk.Radiobutton(
            operation_box,
            text="Ambos",
            value="Ambos",
            variable=self.compare_operation_scope_var,
            command=self.clear_compare_results_for_operation_change,
        ).pack(side=LEFT, padx=(8, 12))
        ttk.Radiobutton(
            operation_box,
            text="Entradas",
            value="Entrada",
            variable=self.compare_operation_scope_var,
            command=self.clear_compare_results_for_operation_change,
        ).pack(side=LEFT, padx=(0, 12))
        ttk.Radiobutton(
            operation_box,
            text="Saidas",
            value="Saida",
            variable=self.compare_operation_scope_var,
            command=self.clear_compare_results_for_operation_change,
        ).pack(side=LEFT)

        actions_box = ttk.Frame(controls_box)
        actions_box.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.compare_run_btn = ttk.Button(actions_box, text="Comparar notas", style="Primary.TButton", command=self.run_compare_invoices)
        compare_action_buttons = [
            self.compare_run_btn,
            ttk.Button(actions_box, text="Marcar nao lancadas", style="Secondary.TButton", command=self.mark_all_missing_compare_invoices),
            ttk.Button(actions_box, text="Limpar marcas", style="Secondary.TButton", command=self.clear_compare_invoice_marks),
            ttk.Button(actions_box, text="Lancar marcadas", style="Secondary.TButton", command=self.launch_marked_compare_invoices),
            ttk.Button(actions_box, text="Investigar diferenca", style="Secondary.TButton", command=self.open_compare_difference_investigation),
            ttk.Button(actions_box, text="Exportar Excel", style="Secondary.TButton", command=self.export_compare_invoices_excel),
            ttk.Button(actions_box, text="Exportar CSV", style="Secondary.TButton", command=self.export_compare_invoices_csv),
            ttk.Button(actions_box, text="Exportar ignorados", style="Secondary.TButton", command=self.export_compare_ignored_xmls),
            ttk.Button(actions_box, text="Limpar Tela", style="Secondary.TButton", command=self.clear_compare_screen),
        ]
        for button_index, button in enumerate(compare_action_buttons):
            button.grid(row=button_index // 4, column=button_index % 4, sticky="w", padx=(0 if button_index % 4 == 0 else 8, 0), pady=(0 if button_index < 4 else 5, 0))

        ttk.Label(setup_box, textvariable=self.compare_status_var, wraplength=1280).grid(row=4, column=0, columnspan=6, sticky="w", pady=(4, 0))
        self.compare_progress_bar = ttk.Progressbar(
            setup_box,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.compare_progress_var,
        )
        self.compare_progress_bar.grid(row=5, column=0, columnspan=6, sticky="ew", pady=(3, 0))

        results_box = ttk.LabelFrame(parent, text="Resultado da Comparacao", padding=6)
        results_box.grid(row=1, column=0, sticky="nsew")
        results_box.columnconfigure(0, weight=1)
        results_box.rowconfigure(1, weight=1)

        ttk.Label(
            results_box,
            textvariable=self.compare_stats_var,
            style="Strong.TLabel",
            wraplength=1280,
            justify=LEFT,
        ).grid(row=0, column=0, sticky="ew", pady=(0, 4))

        tree_box = ttk.Frame(results_box)
        tree_box.grid(row=1, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = ("launch_mark", "status", "operation_type", "model", "key", "number", "series", "issue_date", "issuer_cnpj", "total_value", "file_path")
        headings = {
            "launch_mark": "Lancar",
            "status": "Status",
            "operation_type": "Operacao",
            "model": "Modelo",
            "key": "Chave",
            "number": "Numero",
            "series": "Serie",
            "issue_date": "Data Emissao",
            "issuer_cnpj": "CNPJ Emitente",
            "total_value": "Valor",
            "file_path": "Arquivo Origem",
        }
        widths = {
            "launch_mark": 70,
            "status": 140,
            "operation_type": 90,
            "model": 70,
            "key": 260,
            "number": 90,
            "series": 70,
            "issue_date": 150,
            "issuer_cnpj": 150,
            "total_value": 100,
            "file_path": 520,
        }
        self.compare_tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=10, selectmode="extended")
        for column_id in columns:
            self.compare_tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(self.compare_tree, columns, headings)
        self.compare_tree.grid(row=0, column=0, sticky="nsew")
        self.compare_tree.tag_configure("present", background="#f4fff4")
        self.compare_tree.tag_configure("missing", background="#ffe8e8")
        self.compare_tree.tag_configure("cancelled", background="#ffd6d6", foreground="#8a1f1f")
        self.compare_tree.bind("<Button-1>", self.toggle_compare_invoice_mark)
        self.compare_tree.bind("<Double-1>", self.open_compare_invoice_mirror_popup)
        self.compare_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(self.compare_tree))
        self.compare_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, self.compare_tree))

        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=self.compare_tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=self.compare_tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.compare_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    def build_entry_exit_analysis_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        setup_box = ttk.LabelFrame(parent, text="SPED Fiscal", padding=8)
        setup_box.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        setup_box.columnconfigure(1, weight=1)

        ttk.Label(
            setup_box,
            text="Analise de entrada e saida baseada nos resumos fiscais do SPED (C190).",
            style="Title.TLabel",
        ).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 5))
        ttk.Label(setup_box, text="SPEDs:").grid(row=1, column=0, sticky="w")
        ttk.Entry(setup_box, textvariable=self.entry_exit_sped_paths_var).grid(row=1, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(setup_box, text="Selecionar", style="Secondary.TButton", command=self.select_entry_exit_sped_files).grid(row=1, column=2, sticky="w")
        ttk.Button(setup_box, text="Usar consultas", style="Secondary.TButton", command=self.use_consult_speds_for_entry_exit_analysis).grid(row=1, column=3, sticky="w", padx=(6, 0))
        ttk.Button(setup_box, text="Processar", style="Primary.TButton", command=self.process_entry_exit_analysis).grid(row=1, column=4, sticky="w", padx=(6, 0))

        actions_box = ttk.Frame(setup_box)
        actions_box.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(6, 0))
        ttk.Button(actions_box, text="Exportar Excel", style="Secondary.TButton", command=self.export_entry_exit_analysis_excel).pack(side=LEFT)
        ttk.Label(actions_box, textvariable=self.entry_exit_status_var, wraplength=1050).pack(side=LEFT, padx=(12, 0), fill="x", expand=True)

        results_box = ttk.LabelFrame(parent, text="Resumo Entrada x Saida", padding=8)
        results_box.grid(row=1, column=0, sticky="nsew")
        results_box.columnconfigure(0, weight=1)
        results_box.rowconfigure(1, weight=1)

        totals_box = ttk.Frame(results_box)
        totals_box.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(totals_box, textvariable=self.entry_exit_total_saida_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 16))
        ttk.Label(totals_box, textvariable=self.entry_exit_total_entrada_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 16))
        ttk.Label(totals_box, textvariable=self.entry_exit_recolher_var, style="Strong.TLabel").pack(side=LEFT)

        analysis_notebook = ttk.Notebook(results_box)
        analysis_notebook.grid(row=1, column=0, sticky="nsew")

        entry_tab = ttk.Frame(analysis_notebook, padding=4, style="Card.TFrame")
        exit_tab = ttk.Frame(analysis_notebook, padding=4, style="Card.TFrame")
        analysis_notebook.add(entry_tab, text="Entradas")
        analysis_notebook.add(exit_tab, text="Saidas")

        columns = (
            "marker",
            "category",
            "cst_icms",
            "cfop",
            "icms_rate",
            "effective_rate",
            "icms_value",
            "total_operation_value",
            "base_icms",
            "base_icms_st",
            "icms_st_value",
            "base_ipi",
            "ipi_rate",
            "ipi_value",
            "row_count",
        )
        headings = {
            "marker": "Tipo",
            "category": "Classificacao",
            "cst_icms": "CST",
            "cfop": "CFOP",
            "icms_rate": "Aliq ICMS",
            "effective_rate": "Aliq Efetiva",
            "icms_value": "Valor ICMS",
            "total_operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "base_ipi": "Base IPI",
            "ipi_rate": "Aliq IPI",
            "ipi_value": "Valor IPI",
            "row_count": "Linhas C190",
        }
        widths = {
            "marker": 55,
            "category": 230,
            "cst_icms": 70,
            "cfop": 80,
            "icms_rate": 95,
            "effective_rate": 100,
            "icms_value": 115,
            "total_operation_value": 130,
            "base_icms": 120,
            "base_icms_st": 120,
            "icms_st_value": 120,
            "base_ipi": 110,
            "ipi_rate": 90,
            "ipi_value": 110,
            "row_count": 90,
        }
        footer_columns = (
            "label",
            "icms_value",
            "total_operation_value",
            "base_icms",
            "base_icms_st",
            "icms_st_value",
            "base_ipi",
            "ipi_value",
        )
        footer_headings = {
            "label": "Rodape",
            "icms_value": "Valor ICMS",
            "total_operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "base_ipi": "Base IPI",
            "ipi_value": "Valor IPI",
        }
        footer_widths = {
            "label": 300,
            "icms_value": 130,
            "total_operation_value": 150,
            "base_icms": 130,
            "base_icms_st": 130,
            "icms_st_value": 130,
            "base_ipi": 120,
            "ipi_value": 120,
        }

        def build_operation_panel(tab: ttk.Frame, operation_tag: str) -> tuple[ttk.Treeview, ttk.Treeview]:
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)

            detail_box = ttk.Frame(tab)
            detail_box.grid(row=0, column=0, sticky="nsew")
            detail_box.columnconfigure(0, weight=1)
            detail_box.rowconfigure(0, weight=1)

            tree = ttk.Treeview(detail_box, columns=columns, show="headings", height=11, selectmode="extended")
            for column_id in columns:
                tree.heading(column_id, text=headings[column_id])
                tree.column(column_id, width=widths[column_id], anchor="center")
            tree.grid(row=0, column=0, sticky="nsew")
            tree.tag_configure(operation_tag, background="#eef8ff" if operation_tag == "entrada" else "#fff4e8")
            tree.bind("<Control-c>", lambda event, current_tree=tree: self.copy_treeview_selection(current_tree))
            tree.bind("<Button-3>", lambda event, current_tree=tree: self.show_treeview_copy_menu(event, current_tree))

            scroll_y = ttk.Scrollbar(detail_box, orient="vertical", command=tree.yview)
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x = ttk.Scrollbar(detail_box, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

            footer_box = ttk.LabelFrame(tab, text="Rodape", padding=4)
            footer_box.grid(row=1, column=0, sticky="ew", pady=(6, 0))
            footer_box.columnconfigure(0, weight=1)
            footer_tree = ttk.Treeview(footer_box, columns=footer_columns, show="headings", height=6, selectmode="browse")
            for column_id in footer_columns:
                footer_tree.heading(column_id, text=footer_headings[column_id])
                footer_tree.column(column_id, width=footer_widths[column_id], anchor="center")
            footer_tree.grid(row=0, column=0, sticky="ew")
            footer_tree.tag_configure("total", background="#d8eefc")
            footer_tree.bind("<Control-c>", lambda event, current_tree=footer_tree: self.copy_treeview_selection(current_tree))
            footer_tree.bind("<Button-3>", lambda event, current_tree=footer_tree: self.show_treeview_copy_menu(event, current_tree))
            return tree, footer_tree

        self.entry_exit_entry_tree, self.entry_exit_entry_footer_tree = build_operation_panel(entry_tab, "entrada")
        self.entry_exit_exit_tree, self.entry_exit_exit_footer_tree = build_operation_panel(exit_tab, "saida")

    def select_entry_exit_sped_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Selecione os SPEDs fiscais",
            filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
        )
        if not selected:
            return
        current_paths = parse_selected_paths(self.entry_exit_sped_paths_var.get())
        seen = {str(path.resolve()).lower() for path in current_paths if path.exists()}
        for raw_path in selected:
            path = Path(raw_path)
            key = str(path.resolve()).lower()
            if key not in seen:
                current_paths.append(path)
                seen.add(key)
        self.entry_exit_sped_paths_var.set(format_selected_paths(current_paths))
        self.entry_exit_status_var.set(f"{len(current_paths)} SPED(s) selecionado(s).")

    def use_consult_speds_for_entry_exit_analysis(self) -> None:
        paths: list[Path] = []
        paths.extend(parse_selected_paths(self.consult_sped_paths_var.get()))
        paths.extend(parse_selected_paths(self.sales_consult_sped_paths_var.get()))
        unique_paths: list[Path] = []
        seen: set[str] = set()
        for path in paths:
            key = str(path.resolve()).lower() if path.exists() else str(path).lower()
            if key in seen:
                continue
            seen.add(key)
            unique_paths.append(path)
        self.entry_exit_sped_paths_var.set(format_selected_paths(unique_paths))
        self.entry_exit_status_var.set(f"SPEDs das consultas carregados: {len(unique_paths)} arquivo(s).")

    def process_entry_exit_analysis(self) -> None:
        sped_paths = parse_selected_paths(self.entry_exit_sped_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione ao menos um SPED fiscal.")
            return
        missing = [str(path) for path in sped_paths if not path.exists()]
        if missing:
            messagebox.showerror("Arquivo nao encontrado", "Os seguintes SPEDs nao existem:\n" + "\n".join(missing))
            return
        try:
            self.entry_exit_status_var.set("Processando resumos C190 dos SPEDs...")
            self.root.update_idletasks()
            rows, excel_rows, totals = build_entry_exit_analysis_rows(sped_paths)
            self.entry_exit_rows = rows
            self.entry_exit_excel_rows = excel_rows
            self.entry_exit_totals = totals
            self.refresh_entry_exit_analysis_tree()
            self.entry_exit_total_saida_var.set(
                f"Saidas ICMS: {format_decimal_sped(Decimal(totals.get('saida_icms', Decimal('0'))))}"
            )
            self.entry_exit_total_entrada_var.set(
                f"Entradas ICMS: {format_decimal_sped(Decimal(totals.get('entrada_icms', Decimal('0'))))}"
            )
            self.entry_exit_recolher_var.set(
                f"A recolher: {format_decimal_sped(Decimal(totals.get('recolher', Decimal('0'))))}"
            )
            self.entry_exit_status_var.set(f"Analise montada com {len(rows)} agrupamento(s) de CST/CFOP/aliquota.")
            self.write_audit_log("ANALISE_ENTRADA_SAIDA", f"speds={self.format_audit_paths(sped_paths)}; linhas={len(rows)}")
        except Exception as exc:
            self.entry_exit_status_var.set("Falha ao processar analise de entrada e saida.")
            messagebox.showerror("Analise Entrada e Saida", str(exc))

    def refresh_entry_exit_analysis_tree(self) -> None:
        if not hasattr(self, "entry_exit_entry_tree") or not hasattr(self, "entry_exit_exit_tree"):
            return
        operation_trees = {
            "Entrada": (self.entry_exit_entry_tree, self.entry_exit_entry_footer_tree, "entrada"),
            "Saida": (self.entry_exit_exit_tree, self.entry_exit_exit_footer_tree, "saida"),
        }
        for tree, footer_tree, _tag in operation_trees.values():
            tree.delete(*tree.get_children())
            footer_tree.delete(*footer_tree.get_children())

        for operation, (tree, _footer_tree, tag) in operation_trees.items():
            operation_rows = [row for row in self.entry_exit_rows if row.get("operation_type") == operation]
            for row in operation_rows:
                tree.insert(
                    "",
                    END,
                    values=(
                        row.get("marker", ""),
                        row.get("category", ""),
                        row.get("cst_icms", ""),
                        row.get("cfop", ""),
                        format_decimal_sped(Decimal(row.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("effective_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("icms_value", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("total_operation_value", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("base_icms", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("base_icms_st", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("icms_st_value", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("base_ipi", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("ipi_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("ipi_value", Decimal("0")))),
                        row.get("row_count", 0),
                    ),
                    tags=(tag,),
                )

        for operation, (_tree, footer_tree, _tag) in operation_trees.items():
            operation_rows = [row for row in self.entry_exit_rows if row.get("operation_type") == operation]
            for footer_row in build_entry_exit_footer_rows(operation_rows, operation):
                footer_tree.insert(
                    "",
                    END,
                    values=(
                        footer_row.get("label", ""),
                        format_decimal_sped(Decimal(footer_row.get("icms_value", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("total_operation_value", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("base_icms", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("base_icms_st", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("icms_st_value", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("base_ipi", Decimal("0")))),
                        format_decimal_sped(Decimal(footer_row.get("ipi_value", Decimal("0")))),
                    ),
                    tags=("total",) if str(footer_row.get("label", "")).lower().startswith("total") else (),
                )

    def export_entry_exit_analysis_excel(self) -> None:
        if not self.entry_exit_excel_rows:
            self.process_entry_exit_analysis()
            if not self.entry_exit_excel_rows:
                return
        selected = filedialog.asksaveasfilename(
            title="Salvar Analise Entrada e Saida",
            defaultextension=".xlsx",
            initialfile="entradas_x_saidas_sped.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not selected:
            return
        try:
            write_entry_exit_analysis_excel(Path(selected), self.entry_exit_excel_rows)
            self.entry_exit_status_var.set(f"Relatorio exportado em: {selected}")
            messagebox.showinfo("Analise Entrada e Saida", f"Relatorio gerado com sucesso:\n{selected}")
        except Exception as exc:
            messagebox.showerror("Analise Entrada e Saida", str(exc))

    def load_app_config(self) -> dict[str, str]:
        if not hasattr(self, "app_config_path"):
            return dict(APP_DEFAULT_CONFIG)
        if not self.app_config_path.exists():
            self.save_app_config_payload(APP_DEFAULT_CONFIG)
            return dict(APP_DEFAULT_CONFIG)
        try:
            loaded = json.loads(self.app_config_path.read_text(encoding="utf-8"))
        except Exception:
            loaded = {}
        config = dict(APP_DEFAULT_CONFIG)
        if isinstance(loaded, dict):
            for key, default_value in APP_DEFAULT_CONFIG.items():
                config[key] = str(loaded.get(key, default_value) or default_value)
        return config

    def save_app_config_payload(self, config: dict[str, str]) -> None:
        payload = {
            key: str(config.get(key, APP_DEFAULT_CONFIG[key]) or APP_DEFAULT_CONFIG[key])
            for key in APP_DEFAULT_CONFIG
        }
        self.app_config_path.parent.mkdir(parents=True, exist_ok=True)
        self.app_config_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")

    def get_app_form_config(self) -> dict[str, str]:
        return {
            "window_title": self.app_window_title_var.get().strip() or APP_DEFAULT_CONFIG["window_title"],
            "home_title": self.app_home_title_var.get().strip() or APP_DEFAULT_CONFIG["home_title"],
        }

    def apply_app_config(self) -> None:
        self.root.title(self.app_window_title_var.get().strip() or APP_DEFAULT_CONFIG["window_title"])

    def save_app_config(self) -> None:
        self.save_app_config_payload(self.get_app_form_config())
        self.apply_app_config()
        self.app_config_status_var.set("Configuracao da aplicacao salva.")
        self.log_message(f"Configuracao da aplicacao salva em: {self.app_config_path}")

    def reset_app_config(self) -> None:
        self.app_window_title_var.set(APP_DEFAULT_CONFIG["window_title"])
        self.app_home_title_var.set(APP_DEFAULT_CONFIG["home_title"])
        self.save_app_config()

    def build_app_config_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        config_box = ttk.LabelFrame(parent, text="Identificacao da Aplicacao", padding=10)
        config_box.grid(row=0, column=0, sticky="ew")
        config_box.columnconfigure(1, weight=1)

        ttk.Label(config_box, text="Titulo da janela").grid(row=0, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.app_window_title_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(8, 0),
            pady=(0, 8),
        )
        ttk.Label(config_box, text="Nome central da tela inicial").grid(row=1, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.app_home_title_var).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(8, 0),
            pady=(0, 8),
        )

        actions = ttk.Frame(config_box)
        actions.grid(row=2, column=1, sticky="e")
        ttk.Button(actions, text="Salvar Config", style="Primary.TButton", command=self.save_app_config).pack(side=LEFT)
        ttk.Button(actions, text="Restaurar Padrao", style="Secondary.TButton", command=self.reset_app_config).pack(side=LEFT, padx=(8, 0))
        ttk.Label(config_box, textvariable=self.app_config_status_var, wraplength=900).grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(10, 0),
        )

    def load_mysql_config(self) -> None:
        config = self.mysql_repo.load_config()
        self.db_host_var.set(config["host"])
        self.db_port_var.set(config["port"])
        self.db_user_var.set(config["user"])
        self.db_password_var.set(config["password"])
        self.db_name_var.set(config["database"])
        if not self.mysql_repo.mysql_available():
            self.db_status_var.set("Driver MySQL ausente. Instale mysql-connector-python na venv para habilitar os cadastros.")

    def get_mysql_form_config(self) -> dict[str, str]:
        return {
            "host": self.db_host_var.get().strip(),
            "port": self.db_port_var.get().strip() or "3306",
            "user": self.db_user_var.get().strip(),
            "password": self.db_password_var.get(),
            "database": self.db_name_var.get().strip() or "sped_icms",
        }

    def save_mysql_config(self, log_success: bool = True) -> None:
        self.mysql_repo.save_config(self.get_mysql_form_config())
        self.db_status_var.set("Configuracao MySQL salva.")
        if log_success:
            self.log_message(f"Configuracao MySQL salva em: {self.mysql_repo.config_path}")

    def test_mysql_connection(self) -> None:
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.test_connection()
            message = "Conexao com o servidor MySQL validada com sucesso."
            self.db_status_var.set(message)
            self.log_message(message)
            self.show_front_message("showinfo", "MySQL", message)
        except Exception as exc:
            self.db_status_var.set(f"Falha na conexao MySQL: {exc}")
            self.log_message(f"Falha na conexao MySQL: {exc}")
            messagebox.showerror("MySQL", str(exc))

    def create_mysql_schema(self) -> None:
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            message = f"Banco {self.db_name_var.get().strip()} e tabelas iniciais criados/atualizados."
            self.db_status_var.set(message)
            self.log_message(message)
            self.refresh_company_tree()
            self.show_front_message("showinfo", "MySQL", message)
        except Exception as exc:
            self.db_status_var.set(f"Falha ao criar schema MySQL: {exc}")
            self.log_message(f"Falha ao criar schema MySQL: {exc}")
            messagebox.showerror("MySQL", str(exc))

    def build_mysql_connection_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        config_box = ttk.LabelFrame(parent, text="Conexao MySQL", padding=10)
        config_box.grid(row=0, column=0, sticky="ew")
        for index in range(4):
            config_box.columnconfigure(index, weight=1 if index in (1, 3) else 0)

        ttk.Label(config_box, text="Host").grid(row=0, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.db_host_var).grid(row=0, column=1, sticky="ew", padx=(6, 12), pady=(0, 6))
        ttk.Label(config_box, text="Porta").grid(row=0, column=2, sticky="w")
        ttk.Entry(config_box, textvariable=self.db_port_var, width=10).grid(row=0, column=3, sticky="ew", padx=(6, 0), pady=(0, 6))
        ttk.Label(config_box, text="Banco").grid(row=1, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.db_name_var).grid(row=1, column=1, sticky="ew", padx=(6, 12), pady=(0, 6))
        ttk.Label(config_box, text="Usuario").grid(row=1, column=2, sticky="w")
        ttk.Entry(config_box, textvariable=self.db_user_var).grid(row=1, column=3, sticky="ew", padx=(6, 0), pady=(0, 6))
        ttk.Label(config_box, text="Senha").grid(row=2, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.db_password_var, show="*").grid(row=2, column=1, sticky="ew", padx=(6, 12))

        config_actions = ttk.Frame(config_box)
        config_actions.grid(row=2, column=2, columnspan=2, sticky="e")
        ttk.Button(config_actions, text="Salvar Config", style="Secondary.TButton", command=self.save_mysql_config).pack(side=LEFT)
        ttk.Button(config_actions, text="Testar Conexao", style="Secondary.TButton", command=self.test_mysql_connection).pack(side=LEFT, padx=(8, 0))
        ttk.Button(config_actions, text="Criar Banco/Tabelas", style="Primary.TButton", command=self.create_mysql_schema).pack(side=LEFT, padx=(8, 0))
        ttk.Label(config_box, textvariable=self.db_status_var, wraplength=1250).grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 0))

    def build_mysql_cadastro_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        body = ttk.Frame(parent)
        body.grid(row=0, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        product_box = ttk.LabelFrame(body, text="Cadastro de Produtos por Empresa", padding=10)
        product_box.grid(row=0, column=0, sticky="nsew")
        product_box.columnconfigure(0, weight=1)
        product_box.rowconfigure(2, weight=1)

        product_header = ttk.Frame(product_box)
        product_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        product_header.columnconfigure(1, weight=1)
        ttk.Label(product_header, text="Empresa para consulta/cadastro").grid(row=0, column=0, sticky="w")
        self.company_selector_combo = ttk.Combobox(product_header, textvariable=self.company_selector_var)
        self.company_selector_combo.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.company_selector_combo.bind("<<ComboboxSelected>>", self.handle_company_selector_change)
        self.company_selector_combo.bind("<KeyRelease>", self.filter_company_selector_options)
        self.company_selector_combo.bind("<Return>", self.apply_company_selector_typed_value)
        self.company_selector_combo.bind("<FocusIn>", self.restore_company_selector_options)
        self.company_selector_combo.bind("<FocusOut>", self.apply_company_selector_typed_value)
        ttk.Button(product_header, text="Nova Empresa", style="Secondary.TButton", command=self.open_company_popup).grid(row=0, column=2, sticky="e")
        ttk.Label(product_box, textvariable=self.product_company_var, style="Title.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 8))

        self.product_editor_notebook = ttk.Notebook(product_box)
        self.product_editor_notebook.grid(row=2, column=0, sticky="nsew")

        consulta_tab = ttk.Frame(self.product_editor_notebook, padding=10)
        consulta_tab.columnconfigure(0, weight=1)
        consulta_tab.rowconfigure(1, weight=1)

        cadastro_tab = ttk.Frame(self.product_editor_notebook, padding=10)
        cadastro_tab.columnconfigure(0, weight=1)
        cadastro_tab.rowconfigure(0, weight=1)

        self.product_editor_notebook.add(consulta_tab, text="Consulta de Produtos")
        self.product_editor_notebook.add(cadastro_tab, text="Cadastro / Edicao")

        product_filter_box = ttk.Frame(consulta_tab)
        product_filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        product_filter_box.columnconfigure(1, weight=1)
        ttk.Label(product_filter_box, text="Buscar").grid(row=0, column=0, sticky="w")
        ttk.Entry(product_filter_box, textvariable=self.product_search_var).grid(row=0, column=1, sticky="ew", padx=(6, 12))
        ttk.Label(product_filter_box, text="Status").grid(row=0, column=2, sticky="w")
        ttk.Combobox(product_filter_box, textvariable=self.product_status_filter_var, values=("Todos", "Ativo", "Inativo"), state="readonly", width=12).grid(row=0, column=3, sticky="ew", padx=(6, 0))
        ttk.Button(product_filter_box, text="Atualizar Lista", style="Secondary.TButton", command=self.refresh_product_tree).grid(row=0, column=4, sticky="e", padx=(8, 0))
        ttk.Button(product_filter_box, text="Editar Selecionado", style="Secondary.TButton", command=self.edit_selected_product).grid(row=0, column=5, sticky="e", padx=(8, 0))
        ttk.Button(product_filter_box, text="Novo Produto", style="Primary.TButton", command=self.clear_product_form).grid(row=0, column=6, sticky="e", padx=(8, 0))

        product_columns = (
            "id",
            "codigo",
            "codigo_origem",
            "descricao",
            "tipo_produto",
            "ncm",
            "unidade",
            "cst_icms_entrada",
            "cst_icms_saida",
            "icms_entrada",
            "icms_saida",
            "cst_pis_entrada",
            "cst_pis_saida",
            "pis_entrada",
            "pis_saida",
            "cst_cofins_entrada",
            "cst_cofins_saida",
            "cofins_entrada",
            "cofins_saida",
            "status",
        )
        product_tree_frame = ttk.Frame(consulta_tab)
        product_tree_frame.grid(row=1, column=0, sticky="nsew")
        product_tree_frame.columnconfigure(0, weight=1)
        product_tree_frame.rowconfigure(0, weight=1)
        self.product_tree = ttk.Treeview(product_tree_frame, columns=product_columns, show="headings", height=20, selectmode="browse")
        product_headings = {
            "id": "ID",
            "codigo": "Cod. Interno",
            "codigo_origem": "Cod. Origem",
            "descricao": "Descricao",
            "tipo_produto": "Tipo",
            "ncm": "NCM",
            "unidade": "Unid.",
            "cst_icms_entrada": "CST ICMS Ent.",
            "cst_icms_saida": "CST ICMS Sai.",
            "icms_entrada": "ICMS Ent.",
            "icms_saida": "ICMS Sai.",
            "cst_pis_entrada": "CST PIS Ent.",
            "cst_pis_saida": "CST PIS Sai.",
            "pis_entrada": "PIS Ent.",
            "pis_saida": "PIS Sai.",
            "cst_cofins_entrada": "CST COFINS Ent.",
            "cst_cofins_saida": "CST COFINS Sai.",
            "cofins_entrada": "COFINS Ent.",
            "cofins_saida": "COFINS Sai.",
            "status": "Status",
        }
        product_widths = {
            "id": 70,
            "codigo": 110,
            "codigo_origem": 130,
            "descricao": 280,
            "tipo_produto": 120,
            "ncm": 100,
            "unidade": 70,
            "cst_icms_entrada": 100,
            "cst_icms_saida": 100,
            "icms_entrada": 90,
            "icms_saida": 90,
            "cst_pis_entrada": 100,
            "cst_pis_saida": 100,
            "pis_entrada": 90,
            "pis_saida": 90,
            "cst_cofins_entrada": 115,
            "cst_cofins_saida": 115,
            "cofins_entrada": 110,
            "cofins_saida": 110,
            "status": 90,
        }
        for column_id in product_columns:
            self.product_tree.heading(column_id, text=product_headings[column_id])
            self.product_tree.column(column_id, width=product_widths[column_id], anchor="center")
        self.product_tree.tag_configure("tipo_revenda", background="#f4fff4")
        self.product_tree.tag_configure("tipo_fabricacao", background="#e8f2ff")
        self.product_tree.tag_configure("tipo_materia_prima", background="#fff8e1")
        self.product_tree.grid(row=0, column=0, sticky="nsew")
        self.product_tree.bind("<<TreeviewSelect>>", self.handle_product_tree_select)
        self.product_tree.bind("<Double-1>", self.open_selected_product_details_popup)
        product_scrollbar = ttk.Scrollbar(product_tree_frame, orient="vertical", command=self.product_tree.yview)
        product_scrollbar.grid(row=0, column=1, sticky="ns")
        product_x_scrollbar = ttk.Scrollbar(product_tree_frame, orient="horizontal", command=self.product_tree.xview)
        product_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.product_tree.configure(yscrollcommand=product_scrollbar.set, xscrollcommand=product_x_scrollbar.set)

        cadastro_container = ttk.Frame(cadastro_tab)
        cadastro_container.grid(row=0, column=0, sticky="nsew")
        cadastro_container.columnconfigure(0, weight=1)

        product_form = ttk.Frame(cadastro_container)
        product_form.grid(row=0, column=0, sticky="ew")
        product_form.columnconfigure(0, weight=3)
        product_form.columnconfigure(1, weight=2)
        product_form.columnconfigure(2, weight=2)

        cadastro_box = ttk.LabelFrame(product_form, text="Dados do Produto", padding=8)
        cadastro_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        cadastro_box.columnconfigure(1, weight=1)
        cadastro_box.columnconfigure(3, weight=1)
        cadastro_fields = (
            ("Codigo Interno", self.product_codigo_var, 0, 0, "readonly"),
            ("Status", self.product_status_var, 0, 2, "readonly"),
            ("Codigo Origem XML/SPED/GTIN", self.product_codigo_origem_var, 1, 0, "normal"),
            ("Descricao", self.product_descricao_var, 1, 2, "normal"),
            ("NCM", self.product_ncm_var, 2, 0, "normal"),
            ("Unidade", self.product_unidade_var, 2, 2, "normal"),
        )
        for label, variable, row_index, column_index, field_state in cadastro_fields:
            ttk.Label(cadastro_box, text=label).grid(row=row_index, column=column_index, sticky="w", pady=(0, 4))
            ttk.Entry(cadastro_box, textvariable=variable, state=field_state).grid(
                row=row_index,
                column=column_index + 1,
                sticky="ew",
                padx=(6, 10),
                pady=(0, 6),
            )
        ttk.Label(cadastro_box, text="Tipo do Produto").grid(row=3, column=0, sticky="w", pady=(0, 4))
        ttk.Combobox(
            cadastro_box,
            textvariable=self.product_tipo_var,
            values=("Revenda", "Fabricacao", "Materia Prima"),
            state="readonly",
        ).grid(row=3, column=1, sticky="ew", padx=(6, 10), pady=(0, 6))

        entrada_box = ttk.LabelFrame(product_form, text="Tributacao Entrada", padding=8)
        entrada_box.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        entrada_box.columnconfigure(1, weight=1)
        entrada_fields = (
            ("CST ICMS", self.product_cst_icms_entrada_var),
            ("% ICMS", self.product_icms_entrada_var),
            ("CST PIS", self.product_cst_pis_entrada_var),
            ("% PIS", self.product_pis_entrada_var),
            ("CST COFINS", self.product_cst_cofins_entrada_var),
            ("% COFINS", self.product_cofins_entrada_var),
        )
        for row_index, (label, variable) in enumerate(entrada_fields):
            ttk.Label(entrada_box, text=label).grid(row=row_index, column=0, sticky="w", pady=(0, 4))
            ttk.Entry(entrada_box, textvariable=variable).grid(row=row_index, column=1, sticky="ew", padx=(6, 0), pady=(0, 6))

        saida_box = ttk.LabelFrame(product_form, text="Tributacao Saida", padding=8)
        saida_box.grid(row=0, column=2, sticky="nsew")
        saida_box.columnconfigure(1, weight=1)
        saida_fields = (
            ("CST ICMS", self.product_cst_icms_saida_var),
            ("% ICMS", self.product_icms_saida_var),
            ("CST PIS", self.product_cst_pis_saida_var),
            ("% PIS", self.product_pis_saida_var),
            ("CST COFINS", self.product_cst_cofins_saida_var),
            ("% COFINS", self.product_cofins_saida_var),
        )
        for row_index, (label, variable) in enumerate(saida_fields):
            ttk.Label(saida_box, text=label).grid(row=row_index, column=0, sticky="w", pady=(0, 4))
            ttk.Entry(saida_box, textvariable=variable).grid(row=row_index, column=1, sticky="ew", padx=(6, 0), pady=(0, 6))

        product_actions = ttk.Frame(cadastro_container)
        product_actions.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(product_actions, text="Novo Produto", style="Secondary.TButton", command=self.clear_product_form).pack(side=LEFT)
        ttk.Button(product_actions, text="Desfazer Alteracoes", style="Secondary.TButton", command=self.undo_product_changes).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Salvar Produto", style="Primary.TButton", command=self.save_product_record).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Salvar Alteracoes", style="Primary.TButton", command=self.save_product_record).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Inativar", style="Secondary.TButton", command=self.deactivate_selected_product).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Reativar", style="Secondary.TButton", command=self.reactivate_selected_product).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Puxar Cod. Origem", style="Secondary.TButton", command=self.fill_product_origin_code_from_sources).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Importar Lote", style="Secondary.TButton", command=self.open_product_import_dialog).pack(side=LEFT, padx=(8, 0))

        self.product_search_var.trace_add("write", lambda *_args: self.refresh_product_tree())
        self.product_status_filter_var.trace_add("write", lambda *_args: self.refresh_product_tree())

    def refresh_company_tree(self, companies: list[dict[str, object]] | None = None) -> None:
        if not self.mysql_repo.mysql_available():
            return
        if companies is None:
            try:
                companies = self.mysql_repo.list_companies()
            except Exception as exc:
                self.db_status_var.set(f"MySQL indisponivel para listar empresas: {exc}")
                return
        if hasattr(self, "company_tree"):
            for item_id in self.company_tree.get_children():
                self.company_tree.delete(item_id)
        selector_values: list[str] = []
        selector_index: dict[int, str] = {}
        search_text = normalize_text(self.company_search_var.get())
        status_filter = self.company_status_filter_var.get().strip()
        for company in companies:
            status_label = "Ativa" if int(company.get("ativo", 1)) else "Inativa"
            selector_label = self.format_company_selector_label(company)
            selector_values.append(selector_label)
            selector_index[int(company["id"])] = selector_label
            if status_filter and status_filter != "Todos" and status_label != status_filter:
                continue
            searchable = normalize_text(f"{company['id']} {company['razao_social']} {company['cnpj']} {company.get('nome_fantasia', '')}")
            if search_text and search_text not in searchable:
                continue
            if hasattr(self, "company_tree"):
                self.company_tree.insert(
                    "",
                    END,
                    iid=str(company["id"]),
                    values=(company["id"], company["razao_social"], company["cnpj"], status_label),
                )
        self.company_selector_all_values = selector_values
        self.company_selector_index = selector_index
        if hasattr(self, "company_selector_combo"):
            self.company_selector_combo.configure(values=selector_values)
            if self.selected_company_id:
                selected_company = next((row for row in companies if int(row["id"]) == self.selected_company_id), None)
                if selected_company is not None:
                    self.company_selector_var.set(self.format_company_selector_label(selected_company))
            elif selector_values and not self.company_selector_var.get().strip():
                self.company_selector_var.set(selector_values[0])
                self.handle_company_selector_change()
        self.db_status_var.set(f"{len(companies)} empresa(s) carregada(s).")

    def format_company_selector_label(self, company: dict[str, object]) -> str:
        cnpj = str(company.get("cnpj") or "").strip()
        cnpj_text = cnpj if cnpj else "Sem CNPJ"
        return f"{company['id']} - {company['razao_social']} | CNPJ: {cnpj_text}"

    def restore_company_selector_options(self, _event: object = None) -> None:
        if hasattr(self, "company_selector_combo"):
            self.company_selector_combo.configure(values=self.company_selector_all_values)

    def filter_company_selector_options(self, _event: object = None) -> None:
        if not hasattr(self, "company_selector_combo"):
            return
        typed_text = self.company_selector_var.get().strip()
        normalized_filter = normalize_text(typed_text)
        if not normalized_filter:
            filtered_values = self.company_selector_all_values
        else:
            filtered_values = [
                value
                for value in self.company_selector_all_values
                if normalized_filter in normalize_text(value)
            ]
        self.company_selector_combo.configure(values=filtered_values or self.company_selector_all_values)

    def apply_company_selector_typed_value(self, _event: object = None) -> None:
        typed_text = self.company_selector_var.get().strip()
        if not typed_text:
            return
        if typed_text in self.company_selector_all_values:
            self.handle_company_selector_change()
            return
        normalized_filter = normalize_text(typed_text)
        matches = [
            value
            for value in self.company_selector_all_values
            if normalized_filter in normalize_text(value)
        ]
        if len(matches) == 1:
            self.company_selector_var.set(matches[0])
            self.handle_company_selector_change()

    def handle_company_tree_select(self, _event: object = None) -> None:
        if not hasattr(self, "company_tree"):
            return
        selected = self.company_tree.selection()
        if not selected:
            return
        company_id = int(selected[0])
        self.selected_company_id = company_id
        try:
            companies = self.mysql_repo.list_companies()
        except Exception as exc:
            self.db_status_var.set(f"Falha ao carregar empresa: {exc}")
            return
        selected_company = next((row for row in companies if int(row["id"]) == company_id), None)
        if selected_company is None:
            return
        self.company_id_var.set(str(selected_company["id"]))
        self.company_razao_var.set(str(selected_company["razao_social"] or ""))
        self.company_fantasia_var.set(str(selected_company["nome_fantasia"] or ""))
        self.company_cnpj_var.set(str(selected_company["cnpj"] or ""))
        self.company_ie_var.set(str(selected_company["inscricao_estadual"] or ""))
        self.company_status_var.set("Ativa" if int(selected_company.get("ativo", 1)) else "Inativa")
        self.product_company_var.set(f"Empresa selecionada: {selected_company['razao_social']} (ID {company_id})")
        if hasattr(self, "company_selector_combo"):
            self.company_selector_var.set(self.format_company_selector_label(selected_company))
        self.clear_product_form(clear_company_label=False)
        self.refresh_product_tree()

    def handle_company_selector_change(self, _event: object = None) -> None:
        selected_text = self.company_selector_var.get().strip()
        if not selected_text:
            return
        company_id_text = selected_text.split(" - ", 1)[0].strip()
        if not company_id_text.isdigit():
            return
        company_id = int(company_id_text)
        if hasattr(self, "company_tree") and self.company_tree.exists(str(company_id)):
            self.company_tree.selection_set(str(company_id))
            self.company_tree.focus(str(company_id))
        self.selected_company_id = company_id
        try:
            companies = self.mysql_repo.list_companies()
        except Exception as exc:
            self.db_status_var.set(f"Falha ao carregar empresa: {exc}")
            return
        selected_company = next((row for row in companies if int(row["id"]) == company_id), None)
        if selected_company is None:
            return
        self.company_id_var.set(str(selected_company["id"]))
        self.company_razao_var.set(str(selected_company["razao_social"] or ""))
        self.company_fantasia_var.set(str(selected_company["nome_fantasia"] or ""))
        self.company_cnpj_var.set(str(selected_company["cnpj"] or ""))
        self.company_ie_var.set(str(selected_company["inscricao_estadual"] or ""))
        self.company_status_var.set("Ativa" if int(selected_company.get("ativo", 1)) else "Inativa")
        self.product_company_var.set(f"Empresa selecionada: {selected_company['razao_social']} (ID {company_id})")
        self.clear_product_form(clear_company_label=False)
        self.refresh_product_tree()

    def clear_company_form(self) -> None:
        self.selected_company_id = None
        self.company_id_var.set("")
        self.company_razao_var.set("")
        self.company_fantasia_var.set("")
        self.company_cnpj_var.set("")
        self.company_ie_var.set("")
        self.company_status_var.set("Ativa")
        self.product_company_var.set("Empresa nao selecionada")
        self.clear_product_form()
        if hasattr(self, "product_tree"):
            for item_id in self.product_tree.get_children():
                self.product_tree.delete(item_id)

    def edit_selected_company(self) -> None:
        if not self.company_tree.selection():
            messagebox.showwarning("Empresa", "Selecione uma empresa para editar.")
            return
        self.handle_company_tree_select()

    def open_company_popup(self) -> None:
        dialog = Toplevel(self.root)
        dialog.title("Cadastrar Empresa")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 760, 420, 680, 360, margin_y=180)

        razao_var = StringVar()
        fantasia_var = StringVar()
        cnpj_var = StringVar()
        ie_var = StringVar()

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Nova Empresa", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(container, text="Razao Social").grid(row=1, column=0, sticky="w")
        ttk.Entry(container, textvariable=razao_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))
        ttk.Label(container, text="Nome Fantasia").grid(row=2, column=0, sticky="w")
        ttk.Entry(container, textvariable=fantasia_var).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))
        ttk.Label(container, text="CNPJ").grid(row=3, column=0, sticky="w")
        ttk.Entry(container, textvariable=cnpj_var).grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))
        ttk.Label(container, text="Inscricao Estadual").grid(row=4, column=0, sticky="w")
        ttk.Entry(container, textvariable=ie_var).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))

        actions = ttk.Frame(container)
        actions.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))

        def save_popup_company() -> None:
            razao_social = razao_var.get().strip()
            cnpj = cnpj_var.get().strip()
            if not razao_social:
                messagebox.showwarning("Empresa", "Informe a razao social da empresa.", parent=dialog)
                return
            if not cnpj:
                messagebox.showwarning("Empresa", "Informe o CNPJ da empresa.", parent=dialog)
                return
            data = {
                "razao_social": razao_social,
                "nome_fantasia": fantasia_var.get().strip(),
                "cnpj": cnpj,
                "inscricao_estadual": ie_var.get().strip(),
                "ativo": 1,
            }
            try:
                self.save_mysql_config(log_success=False)
                self.mysql_repo.ensure_schema()
                saved_id = self.mysql_repo.save_company(None, data)
            except Exception as exc:
                self.db_status_var.set(f"Falha ao salvar empresa: {exc}")
                self.log_message(f"Falha ao salvar empresa: {exc}")
                messagebox.showerror("Empresa", str(exc), parent=dialog)
                return
            self.refresh_company_tree()
            self.company_selector_var.set(
                self.company_selector_index.get(saved_id, f"{saved_id} - {razao_social} | CNPJ: {cnpj}")
            )
            self.handle_company_selector_change()
            dialog.destroy()
            self.log_message(f"Empresa {saved_id} cadastrada pelo popup.")

        ttk.Button(actions, text="Cancelar", style="Secondary.TButton", command=dialog.destroy).pack(side=LEFT)
        ttk.Button(actions, text="Salvar Empresa", style="Primary.TButton", command=save_popup_company).pack(side=LEFT, padx=(8, 0))

    def deactivate_selected_company(self) -> None:
        if not self.selected_company_id:
            messagebox.showwarning("Empresa", "Selecione uma empresa para inativar.")
            return
        if not messagebox.askyesno("Empresa", "Deseja inativar a empresa selecionada?"):
            return
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            self.mysql_repo.deactivate_company(self.selected_company_id)
            self.refresh_company_tree()
            self.company_status_var.set("Inativa")
            self.log_message(f"Empresa {self.selected_company_id} inativada.")
        except Exception as exc:
            self.db_status_var.set(f"Falha ao inativar empresa: {exc}")
            self.log_message(f"Falha ao inativar empresa: {exc}")
            messagebox.showerror("Empresa", str(exc))

    def reactivate_selected_company(self) -> None:
        if not self.selected_company_id:
            messagebox.showwarning("Empresa", "Selecione uma empresa para reativar.")
            return
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            self.mysql_repo.reactivate_company(self.selected_company_id)
            self.refresh_company_tree()
            self.company_status_var.set("Ativa")
            self.log_message(f"Empresa {self.selected_company_id} reativada.")
        except Exception as exc:
            self.db_status_var.set(f"Falha ao reativar empresa: {exc}")
            self.log_message(f"Falha ao reativar empresa: {exc}")
            messagebox.showerror("Empresa", str(exc))

    def save_company_record(self) -> None:
        razao_social = self.company_razao_var.get().strip()
        cnpj = self.company_cnpj_var.get().strip()
        if not razao_social:
            messagebox.showwarning("Empresa", "Informe a razao social da empresa.")
            return
        if not cnpj:
            messagebox.showwarning("Empresa", "Informe o CNPJ da empresa.")
            return
        data = {
            "razao_social": razao_social,
            "nome_fantasia": self.company_fantasia_var.get().strip(),
            "cnpj": cnpj,
            "inscricao_estadual": self.company_ie_var.get().strip(),
            "ativo": 1 if self.company_status_var.get().strip() != "Inativa" else 0,
        }
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            saved_id = self.mysql_repo.save_company(self.selected_company_id, data)
            self.selected_company_id = saved_id
            self.company_id_var.set(str(saved_id))
            self.refresh_company_tree()
            self.company_selector_var.set(self.company_selector_index.get(saved_id, self.company_selector_var.get()))
            if hasattr(self, "company_tree") and self.company_tree.exists(str(saved_id)):
                self.company_tree.selection_set(str(saved_id))
                self.company_tree.focus(str(saved_id))
            message = f"Empresa salva com ID {saved_id}."
            self.db_status_var.set(message)
            self.log_message(message)
        except Exception as exc:
            self.db_status_var.set(f"Falha ao salvar empresa: {exc}")
            self.log_message(f"Falha ao salvar empresa: {exc}")
            messagebox.showerror("Empresa", str(exc))

    def refresh_product_tree(self) -> None:
        if not hasattr(self, "product_tree"):
            return
        self.product_rows_by_id = {}
        for item_id in self.product_tree.get_children():
            self.product_tree.delete(item_id)
        if not self.selected_company_id or not self.mysql_repo.mysql_available():
            return
        try:
            products = self.mysql_repo.list_products(self.selected_company_id)
        except Exception as exc:
            self.db_status_var.set(f"Falha ao listar produtos: {exc}")
            return
        search_text = normalize_text(self.product_search_var.get())
        status_filter = self.product_status_filter_var.get().strip()
        for product in products:
            product_id = int(product["id"])
            self.product_rows_by_id[product_id] = dict(product)
            status_label = "Ativo" if int(product.get("ativo", 1)) else "Inativo"
            product_type = str(product.get("tipo_produto", "Revenda") or "Revenda").strip()
            if status_filter and status_filter != "Todos" and status_label != status_filter:
                continue
            searchable = normalize_text(
                f"{product['id']} {product['codigo']} {product.get('codigo_origem', '')} {product['descricao']} {product_type} {product['ncm']} {product['unidade']}"
            )
            if search_text and search_text not in searchable:
                continue
            item_tag = self.get_product_type_tag(product_type)
            self.product_tree.insert(
                "",
                END,
                iid=str(product_id),
                tags=(item_tag,),
                values=(
                    product_id,
                    product["codigo"],
                    product["codigo_origem"],
                    product["descricao"],
                    product_type,
                    product["ncm"],
                    product["unidade"],
                    product.get("cst_icms_entrada", ""),
                    product.get("cst_icms_saida", ""),
                    compare_format_float(float(product["icms_entrada"] or 0), 4),
                    compare_format_float(float(product["icms_saida"] or 0), 4),
                    product.get("cst_pis_entrada", ""),
                    product.get("cst_pis_saida", ""),
                    compare_format_float(float(product["pis_entrada"] or 0), 4),
                    compare_format_float(float(product["pis_saida"] or 0), 4),
                    product.get("cst_cofins_entrada", ""),
                    product.get("cst_cofins_saida", ""),
                    compare_format_float(float(product["cofins_entrada"] or 0), 4),
                    compare_format_float(float(product["cofins_saida"] or 0), 4),
                    status_label,
                ),
            )

    def get_product_type_tag(self, product_type: str) -> str:
        normalized_type = normalize_text(product_type)
        if normalized_type == "fabricacao":
            return "tipo_fabricacao"
        if normalized_type == "materia prima":
            return "tipo_materia_prima"
        return "tipo_revenda"

    def get_selected_product_row(self) -> dict[str, object] | None:
        if not hasattr(self, "product_tree"):
            return None
        selected = self.product_tree.selection()
        if not selected:
            return None
        product_id = int(selected[0])
        product = self.product_rows_by_id.get(product_id)
        if product is not None:
            return product
        try:
            products = self.mysql_repo.list_products(self.selected_company_id) if self.selected_company_id else []
        except Exception:
            return None
        product = next((row for row in products if int(row.get("id", 0)) == product_id), None)
        if product is not None:
            self.product_rows_by_id[product_id] = dict(product)
        return product

    def open_selected_product_details_popup(self, _event: object = None) -> None:
        product = self.get_selected_product_row()
        if product is None:
            return
        dialog = Toplevel(self.root)
        dialog.title(f"Detalhes do Produto {product.get('codigo', '')}")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 980, 620, 860, 520, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(3, weight=1)

        ttk.Label(container, text="Detalhes do Produto", style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        detail_fields = (
            ("ID", display_text(product.get("id", ""))),
            ("Codigo Interno", display_text(product.get("codigo", ""))),
            ("Codigo Origem", display_text(product.get("codigo_origem", ""))),
            ("Descricao", display_text(product.get("descricao", ""))),
            ("Tipo do Produto", display_text(product.get("tipo_produto", "Revenda"))),
            ("NCM", display_text(product.get("ncm", ""))),
            ("Unidade", display_text(product.get("unidade", "UN"))),
            ("Status", "Ativo" if int(product.get("ativo", 1)) else "Inativo"),
            ("CST ICMS Entrada", normalize_tax_code(product.get("cst_icms_entrada", ""), 3)),
            ("CST ICMS Saida", normalize_tax_code(product.get("cst_icms_saida", ""), 3)),
            ("% ICMS Entrada", compare_format_float(float(product.get("icms_entrada", 0) or 0), 4)),
            ("% ICMS Saida", compare_format_float(float(product.get("icms_saida", 0) or 0), 4)),
            ("CST PIS Entrada", normalize_tax_code(product.get("cst_pis_entrada", ""), 2)),
            ("CST PIS Saida", normalize_tax_code(product.get("cst_pis_saida", ""), 2)),
            ("% PIS Entrada", compare_format_float(float(product.get("pis_entrada", 0) or 0), 4)),
            ("% PIS Saida", compare_format_float(float(product.get("pis_saida", 0) or 0), 4)),
            ("CST COFINS Entrada", normalize_tax_code(product.get("cst_cofins_entrada", ""), 2)),
            ("CST COFINS Saida", normalize_tax_code(product.get("cst_cofins_saida", ""), 2)),
            ("% COFINS Entrada", compare_format_float(float(product.get("cofins_entrada", 0) or 0), 4)),
            ("% COFINS Saida", compare_format_float(float(product.get("cofins_saida", 0) or 0), 4)),
        )
        for index, (label_text, value_text) in enumerate(detail_fields, start=1):
            row_index = ((index - 1) // 2) + 1
            column_group = ((index - 1) % 2) * 2
            ttk.Label(container, text=label_text).grid(row=row_index, column=column_group, sticky="w", pady=(0, 6))
            ttk.Entry(container, state="readonly").grid(row=row_index, column=column_group + 1, sticky="ew", padx=(8, 14), pady=(0, 6))
            entry_widget = container.grid_slaves(row=row_index, column=column_group + 1)[0]
            entry_widget.configure(state="normal")
            entry_widget.delete(0, END)
            entry_widget.insert(0, value_text)
            entry_widget.configure(state="readonly")

        def edit_from_popup() -> None:
            dialog.destroy()
            current_product = self.get_selected_product_row() or product
            self.load_product_into_form(current_product, switch_tab=True)

        def duplicate_from_popup() -> None:
            dialog.destroy()
            self.load_product_into_form(product, duplicate=True, switch_tab=True)

        def toggle_status_from_popup() -> None:
            dialog.destroy()
            current_product = self.get_selected_product_row() or product
            if int(current_product.get("ativo", 1)):
                self.deactivate_selected_product()
            else:
                self.reactivate_selected_product()

        action_label = "Inativar" if int(product.get("ativo", 1)) else "Reativar"
        footer = ttk.Frame(container)
        footer.grid(row=((len(detail_fields) - 1) // 2) + 2, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(footer, text="Editar", style="Secondary.TButton", command=edit_from_popup).pack(side=LEFT)
        ttk.Button(footer, text="Duplicar Produto", style="Secondary.TButton", command=duplicate_from_popup).pack(side=LEFT, padx=(8, 0))
        ttk.Button(footer, text=action_label, style="Secondary.TButton", command=toggle_status_from_popup).pack(side=LEFT, padx=(8, 0))
        ttk.Button(footer, text="Fechar", style="Primary.TButton", command=dialog.destroy).pack(side=LEFT)

    def load_product_into_form(self, product: dict[str, object], duplicate: bool = False, switch_tab: bool = False) -> None:
        self.selected_product_id = None if duplicate else int(product.get("id", 0) or 0)
        self.product_id_var.set("" if duplicate else display_text(product.get("id", "")))
        self.product_codigo_var.set("" if duplicate else display_text(product.get("codigo", "")))
        self.product_codigo_origem_var.set(display_text(product.get("codigo_origem", "")))
        self.product_descricao_var.set(display_text(product.get("descricao", "")))
        self.product_tipo_var.set(display_text(product.get("tipo_produto", "Revenda"), "Revenda") or "Revenda")
        self.product_ncm_var.set(display_text(product.get("ncm", "")))
        self.product_unidade_var.set(display_text(product.get("unidade", "UN"), "UN") or "UN")
        self.product_cst_icms_entrada_var.set(normalize_tax_code(product.get("cst_icms_entrada", ""), 3))
        self.product_cst_icms_saida_var.set(normalize_tax_code(product.get("cst_icms_saida", ""), 3))
        self.product_icms_entrada_var.set(compare_format_float(float(product.get("icms_entrada", 0) or 0), 4))
        self.product_icms_saida_var.set(compare_format_float(float(product.get("icms_saida", 0) or 0), 4))
        self.product_cst_pis_entrada_var.set(normalize_tax_code(product.get("cst_pis_entrada", ""), 2))
        self.product_cst_pis_saida_var.set(normalize_tax_code(product.get("cst_pis_saida", ""), 2))
        self.product_pis_entrada_var.set(compare_format_float(float(product.get("pis_entrada", 0) or 0), 4))
        self.product_pis_saida_var.set(compare_format_float(float(product.get("pis_saida", 0) or 0), 4))
        self.product_cst_cofins_entrada_var.set(normalize_tax_code(product.get("cst_cofins_entrada", ""), 2))
        self.product_cst_cofins_saida_var.set(normalize_tax_code(product.get("cst_cofins_saida", ""), 2))
        self.product_cofins_entrada_var.set(compare_format_float(float(product.get("cofins_entrada", 0) or 0), 4))
        self.product_cofins_saida_var.set(compare_format_float(float(product.get("cofins_saida", 0) or 0), 4))
        self.product_status_var.set("Ativo" if duplicate else ("Ativo" if int(product.get("ativo", 1)) else "Inativo"))
        if switch_tab and hasattr(self, "product_editor_notebook"):
            self.product_editor_notebook.select(1)

    def handle_product_tree_select(self, _event: object = None) -> None:
        product = self.get_selected_product_row()
        if product is None:
            return
        self.load_product_into_form(product)

    def clear_product_form(self, clear_company_label: bool = True) -> None:
        self.selected_product_id = None
        self.product_id_var.set("")
        self.product_codigo_var.set("")
        self.product_codigo_origem_var.set("")
        self.product_descricao_var.set("")
        self.product_tipo_var.set("Revenda")
        self.product_ncm_var.set("")
        self.product_unidade_var.set("UN")
        self.product_cst_icms_entrada_var.set("")
        self.product_cst_icms_saida_var.set("")
        self.product_icms_entrada_var.set("")
        self.product_icms_saida_var.set("")
        self.product_cst_pis_entrada_var.set("")
        self.product_cst_pis_saida_var.set("")
        self.product_pis_entrada_var.set("")
        self.product_pis_saida_var.set("")
        self.product_cst_cofins_entrada_var.set("")
        self.product_cst_cofins_saida_var.set("")
        self.product_cofins_entrada_var.set("")
        self.product_cofins_saida_var.set("")
        self.product_status_var.set("Ativo")
        if clear_company_label and not self.selected_company_id:
            self.product_company_var.set("Empresa nao selecionada")
        if hasattr(self, "product_editor_notebook"):
            self.product_editor_notebook.select(1)

    def undo_product_changes(self) -> None:
        if self.selected_product_id:
            product = self.product_rows_by_id.get(self.selected_product_id)
            if product is None and self.selected_company_id:
                try:
                    products = self.mysql_repo.list_products(self.selected_company_id)
                except Exception as exc:
                    self.db_status_var.set(f"Falha ao recarregar produto: {exc}")
                    self.log_message(f"Falha ao recarregar produto: {exc}")
                    messagebox.showerror("Produto", str(exc))
                    return
                product = next((row for row in products if int(row.get("id", 0)) == self.selected_product_id), None)
                if product is not None:
                    self.product_rows_by_id[self.selected_product_id] = dict(product)
            if product is not None:
                self.load_product_into_form(product, switch_tab=True)
                self.show_front_message("showinfo", "Produto", "Alteracoes desfeitas com sucesso.")
                return
        self.clear_product_form()
        self.show_front_message("showinfo", "Produto", "Formulario limpo.")

    def edit_selected_product(self) -> None:
        if not self.product_tree.selection():
            messagebox.showwarning("Produto", "Selecione um produto para editar.")
            return
        self.handle_product_tree_select()
        if hasattr(self, "product_editor_notebook"):
            self.product_editor_notebook.select(1)

    def deactivate_selected_product(self) -> None:
        if not self.selected_product_id:
            messagebox.showwarning("Produto", "Selecione um produto para inativar.")
            return
        if not messagebox.askyesno("Produto", "Deseja inativar o produto selecionado?"):
            return
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            self.mysql_repo.deactivate_product(self.selected_product_id)
            self.refresh_product_tree()
            self.product_status_var.set("Inativo")
            self.log_message(f"Produto {self.selected_product_id} inativado.")
        except Exception as exc:
            self.db_status_var.set(f"Falha ao inativar produto: {exc}")
            self.log_message(f"Falha ao inativar produto: {exc}")
            messagebox.showerror("Produto", str(exc))

    def reactivate_selected_product(self) -> None:
        if not self.selected_product_id:
            messagebox.showwarning("Produto", "Selecione um produto para reativar.")
            return
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            self.mysql_repo.reactivate_product(self.selected_product_id)
            self.refresh_product_tree()
            self.product_status_var.set("Ativo")
            self.log_message(f"Produto {self.selected_product_id} reativado.")
        except Exception as exc:
            self.db_status_var.set(f"Falha ao reativar produto: {exc}")
            self.log_message(f"Falha ao reativar produto: {exc}")
            messagebox.showerror("Produto", str(exc))

    def fill_product_origin_code_from_sources(self) -> None:
        description = self.product_descricao_var.get().strip()
        ncm = self.product_ncm_var.get().strip()
        if not description and not ncm:
            messagebox.showwarning("Produto", "Informe ao menos descricao ou NCM para localizar o codigo de origem.")
            return
        candidates: list[dict[str, str]] = []
        xml_source_paths: list[Path] = []
        xml_source_paths.extend(parse_selected_paths(self.xml_path_var.get()))
        compare_xml_folder = self.compare_xml_folder_var.get().strip()
        if compare_xml_folder:
            xml_source_paths.append(Path(compare_xml_folder))
        xml_source_paths.extend(parse_selected_paths(self.sales_consult_xml_paths_var.get()))
        unique_xml_sources: list[Path] = []
        seen_xml_sources: set[str] = set()
        for path in xml_source_paths:
            normalized = str(path)
            if normalized not in seen_xml_sources:
                seen_xml_sources.add(normalized)
                unique_xml_sources.append(path)
        if unique_xml_sources:
            candidates.extend(build_product_origin_candidates_from_xml_sources(unique_xml_sources))

        sped_paths: list[Path] = []
        for raw_path in (
            self.sped_path_var.get().strip(),
            self.compare_sped_icms_path_var.get().strip(),
            self.compare_sped_contrib_path_var.get().strip(),
        ):
            if raw_path:
                sped_paths.append(Path(raw_path))
        sped_paths.extend(parse_selected_paths(self.consult_sped_paths_var.get()))
        sped_paths.extend(parse_selected_paths(self.sales_consult_sped_paths_var.get()))
        unique_sped_paths: list[Path] = []
        seen_sped_paths: set[str] = set()
        for path in sped_paths:
            normalized = str(path)
            if normalized not in seen_sped_paths and path.exists():
                seen_sped_paths.add(normalized)
                unique_sped_paths.append(path)
        for sped_path in unique_sped_paths[:5]:
            candidates.extend(build_product_origin_candidates_from_sped_file(sped_path))

        if not candidates:
            messagebox.showwarning("Produto", "Nenhum XML ou SPED carregado foi encontrado para pesquisar o codigo de origem.")
            return

        target_description = normalize_text(description)
        target_ncm = normalize_ncm(ncm)
        exact_matches = [
            item for item in candidates
            if normalize_text(item.get("description", "")) == target_description
            and (not target_ncm or normalize_ncm(item.get("ncm", "")) == target_ncm)
        ]
        fallback_matches = [
            item for item in candidates
            if target_description and target_description in normalize_text(item.get("description", ""))
        ]
        selected_match = (exact_matches or fallback_matches or [None])[0]
        if not selected_match:
            messagebox.showinfo("Produto", "Nao foi encontrado codigo de origem compativel nos XMLs/SPEDs carregados.")
            return
        origin_code = str(selected_match.get("ean", "")).strip() or str(selected_match.get("code", "")).strip() or str(selected_match.get("origin_code", "")).strip()
        self.product_codigo_origem_var.set(origin_code)
        if not self.product_ncm_var.get().strip():
            self.product_ncm_var.set(str(selected_match.get("ncm", "")).strip())
        message = f"Codigo de origem preenchido automaticamente: {origin_code}"
        self.db_status_var.set(message)
        self.log_message(message)

    def open_product_import_dialog(self) -> None:
        if not self.selected_company_id:
            messagebox.showwarning("Importacao", "Selecione uma empresa antes de importar produtos.")
            return

        dialog = Toplevel(self.root)
        dialog.title("Importar Produtos")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 920, 360, 820, 320, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(1, weight=1)

        source_type_var = StringVar(value="Consolidado")
        import_mode_var = StringVar(value="Criar e atualizar")
        source_path_var = StringVar(value="Usar XML + SPED Fiscal + SPED PIS/COFINS carregados no sistema")

        def update_default_source() -> None:
            if source_type_var.get() == "Consolidado":
                source_path_var.set("Usar XML + SPED Fiscal + SPED PIS/COFINS carregados no sistema")
            elif source_type_var.get() == "SPED 0200":
                source_path_var.set(self.sped_path_var.get().strip() or self.compare_sped_icms_path_var.get().strip())
            else:
                source_path_var.set(self.xml_path_var.get().strip() or self.compare_xml_folder_var.get().strip())

        def choose_source() -> None:
            if source_type_var.get() == "Consolidado":
                return
            if source_type_var.get() == "SPED 0200":
                selected = filedialog.askopenfilename(
                    title="Selecione o arquivo SPED para importar produtos 0200",
                    filetypes=[("Arquivos SPED", "*.txt *.sped *.efd"), ("Todos os arquivos", "*.*")],
                )
                if selected:
                    source_path_var.set(selected)
            else:
                selected_directory = filedialog.askdirectory(title="Selecione a pasta com XMLs")
                if selected_directory:
                    source_path_var.set(selected_directory)
                    source_type_var.set("XML")

        ttk.Label(container, text=f"Empresa selecionada: {self.product_company_var.get()}", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        ttk.Label(container, text="Fonte").grid(row=1, column=0, sticky="w")
        source_combo = ttk.Combobox(container, textvariable=source_type_var, values=("Consolidado", "SPED 0200", "XML"), state="readonly")
        source_combo.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        source_combo.bind("<<ComboboxSelected>>", lambda _event: update_default_source())

        ttk.Label(container, text="Modo").grid(row=2, column=0, sticky="w")
        ttk.Combobox(container, textvariable=import_mode_var, values=("Criar e atualizar", "Somente criar", "Somente atualizar"), state="readonly").grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))

        ttk.Label(container, text="Caminho").grid(row=3, column=0, sticky="w")
        ttk.Entry(container, textvariable=source_path_var).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(container, text="Selecionar", style="Secondary.TButton", command=choose_source).grid(row=3, column=2, sticky="w")

        ttk.Label(
            container,
            text=(
                "Consolidado usa XML + SPED Fiscal + SPED PIS/COFINS carregados no sistema. "
                "Prioridade: ICMS pelo SPED Fiscal, PIS/COFINS pelo SPED Contribuicoes e complemento pelo XML. "
                "SPED 0200 importa cadastro base. XML importa produtos por codigo/GTIN, descricao, NCM e aliquotas encontradas."
            ),
            wraplength=860,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 12))

        actions = ttk.Frame(container)
        actions.grid(row=5, column=0, columnspan=3, sticky="e")

        def run_import() -> None:
            source_text = source_path_var.get().strip()
            if source_type_var.get() != "Consolidado" and not source_text:
                messagebox.showwarning("Importacao", "Selecione a fonte da importacao.")
                return
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Preparando importacao",
                "Montando previa da importacao...",
            )
            try:
                preview = self.build_product_import_preview(
                    source_type_var.get(),
                    source_text,
                    import_mode_var.get(),
                    progress_callback=lambda current, total, message: self.update_progress_dialog(
                        progress_dialog,
                        progress_message_var,
                        progress_percent_var,
                        current,
                        total,
                        message,
                    ),
                )
            except Exception as exc:
                self.close_progress_dialog(progress_dialog)
                self.db_status_var.set(f"Falha na importacao de produtos: {exc}")
                self.log_message(f"Falha na importacao de produtos: {exc}")
                messagebox.showerror("Importacao", str(exc))
                return
            self.close_progress_dialog(progress_dialog)
            self.open_product_import_preview_dialog(dialog, preview, source_type_var.get(), import_mode_var.get(), source_text)

        ttk.Button(actions, text="Importar", style="Primary.TButton", command=run_import).pack(side=RIGHT)

    def open_product_import_preview_dialog(
        self,
        parent_dialog: Toplevel,
        preview: dict[str, object],
        source_type: str,
        import_mode: str,
        source_text: str,
    ) -> None:
        plan_rows = preview.get("plan_rows", [])
        summary = preview.get("summary", {})

        dialog = Toplevel(self.root)
        dialog.title("Previa da Importacao")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1360, 760, 1120, 620, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(container, text="Previa da Importacao de Produtos", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=(
                f"Fonte: {source_type} | Modo: {import_mode} | "
                f"Lidos: {summary.get('read', 0)} | Novos: {summary.get('created', 0)} | "
                f"Atualizacoes: {summary.get('updated', 0)} | Ignorados: {summary.get('skipped', 0)} | "
                f"Divergencias: {summary.get('divergences', 0)}"
            ),
            wraplength=1320,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        notebook = ttk.Notebook(container)
        notebook.grid(row=2, column=0, sticky="nsew")

        tab_specs = (
            ("Novos", [row for row in plan_rows if str(row.get("action", "")) == "create"]),
            ("Atualizar", [row for row in plan_rows if str(row.get("action", "")) == "update"]),
            ("Ignorados", [row for row in plan_rows if str(row.get("action", "")) == "ignored"]),
            ("Divergencias", [row for row in plan_rows if bool(row.get("divergent"))]),
        )
        columns = ("acao", "codigo_origem", "descricao", "ncm", "codigo_atual", "descricao_atual", "ncm_atual", "motivo")
        headings = {
            "acao": "Acao",
            "codigo_origem": "Cod. Origem",
            "descricao": "Descricao Importada",
            "ncm": "NCM Importado",
            "codigo_atual": "Cod. Atual",
            "descricao_atual": "Descricao Atual",
            "ncm_atual": "NCM Atual",
            "motivo": "Motivo",
        }
        widths = {
            "acao": 90,
            "codigo_origem": 130,
            "descricao": 260,
            "ncm": 110,
            "codigo_atual": 100,
            "descricao_atual": 260,
            "ncm_atual": 110,
            "motivo": 320,
        }

        for tab_name, rows in tab_specs:
            tab = ttk.Frame(notebook, padding=10, style="Card.TFrame")
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)
            notebook.add(tab, text=f"{tab_name} ({len(rows)})")
            tree_box = ttk.Frame(tab)
            tree_box.grid(row=0, column=0, sticky="nsew")
            tree_box.columnconfigure(0, weight=1)
            tree_box.rowconfigure(0, weight=1)
            tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=18, selectmode="extended")
            for column_id in columns:
                tree.heading(column_id, text=headings[column_id])
                tree.column(column_id, width=widths[column_id], anchor="center")
            tree.grid(row=0, column=0, sticky="nsew")
            tree.bind("<Control-c>", lambda event, current_tree=tree: self.copy_treeview_selection(current_tree))
            tree.bind("<Button-3>", lambda event, current_tree=tree: self.show_treeview_copy_menu(event, current_tree))
            scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            for row in rows:
                tree.insert(
                    "",
                    END,
                    values=(
                        str(row.get("action", "")).upper(),
                        str(row.get("codigo_origem", "")),
                        str(row.get("descricao", "")),
                        str(row.get("ncm", "")),
                        str(row.get("existing_codigo", "")),
                        str(row.get("existing_descricao", "")),
                        str(row.get("existing_ncm", "")),
                        str(row.get("reason", "")),
                    ),
                )

        footer = ttk.Frame(container)
        footer.grid(row=3, column=0, sticky="e", pady=(10, 0))

        def confirm_import() -> None:
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Importando produtos",
                "Gravando produtos no banco...",
            )
            try:
                result = self.apply_product_import_plan(
                    plan_rows,
                    progress_callback=lambda current, total, message: self.update_progress_dialog(
                        progress_dialog,
                        progress_message_var,
                        progress_percent_var,
                        current,
                        total,
                        message,
                    ),
                )
            except Exception as exc:
                self.close_progress_dialog(progress_dialog)
                self.db_status_var.set(f"Falha na importacao de produtos: {exc}")
                self.log_message(f"Falha na importacao de produtos: {exc}")
                messagebox.showerror("Importacao", str(exc))
                return
            self.close_progress_dialog(progress_dialog)
            dialog.destroy()
            parent_dialog.destroy()
            self.refresh_product_tree()
            message = (
                f"Importacao concluida. Lidos: {result['read']} | "
                f"Criados: {result['created']} | Atualizados: {result['updated']} | Ignorados: {result['skipped']}"
            )
            self.db_status_var.set(message)
            self.log_message(message)
            self.show_front_message("showinfo", "Importacao", message)

        ttk.Button(footer, text="Cancelar", style="Secondary.TButton", command=dialog.destroy).pack(side=LEFT)
        ttk.Button(footer, text="Confirmar Importacao", style="Primary.TButton", command=confirm_import).pack(side=LEFT, padx=(8, 0))

    def import_products_batch(self, source_type: str, source_text: str, import_mode: str) -> dict[str, int]:
        preview = self.build_product_import_preview(source_type, source_text, import_mode)
        return self.apply_product_import_plan(preview["plan_rows"])

    def collect_runtime_product_import_sources(self) -> dict[str, list[Path]]:
        xml_sources: list[Path] = []
        xml_sources.extend(parse_selected_paths(self.xml_path_var.get()))
        if self.compare_xml_folder_var.get().strip():
            xml_sources.append(Path(self.compare_xml_folder_var.get().strip()))
        xml_sources.extend(parse_selected_paths(self.sales_consult_xml_paths_var.get()))

        sped_fiscal_sources: list[Path] = []
        for raw_path in (
            self.sped_path_var.get().strip(),
            self.compare_sped_icms_path_var.get().strip(),
        ):
            if raw_path:
                sped_fiscal_sources.append(Path(raw_path))
        sped_fiscal_sources.extend(parse_selected_paths(self.consult_sped_paths_var.get()))
        sped_fiscal_sources.extend(parse_selected_paths(self.sales_consult_sped_paths_var.get()))

        sped_contrib_sources: list[Path] = []
        if self.compare_sped_contrib_path_var.get().strip():
            sped_contrib_sources.append(Path(self.compare_sped_contrib_path_var.get().strip()))

        def unique_existing(paths: list[Path]) -> list[Path]:
            unique_paths: list[Path] = []
            seen_paths: set[str] = set()
            for path in paths:
                normalized = str(path)
                if normalized not in seen_paths and path.exists():
                    seen_paths.add(normalized)
                    unique_paths.append(path)
            return unique_paths

        return {
            "xml": unique_existing(xml_sources),
            "sped_fiscal": unique_existing(sped_fiscal_sources),
            "sped_contrib": unique_existing(sped_contrib_sources),
        }

    def prepare_product_import_rows(
        self,
        source_type: str,
        source_text: str,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[dict[str, object]]:
        self.save_mysql_config(log_success=False)
        self.mysql_repo.ensure_schema()
        if not self.selected_company_id:
            raise ValueError("Nenhuma empresa foi selecionada para a importacao.")

        source_path = Path(source_text)
        normalized_source_type = str(source_type or "").strip()
        if normalized_source_type == "Consolidado":
            runtime_sources = self.collect_runtime_product_import_sources()
            if not runtime_sources["xml"] and not runtime_sources["sped_fiscal"] and not runtime_sources["sped_contrib"]:
                raise ValueError("Nenhuma fonte carregada foi encontrada. Carregue XMLs, SPED Fiscal e/ou SPED PIS/COFINS no sistema.")
            return build_import_products_from_consolidated_sources(
                runtime_sources["xml"],
                runtime_sources["sped_fiscal"],
                runtime_sources["sped_contrib"],
                progress_callback=progress_callback,
            )
        effective_source_type = normalized_source_type
        if source_path.exists():
            if source_path.is_dir():
                effective_source_type = "XML"
            elif source_path.suffix.lower() == ".xml":
                effective_source_type = "XML"
            elif source_path.suffix.lower() in {".txt", ".sped", ".efd"}:
                effective_source_type = "SPED 0200"

        if effective_source_type == "SPED 0200":
            if not source_path.exists():
                raise ValueError("O arquivo SPED informado nao existe.")
            if source_path.is_dir():
                raise ValueError("A fonte SPED 0200 exige um arquivo SPED, nao uma pasta.")
            return build_import_products_from_sped_0200(source_path, progress_callback=progress_callback)
        else:
            xml_sources = parse_selected_paths(source_text)
            if not xml_sources:
                xml_sources = [Path(source_text)]
            xml_sources = [path for path in xml_sources if path.exists()]
            if not xml_sources:
                raise ValueError("Nenhum XML ou pasta de XML valida foi informado.")
            return build_import_products_from_xml_sources(xml_sources, progress_callback=progress_callback)

    def build_product_import_preview(
        self,
        source_type: str,
        source_text: str,
        import_mode: str,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> dict[str, object]:
        import_rows = self.prepare_product_import_rows(source_type, source_text, progress_callback=progress_callback)
        existing_products = self.mysql_repo.list_products(self.selected_company_id)
        by_origin: dict[str, dict[str, object]] = {}
        by_description_ncm: dict[tuple[str, str], dict[str, object]] = {}
        for product in existing_products:
            origin_code = str(product.get("codigo_origem", "")).strip()
            if origin_code:
                by_origin[origin_code] = product
            by_description_ncm[(normalize_text(str(product.get("descricao", ""))), normalize_ncm(str(product.get("ncm", ""))))] = product

        created = 0
        updated = 0
        skipped = 0
        divergence_count = 0
        plan_rows: list[dict[str, object]] = []
        total_rows = max(len(import_rows), 1)
        for index, row in enumerate(import_rows, start=1):
            codigo_origem = str(row.get("codigo_origem", "")).strip()
            descricao = str(row.get("descricao", "")).strip()
            ncm = str(row.get("ncm", "")).strip()
            if not codigo_origem and not descricao:
                plan_rows.append(
                    {
                        "action": "ignored",
                        "reason": "Sem codigo de origem e sem descricao.",
                        "codigo_origem": codigo_origem,
                        "descricao": descricao,
                        "ncm": ncm,
                        "existing_id": "",
                        "existing_codigo": "",
                        "existing_descricao": "",
                        "existing_ncm": "",
                        "divergent": False,
                    }
                )
                skipped += 1
                if progress_callback:
                    progress_callback(index, total_rows, f"Analisando produtos... {index}/{total_rows}")
                continue
            existing = by_origin.get(codigo_origem) if codigo_origem else None
            matched_by = "codigo_origem" if existing is not None else ""
            if existing is None:
                existing = by_description_ncm.get((normalize_text(descricao), normalize_ncm(ncm)))
                if existing is not None:
                    matched_by = "descricao_ncm"

            existing_description = str(existing.get("descricao", "")).strip() if existing else ""
            existing_ncm = str(existing.get("ncm", "")).strip() if existing else ""
            divergent_description = bool(existing and normalize_text(existing_description) != normalize_text(descricao))
            divergent_ncm = bool(existing and normalize_ncm(existing_ncm) != normalize_ncm(ncm))
            divergent = divergent_description or divergent_ncm

            action = ""
            reason = ""
            if existing and import_mode == "Somente criar":
                action = "ignored"
                reason = "Produto ja existe e o modo e Somente criar."
                skipped += 1
            elif existing is None and import_mode == "Somente atualizar":
                action = "ignored"
                reason = "Produto nao existe e o modo e Somente atualizar."
                skipped += 1
            elif existing:
                action = "update"
                reason = f"Produto localizado por {matched_by or 'criterio auxiliar'}."
                updated += 1
            else:
                action = "create"
                reason = "Produto novo para a empresa selecionada."
                created += 1

            if divergent:
                divergence_count += 1
                if divergent_description and divergent_ncm:
                    reason = f"{reason} Divergencia em descricao e NCM."
                elif divergent_description:
                    reason = f"{reason} Divergencia em descricao."
                elif divergent_ncm:
                    reason = f"{reason} Divergencia em NCM."
            import_warnings = [str(item).strip() for item in row.get("import_warnings", []) if str(item).strip()]
            if import_warnings:
                unique_warnings = []
                seen_warnings: set[str] = set()
                for warning in import_warnings:
                    if warning not in seen_warnings:
                        seen_warnings.add(warning)
                        unique_warnings.append(warning)
                reason = f"{reason} {' '.join(unique_warnings)}".strip()

            payload = {
                "codigo_origem": codigo_origem,
                "descricao": descricao,
                "ncm": ncm,
                "unidade": str(row.get("unidade", "UN")).strip() or "UN",
                "tipo_produto": str(existing.get("tipo_produto", "Revenda")).strip() if existing else "Revenda",
                "cst_icms_entrada": normalize_tax_code(row.get("cst_icms_entrada", ""), 3),
                "cst_icms_saida": normalize_tax_code(row.get("cst_icms_saida", ""), 3),
                "icms_entrada": Decimal(row.get("icms_entrada", Decimal("0"))).quantize(Decimal("0.0001")),
                "icms_saida": Decimal(row.get("icms_saida", Decimal("0"))).quantize(Decimal("0.0001")),
                "cst_pis_entrada": normalize_tax_code(row.get("cst_pis_entrada", ""), 2),
                "cst_pis_saida": normalize_tax_code(row.get("cst_pis_saida", ""), 2),
                "pis_entrada": Decimal(row.get("pis_entrada", Decimal("0"))).quantize(Decimal("0.0001")),
                "pis_saida": Decimal(row.get("pis_saida", Decimal("0"))).quantize(Decimal("0.0001")),
                "cst_cofins_entrada": normalize_tax_code(row.get("cst_cofins_entrada", ""), 2),
                "cst_cofins_saida": normalize_tax_code(row.get("cst_cofins_saida", ""), 2),
                "cofins_entrada": Decimal(row.get("cofins_entrada", Decimal("0"))).quantize(Decimal("0.0001")),
                "cofins_saida": Decimal(row.get("cofins_saida", Decimal("0"))).quantize(Decimal("0.0001")),
                "ativo": 1,
            }
            plan_rows.append(
                {
                    "action": action,
                    "reason": reason.strip(),
                    "codigo_origem": codigo_origem,
                    "descricao": descricao,
                    "ncm": ncm,
                    "payload": payload,
                    "existing_id": str(existing.get("id", "")) if existing else "",
                    "existing_codigo": str(existing.get("codigo", "")) if existing else "",
                    "existing_descricao": existing_description,
                    "existing_ncm": existing_ncm,
                    "divergent": divergent,
                }
            )
            if progress_callback:
                progress_callback(index, total_rows, f"Analisando produtos... {index}/{total_rows}")
        return {
            "plan_rows": plan_rows,
            "summary": {"read": len(import_rows), "created": created, "updated": updated, "skipped": skipped, "divergences": divergence_count},
        }

    def apply_product_import_plan(
        self,
        plan_rows: list[dict[str, object]],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> dict[str, int]:
        created = 0
        updated = 0
        skipped = 0
        total_rows = max(len(plan_rows), 1)
        for index, row in enumerate(plan_rows, start=1):
            action = str(row.get("action", "")).strip()
            if action == "create":
                payload = row.get("payload", {})
                if isinstance(payload, dict):
                    self.mysql_repo.save_product(None, self.selected_company_id, payload)
                    created += 1
                else:
                    skipped += 1
            elif action == "update":
                payload = row.get("payload", {})
                existing_id = str(row.get("existing_id", "")).strip()
                if isinstance(payload, dict) and existing_id.isdigit():
                    self.mysql_repo.save_product(int(existing_id), self.selected_company_id, payload)
                    updated += 1
                else:
                    skipped += 1
            else:
                skipped += 1
            if progress_callback:
                progress_callback(index, total_rows, f"Importando produtos... {index}/{total_rows}")
        return {"read": len(plan_rows), "created": created, "updated": updated, "skipped": skipped}

    def save_product_record(self) -> None:
        if not self.selected_company_id:
            messagebox.showwarning("Produto", "Selecione uma empresa antes de cadastrar produtos.")
            return
        descricao = self.product_descricao_var.get().strip()
        if not descricao:
            messagebox.showwarning("Produto", "Informe a descricao do produto.")
            return
        data = {
            "codigo_origem": self.product_codigo_origem_var.get().strip(),
            "descricao": descricao,
            "ncm": self.product_ncm_var.get().strip(),
            "unidade": self.product_unidade_var.get().strip() or "UN",
            "tipo_produto": self.product_tipo_var.get().strip() or "Revenda",
            "cst_icms_entrada": normalize_tax_code(self.product_cst_icms_entrada_var.get(), 3),
            "cst_icms_saida": normalize_tax_code(self.product_cst_icms_saida_var.get(), 3),
            "icms_entrada": parse_decimal(self.product_icms_entrada_var.get()).quantize(Decimal("0.0001")),
            "icms_saida": parse_decimal(self.product_icms_saida_var.get()).quantize(Decimal("0.0001")),
            "cst_pis_entrada": normalize_tax_code(self.product_cst_pis_entrada_var.get(), 2),
            "cst_pis_saida": normalize_tax_code(self.product_cst_pis_saida_var.get(), 2),
            "pis_entrada": parse_decimal(self.product_pis_entrada_var.get()).quantize(Decimal("0.0001")),
            "pis_saida": parse_decimal(self.product_pis_saida_var.get()).quantize(Decimal("0.0001")),
            "cst_cofins_entrada": normalize_tax_code(self.product_cst_cofins_entrada_var.get(), 2),
            "cst_cofins_saida": normalize_tax_code(self.product_cst_cofins_saida_var.get(), 2),
            "cofins_entrada": parse_decimal(self.product_cofins_entrada_var.get()).quantize(Decimal("0.0001")),
            "cofins_saida": parse_decimal(self.product_cofins_saida_var.get()).quantize(Decimal("0.0001")),
            "ativo": 1 if self.product_status_var.get().strip() != "Inativo" else 0,
        }
        is_update = self.selected_product_id is not None
        try:
            self.save_mysql_config(log_success=False)
            self.mysql_repo.ensure_schema()
            saved_id = self.mysql_repo.save_product(self.selected_product_id, self.selected_company_id, data)
            self.selected_product_id = saved_id
            self.product_codigo_var.set(str(saved_id))
            self.product_cst_icms_entrada_var.set(data["cst_icms_entrada"])
            self.product_cst_icms_saida_var.set(data["cst_icms_saida"])
            self.product_cst_pis_entrada_var.set(data["cst_pis_entrada"])
            self.product_cst_pis_saida_var.set(data["cst_pis_saida"])
            self.product_cst_cofins_entrada_var.set(data["cst_cofins_entrada"])
            self.product_cst_cofins_saida_var.set(data["cst_cofins_saida"])
            self.refresh_product_tree()
            if self.product_tree.exists(str(saved_id)):
                self.product_tree.selection_set(str(saved_id))
                self.product_tree.focus(str(saved_id))
            action_text = "alterado" if is_update else "cadastrado"
            message = f"Produto {action_text} com sucesso. ID {saved_id}, empresa {self.selected_company_id}."
            self.db_status_var.set(message)
            self.log_message(message)
            self.show_front_message("showinfo", "Produto", message)
        except Exception as exc:
            self.db_status_var.set(f"Falha ao salvar produto: {exc}")
            self.log_message(f"Falha ao salvar produto: {exc}")
            messagebox.showerror("Produto", str(exc))

    def select_compare_sped_icms_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o arquivo SPED ICMS",
            filetypes=[("Arquivos SPED", "*.txt *.efd *.sped"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.compare_sped_icms_path_var.set(selected)
            self.write_audit_log("ARQUIVO_CARREGADO", f"SPED ICMS comparacao={selected}")

    def select_compare_sped_contrib_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o arquivo SPED Contribuicoes",
            filetypes=[("Arquivos SPED", "*.txt *.efd *.sped"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.compare_sped_contrib_path_var.set(selected)
            self.write_audit_log("ARQUIVO_CARREGADO", f"SPED Contribuicoes comparacao={selected}")

    def select_compare_xml_folder(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta de XML")
        if selected:
            self.compare_xml_folder_var.set(selected)
            self.write_audit_log("ARQUIVO_CARREGADO", f"Pasta XML comparacao={selected}")

    def select_compare_sheet_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm *.xlt *.xltx *.xltm"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.compare_sheet_path_var.set(selected)
            self.write_audit_log("ARQUIVO_CARREGADO", f"Planilha comparacao={selected}")

    def clear_compare_results_for_operation_change(self) -> None:
        if not self.compare_last_results:
            return
        current_scope = normalize_compare_operation_scope(self.compare_operation_scope_var.get())
        if current_scope == self.compare_last_operation_scope:
            return
        self.compare_last_results = []
        self.compare_last_result_origin = ""
        self.compare_last_operation_scope = ""
        self.compare_last_mode = ""
        self.compare_last_stats = {}
        if hasattr(self, "compare_tree"):
            self.compare_tree.delete(*self.compare_tree.get_children())
        self.compare_stats_var.set("Operacao alterada. Execute a comparacao novamente.")
        self.compare_status_var.set("Filtro de operacao alterado; compare novamente para atualizar a grid e a exportacao.")

    def run_compare_invoices(self) -> None:
        compare_mode = self.compare_mode_var.get().strip()
        operation_scope = self.compare_operation_scope_var.get().strip() or "Ambos"
        sped_icms = self.compare_sped_icms_path_var.get().strip()
        sped_contrib = self.compare_sped_contrib_path_var.get().strip()
        target_sped = sped_icms if compare_mode == "icms" else sped_contrib
        xml_folder = self.compare_xml_folder_var.get().strip()
        sheet_file = self.compare_sheet_path_var.get().strip()
        if compare_mode == "icms" and (not sped_icms or not Path(sped_icms).exists()):
            messagebox.showerror("Comparacao SPED", "Selecione um arquivo SPED ICMS valido.")
            return
        if compare_mode == "contrib" and (not sped_contrib or not Path(sped_contrib).exists()):
            messagebox.showerror("Comparacao SPED", "Selecione um arquivo SPED Contribuicoes valido.")
            return
        has_xml = bool(xml_folder and Path(xml_folder).exists())
        has_sheet = bool(sheet_file and Path(sheet_file).exists())
        if not has_xml and not has_sheet:
            messagebox.showerror("Comparacao SPED", "Selecione uma pasta de XML valida ou uma planilha valida.")
            return
        if has_xml and has_sheet:
            messagebox.showerror("Comparacao SPED", "Selecione apenas uma origem por vez: XML ou planilha.")
            return
        if has_sheet:
            sheet_path = Path(sheet_file)
            supported_sheet_suffixes = {".xlsx", ".xls", ".xlsm", ".xlt", ".xltx", ".xltm"}
            if sheet_path.suffix.lower() not in supported_sheet_suffixes:
                messagebox.showerror(
                    "Comparacao SPED",
                    (
                        "A planilha selecionada nao esta em um formato suportado.\n\n"
                        "Use um arquivo .xlsx, .xls, .xlsm, .xlt, .xltx ou .xltm. "
                        "Se o nome terminou como .xlsx.lt, renomeie/salve novamente como .xlsx."
                    ),
                )
                return

        self.compare_run_btn.configure(state="disabled")
        self.update_compare_progress(0, "Iniciando comparacao...")
        source_kind = "sheet" if has_sheet else "xml"
        source_path = Path(sheet_file) if has_sheet else Path(xml_folder)
        self.write_audit_log(
            "COMPARACAO_INICIO",
            (
                f"modo={compare_mode}; operacao={normalize_compare_operation_scope(operation_scope)}; "
                f"sped={target_sped}; origem_tipo={source_kind}; origem={source_path}"
            ),
        )
        threading.Thread(
            target=self.process_compare_invoices_background,
            args=(Path(target_sped), source_path, compare_mode, source_kind, operation_scope),
            daemon=True,
        ).start()

    def update_compare_progress(self, percent: int | float, message: str = "") -> None:
        normalized_percent = max(0, min(100, float(percent)))
        self.compare_progress_var.set(normalized_percent)
        if message:
            self.compare_status_var.set(f"{message} ({normalized_percent:.0f}%)")
        if hasattr(self, "compare_progress_bar"):
            self.compare_progress_bar.update_idletasks()

    def schedule_compare_progress_update(self, current: int, total: int, message: str) -> None:
        percent = int(100 * current / max(total, 1))
        self.root.after(0, lambda percent=percent, message=message: self.update_compare_progress(percent, message))

    def process_compare_invoices_background(
        self,
        sped_file: Path,
        source_path: Path,
        compare_mode: str,
        source_kind: str,
        operation_scope: str,
    ) -> None:
        try:
            if source_kind == "sheet":
                present, missing, stats = compare_sped_with_sheet(
                    sped_file,
                    source_path,
                    operation_scope,
                    self.schedule_compare_progress_update,
                )
                sped_missing_xml: list[CompareSpedDocument] = []
            else:
                present, missing, sped_missing_xml, stats = compare_sped_with_xml_folder(
                    sped_file,
                    source_path,
                    operation_scope,
                    self.schedule_compare_progress_update,
                )
            self.root.after(0, lambda: self.update_compare_progress(95, "Montando resultado na tela."))
            self.root.after(0, lambda: self.render_compare_invoice_results(present, missing, sped_missing_xml, stats, compare_mode, source_kind, operation_scope))
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda message=error_message: self.handle_compare_invoice_error(message))

    def render_compare_invoice_results(
        self,
        present: list[object],
        missing: list[object],
        sped_missing_xml: list[object],
        stats: dict[str, object],
        compare_mode: str,
        source_kind: str,
        operation_scope: str,
    ) -> None:
        self.compare_run_btn.configure(state="normal")
        self.compare_last_results = []
        self.compare_last_result_origin = source_kind
        self.compare_last_operation_scope = normalize_compare_operation_scope(operation_scope)
        self.compare_last_mode = compare_mode
        self.compare_last_stats = stats
        if hasattr(self, "compare_tree"):
            self.compare_tree.delete(*self.compare_tree.get_children())
        self.update_compare_progress(98, "Preenchendo grid de resultado.")

        row_index = 0
        operation_by_key = stats.get("operation_by_key", {})
        if not isinstance(operation_by_key, dict):
            operation_by_key = {}
        source_operation_by_key = stats.get("source_operation_by_key", {})
        if not isinstance(source_operation_by_key, dict):
            source_operation_by_key = {}
        selected_operation = normalize_compare_operation_scope(operation_scope)
        result_groups = [("Lancada", present, "present"), ("Nao lancada", missing, "missing")]
        if source_kind == "xml":
            result_groups.append(("Lancado SPED / Sem XML", sped_missing_xml, "missing"))
            cancelled_invoices = stats.get("cancelled_xml_invoices", [])
            if isinstance(cancelled_invoices, list):
                result_groups.append(("XML cancelado", cancelled_invoices, "cancelled"))
        for status, invoices, tag in result_groups:
            for invoice in invoices:
                operation_type = str(operation_by_key.get(getattr(invoice, "key", ""), "") or "")
                if not operation_type:
                    operation_type = str(source_operation_by_key.get(getattr(invoice, "key", ""), "") or "")
                if not operation_type and selected_operation != "Ambos":
                    operation_type = selected_operation
                self.compare_tree.insert(
                    "",
                    END,
                    iid=str(row_index),
                    values=(
                        COMPARE_MARK_UNCHECKED if status == "Nao lancada" and source_kind == "xml" else "",
                        status,
                        operation_type,
                        getattr(invoice, "model", ""),
                        invoice.key,
                        invoice.number,
                        invoice.series,
                        invoice.issue_date,
                        invoice.issuer_cnpj,
                        invoice.total_value,
                        invoice.file_path,
                    ),
                    tags=(tag,),
                )
                if status != "XML cancelado":
                    self.compare_last_results.append((status, operation_type, invoice))
                row_index += 1
                if row_index % 500 == 0:
                    self.update_compare_progress(98, f"Preenchendo grid: {row_index} registro(s).")

        def summarize_compare_group(label: str, invoices: list[object]) -> str:
            total = sum((self.parse_compare_export_value(getattr(invoice, "total_value", "")) for invoice in invoices), Decimal("0.00"))
            return f"{label}: {len(invoices)} | Valor: {format_decimal_sped(total)}"

        value_summary_parts = [
            summarize_compare_group("Lancadas", present),
            summarize_compare_group("Nao lancadas", missing),
        ]
        if source_kind == "xml":
            value_summary_parts.append(summarize_compare_group("Lancado SPED / Sem XML", sped_missing_xml))
            cancelled_invoices = stats.get("cancelled_xml_invoices", [])
            if isinstance(cancelled_invoices, list) and cancelled_invoices:
                value_summary_parts.append(summarize_compare_group("XML cancelado", cancelled_invoices))
        value_summary_text = "Totais: " + " || ".join(value_summary_parts)

        compare_label = "SPED ICMS" if compare_mode == "icms" else "SPED Contribuicoes"
        operation_label = normalize_compare_operation_scope(operation_scope)
        if source_kind == "sheet":
            number_match_text = f" | Encontradas por numero: {stats.get('match_by_number_total', 0)}" if stats.get("match_by_number_total", 0) else ""
            source_match_text = (
                f" | Linhas casadas: {stats.get('source_present_total', 0)}"
                if stats.get("source_present_total", 0) != stats.get("present_total", 0)
                else ""
            )
            self.compare_stats_var.set(
                f"Operacao: {operation_label} | Docs no {compare_label}: {stats['sped_keys']} | Linhas validas na planilha: {stats['sheet_total']} | "
                f"Linhas ignoradas: {stats['sheet_ignored']}\n"
                f"Docs encontrados no SPED: {stats['present_total']}{source_match_text}{number_match_text} | Nao encontradas: {stats['missing_total']}\n"
                f"{value_summary_text}"
            )
        else:
            number_match_text = f" | Encontradas por numero: {stats.get('match_by_number_total', 0)}" if stats.get("match_by_number_total", 0) else ""
            source_match_text = (
                f" | XMLs casados: {stats.get('source_present_total', 0)}"
                if stats.get("source_present_total", 0) != stats.get("present_total", 0)
                else ""
            )
            self.compare_stats_var.set(
                f"Operacao: {operation_label} | Docs no {compare_label}: {stats['sped_keys']} | XML validos: {stats['xml_total']} | "
                f"XML ignorados: {stats['xml_ignored']} | XML cancelados: {stats.get('xml_cancelled', 0)} | "
                f"Valor cancelado: {format_decimal_sped(Decimal(stats.get('xml_cancelled_value_total', Decimal('0'))))}\n"
                f"Docs encontrados no SPED: {stats['present_total']}{source_match_text}{number_match_text} | "
                f"XML fora do SPED: {stats['missing_total']} | Lancado SPED / Sem XML: {stats.get('sped_missing_xml_total', 0)}\n"
                f"{value_summary_text}"
            )
        if source_kind == "xml" and hasattr(self, "compare_diagnostic_motive_tree"):
            self.refresh_fiscal_diagnostic_tab(show_message=False)
        self.update_compare_progress(100, "Comparacao concluida.")
        self.write_audit_log(
            "COMPARACAO_FIM",
            (
                f"modo={compare_label}; operacao={operation_label}; origem={source_kind}; "
                f"docs_sped={stats.get('sped_keys', 0)}; encontrados={stats.get('present_total', 0)}; "
                f"nao_lancadas_ou_nao_encontradas={stats.get('missing_total', 0)}; "
                f"lancado_sped_sem_xml={stats.get('sped_missing_xml_total', 0)}; {value_summary_text}"
            ),
        )

    def handle_compare_invoice_error(self, message: str) -> None:
        self.compare_run_btn.configure(state="normal")
        self.update_compare_progress(0, "Erro durante o processamento.")
        self.write_audit_log("ERRO_COMPARACAO", message)
        messagebox.showerror("Comparacao SPED", message)

    def parse_compare_export_value(self, value: object) -> Decimal:
        if isinstance(value, Decimal):
            return value.quantize(Decimal("0.01"))
        cleaned = re.sub(r"[^0-9,.\-]", "", str(value or "").strip())
        return parse_decimal(cleaned).quantize(Decimal("0.01"))

    def build_compare_export_rows(self) -> tuple[list[str], list[list[object]], list[str], list[list[object]], Decimal]:
        headers = ["Status", "Operacao", "Modelo", "Chave", "Numero", "Serie", "Data Emissao", "CNPJ Emitente", "Valor", "Arquivo Origem"]
        summary_headers = ["Status", "Operacao", "Quantidade", "Valor"]
        rows: list[list[object]] = []
        summary: dict[tuple[str, str], dict[str, object]] = {}
        total_value = Decimal("0.00")

        for status, operation_type, invoice in self.compare_last_results:
            invoice_value = self.parse_compare_export_value(getattr(invoice, "total_value", ""))
            rows.append(
                [
                    status,
                    operation_type,
                    getattr(invoice, "model", ""),
                    getattr(invoice, "key", ""),
                    getattr(invoice, "number", ""),
                    getattr(invoice, "series", ""),
                    getattr(invoice, "issue_date", ""),
                    getattr(invoice, "issuer_cnpj", ""),
                    invoice_value,
                    getattr(invoice, "file_path", ""),
                ]
            )
            summary_key = (str(status), str(operation_type))
            bucket = summary.setdefault(summary_key, {"quantity": 0, "value": Decimal("0.00")})
            bucket["quantity"] = int(bucket["quantity"]) + 1
            bucket["value"] = Decimal(bucket["value"]) + invoice_value
            total_value += invoice_value

        summary_rows = [
            [status, operation_type, values["quantity"], Decimal(values["value"]).quantize(Decimal("0.01"))]
            for (status, operation_type), values in sorted(summary.items(), key=lambda item: (item[0][0], item[0][1]))
        ]
        return headers, rows, summary_headers, summary_rows, total_value.quantize(Decimal("0.01"))

    def export_compare_invoices_csv(self) -> None:
        if not self.compare_last_results:
            messagebox.showinfo("Exportacao", "Nao ha notas para exportar.")
            return
        selected = filedialog.asksaveasfilename(
            title="Salvar CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="resultado_comparacao_sped.csv",
        )
        if not selected:
            return
        headers, rows, summary_headers, summary_rows, total_value = self.build_compare_export_rows()
        total_row = ["Total", "", "", "", "", "", "", "", total_value, ""]
        summary_total_row = ["Total Geral", "", len(rows), total_value]
        try:
            output_path = Path(selected)
            summary_output = output_path.with_name(f"{output_path.stem}_totais.csv")
            write_simple_csv_file(output_path, headers, rows + [total_row])
            write_simple_csv_file(summary_output, summary_headers, summary_rows + [summary_total_row])
            messagebox.showinfo("Exportacao", f"Arquivos salvos em:\n{output_path}\n{summary_output}")
        except Exception as exc:
            messagebox.showerror("Exportacao", str(exc))

    def export_compare_invoices_excel(self) -> None:
        if not self.compare_last_results:
            messagebox.showinfo("Exportacao", "Nao ha notas para exportar.")
            return
        selected = filedialog.asksaveasfilename(
            title="Salvar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="resultado_comparacao_sped.xlsx",
        )
        if not selected:
            return
        headers, rows, summary_headers, summary_rows, _total_value = self.build_compare_export_rows()
        try:
            write_simple_excel_workbook(
                Path(selected),
                [
                    ("Comparacao", headers, rows),
                    ("Totais", summary_headers, summary_rows),
                ],
            )
            messagebox.showinfo("Exportacao", f"Arquivo Excel salvo em:\n{selected}")
        except Exception as exc:
            messagebox.showerror("Exportacao", str(exc))

    def export_compare_ignored_xmls(self) -> None:
        xml_folder_text = self.compare_xml_folder_var.get().strip()
        if not xml_folder_text or not Path(xml_folder_text).exists():
            messagebox.showinfo("XMLs ignorados", "Selecione uma pasta de XML valida.")
            return
        rows = collect_compare_ignored_xml_rows(Path(xml_folder_text))
        if not rows:
            messagebox.showinfo("XMLs ignorados", "Nao ha XMLs ignorados na pasta selecionada.")
            return
        selected = filedialog.asksaveasfilename(
            title="Exportar XMLs ignorados",
            defaultextension=".xlsx",
            initialfile="xmls_ignorados_comparacao.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not selected:
            return
        output_path = Path(selected)
        headers = ["Arquivo", "Motivo", "Caminho"]
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                write_simple_excel_workbook(output_path, [("XMLs Ignorados", headers, rows)])
            messagebox.showinfo("XMLs ignorados", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("XMLs ignorados", str(exc))

    def build_compare_investigation_dataset(self) -> dict[str, object]:
        if self.compare_last_result_origin != "xml" or not self.compare_last_results:
            raise ValueError("Execute uma comparacao a partir de XML antes de investigar a diferenca.")
        stats = dict(getattr(self, "compare_last_stats", {}) or {})
        fiscal_basis = dict(stats.get("fiscal_basis", {}) or {})
        matched_rows = [
            row
            for row in stats.get("matched_rows", [])
            if isinstance(row, CompareInvestigationMatchedRow)
        ]
        xml_sped_difference_rows: list[list[object]] = []
        xml_sped_difference_total = Decimal("0.00")
        for row in matched_rows:
            difference = compare_decimal_value(row.difference)
            if difference == 0:
                continue
            xml_sped_difference_total += difference
            xml_sped_difference_rows.append(
                [
                    row.key,
                    row.number,
                    row.series,
                    row.model,
                    row.issuer_cnpj,
                    row.issue_date,
                    compare_decimal_value(row.xml_value),
                    compare_decimal_value(row.sped_value),
                    difference,
                    row.match_type,
                    row.reason,
                    row.xml_file,
                    row.sped_key,
                ]
            )
        xml_sped_difference_rows.sort(key=lambda values: abs(Decimal(values[8])), reverse=True)
        xml_value_by_key = {
            row.key: compare_decimal_value(row.xml_value)
            for row in matched_rows
            if str(row.key).strip()
        }

        xml_missing_rows: list[list[object]] = []
        sped_missing_rows: list[list[object]] = []
        present_xml_total = Decimal("0.00")
        missing_xml_total = Decimal("0.00")
        sped_missing_total = Decimal("0.00")
        for status, operation_type, invoice in self.compare_last_results:
            value = compare_decimal_value(getattr(invoice, "total_value", ""))
            if str(status).startswith("Lancada"):
                present_xml_total += value
            elif status == "Nao lancada":
                missing_xml_total += value
                xml_missing_rows.append(
                    [
                        status,
                        operation_type,
                        getattr(invoice, "model", ""),
                        getattr(invoice, "key", ""),
                        getattr(invoice, "number", ""),
                        getattr(invoice, "series", ""),
                        getattr(invoice, "issue_date", ""),
                        getattr(invoice, "issuer_cnpj", ""),
                        value,
                        getattr(invoice, "file_path", ""),
                    ]
                )
            elif status == "Lancado SPED / Sem XML":
                sped_missing_total += value
                sped_missing_rows.append(
                    [
                        status,
                        operation_type,
                        getattr(invoice, "model", ""),
                        getattr(invoice, "key", ""),
                        getattr(invoice, "number", ""),
                        getattr(invoice, "series", ""),
                        getattr(invoice, "issue_date", ""),
                        getattr(invoice, "issuer_cnpj", ""),
                        value,
                        "Servico gerado/CFOP 933" if getattr(invoice, "generated_nfse_service", False) else "",
                    ]
                )

        comparison_total = (present_xml_total + sped_missing_total).quantize(Decimal("0.01"))
        present_sped_total = Decimal(stats.get("present_sped_value_total", Decimal("0.00"))).quantize(Decimal("0.01"))
        document_total = Decimal(fiscal_basis.get("document_total", present_sped_total + sped_missing_total)).quantize(Decimal("0.01"))
        fiscal_total = Decimal(fiscal_basis.get("fiscal_total", document_total)).quantize(Decimal("0.01"))
        xml_sped_difference_total = (present_sped_total - present_xml_total).quantize(Decimal("0.01"))
        document_difference = (document_total - (present_sped_total + sped_missing_total)).quantize(Decimal("0.01"))
        fiscal_document_difference = (fiscal_total - document_total).quantize(Decimal("0.01"))
        total_difference = (fiscal_total - comparison_total).quantize(Decimal("0.01"))

        summary_rows = [
            ["Total XML lancadas", present_xml_total, "Valor das notas encontradas usando o XML."],
            ["Lancado SPED / Sem XML", sped_missing_total, "Valor dos documentos que ja estao lancados no SPED, mas nao tiveram XML localizado na pasta."],
            ["Total da comparacao", comparison_total, "Soma usada no rodape da comparacao."],
            ["Diferenca XML x SPED nas lancadas", xml_sped_difference_total, "Mesmo documento, mas valor do XML diferente do valor no SPED."],
            ["Total lancadas pelo valor SPED", present_sped_total, "Total dos documentos casados usando o valor do SPED."],
            ["Diferenca documentos nao listados", document_difference, "Diferenca entre os documentos SPED coletados e a soma dos grupos da comparacao."],
            ["Total documentos C100/A100", document_total, "Total dos documentos fiscais coletados no SPED."],
            ["Ajuste registros fiscais", fiscal_document_difference, "Diferenca entre os documentos e os registros de resumo fiscal."],
            ["Total fiscal da tela", fiscal_total, "Soma dos resumos fiscais C190/C590/D590."],
            ["Diferenca total explicada", total_difference, "Total fiscal menos total da comparacao."],
        ]
        fiscal_rows = [
            ["C100 documentos", fiscal_basis.get("c100_count", 0), Decimal(fiscal_basis.get("c100_total", Decimal("0.00")))],
            ["A100 documentos", fiscal_basis.get("a100_count", 0), Decimal(fiscal_basis.get("a100_total", Decimal("0.00")))],
            ["C190 resumo fiscal", fiscal_basis.get("c190_count", 0), Decimal(fiscal_basis.get("c190_total", Decimal("0.00")))],
            ["C590 resumo fiscal", fiscal_basis.get("c590_count", 0), Decimal(fiscal_basis.get("c590_total", Decimal("0.00")))],
            ["D590 resumo fiscal", fiscal_basis.get("d590_count", 0), Decimal(fiscal_basis.get("d590_total", Decimal("0.00")))],
            ["Total documentos C100/A100", "", document_total],
            ["Total fiscal C190/C590/D590", "", fiscal_total],
            ["Diferenca fiscal x documentos", "", fiscal_document_difference],
        ]
        fiscal_difference_rows: list[list[object]] = []
        c100_c190_rows = list(fiscal_basis.get("c100_c190_rows", []) or [])
        c500_c590_rows = list(fiscal_basis.get("c500_c590_rows", []) or [])
        d500_d590_rows = list(fiscal_basis.get("d500_d590_rows", []) or [])
        c100_c190_total = sum((Decimal(row[8]) for row in c100_c190_rows), Decimal("0.00")).quantize(Decimal("0.01"))
        c100_c190_positive_total = sum((Decimal(row[8]) for row in c100_c190_rows if Decimal(row[8]) > 0), Decimal("0.00")).quantize(Decimal("0.01"))
        c100_c190_negative_total = sum((Decimal(row[8]) for row in c100_c190_rows if Decimal(row[8]) < 0), Decimal("0.00")).quantize(Decimal("0.01"))
        c500_c590_total = sum((Decimal(row[5]) for row in c500_c590_rows), Decimal("0.00")).quantize(Decimal("0.01"))
        d500_d590_total = sum((Decimal(row[5]) for row in d500_d590_rows), Decimal("0.00")).quantize(Decimal("0.01"))
        xml_sped_ipi_total = Decimal("0.00")
        xml_sped_st_total = Decimal("0.00")
        xml_sped_other_total = Decimal("0.00")
        for row in xml_sped_difference_rows:
            row_difference = Decimal(row[8])
            row_reason = str(row[10]).lower() if len(row) > 10 else ""
            if "ipi" in row_reason:
                xml_sped_ipi_total += row_difference
            elif "st" in row_reason:
                xml_sped_st_total += row_difference
            else:
                xml_sped_other_total += row_difference
        simple_diagnostic_rows = [
            [
                "Valor XML diferente do valor do documento no SPED",
                xml_sped_difference_total,
                len(xml_sped_difference_rows),
                "O mesmo documento existe no XML e no SPED, mas o valor total nao esta igual.",
                "Ver aba XML x SPED diferente.",
            ],
            [
                "Dentro do SPED: resumo C190 diferente do documento C100",
                c100_c190_total,
                len(c100_c190_rows),
                "O C100 mostra o valor do documento, mas o C190 soma bases por CST/CFOP. A diferenca normalmente vem de desconto, ST, IPI ou outros acrescimos.",
                "Ver abas Ajuste fiscal resumo e Ajuste fiscal detalhado.",
            ],
            [
                "Registros C500/C590 entram no total fiscal",
                c500_c590_total,
                len(c500_c590_rows),
                "Esses valores fazem parte do total fiscal da tela, mas nao fazem parte da comparacao principal de XML C100.",
                "Ver aba Ajuste fiscal detalhado filtrando C500/C590.",
            ],
            [
                "Registros D500/D590 entram no total fiscal",
                d500_d590_total,
                len(d500_d590_rows),
                "Esses valores fazem parte do total fiscal da tela, mas nao fazem parte da comparacao principal de XML C100.",
                "Ver aba Ajuste fiscal detalhado filtrando D500/D590.",
            ],
            [
                "Diferenca total explicada",
                total_difference,
                len(xml_sped_difference_rows) + len(c100_c190_rows) + len(c500_c590_rows) + len(d500_d590_rows),
                "Soma dos motivos acima: valor XML x SPED + ajuste fiscal interno do SPED.",
                "O total desta linha deve bater com a diferenca entre comparacao e total fiscal.",
            ],
        ]
        simple_composition_rows = [
            ["XML x SPED - provavel IPI", xml_sped_ipi_total, "SPED menor que XML; diferenca bate com IPI do XML nao levado ao VL_DOC."],
            ["XML x SPED - provavel ST", xml_sped_st_total, "SPED maior que XML; indicio de ST no SPED sem o mesmo valor no XML."],
            ["XML x SPED - outros", xml_sped_other_total, "Diferencas que precisam de conferencia de composicao."],
            ["C100 x C190 - descontos", c100_c190_positive_total, "C190 maior que C100; normalmente o resumo esta antes do desconto do documento."],
            ["C100 x C190 - acrescimos/impostos", c100_c190_negative_total, "C190 menor que C100; normalmente o documento inclui ST, IPI, outros ou acrescimos que nao entram igual no resumo."],
            ["C100 x C190", c100_c190_total, "Diferenca interna entre documento e resumo fiscal no SPED."],
            ["C500/C590", c500_c590_total, "Outros documentos fiscais que entram no total fiscal."],
            ["D500/D590", d500_d590_total, "Outros documentos fiscais que entram no total fiscal."],
        ]
        explained_difference_total = (xml_sped_difference_total + c100_c190_total + c500_c590_total + d500_d590_total).quantize(Decimal("0.01"))
        pending_difference_total = (total_difference - explained_difference_total).quantize(Decimal("0.01"))
        reconciliation_status = "CONCILIADO" if pending_difference_total == 0 else "PENDENTE"
        reconciliation_check_rows = [
            ["Diferenca original", total_difference, "Total fiscal da tela menos total da comparacao."],
            ["Diferenca explicada", explained_difference_total, "Soma dos motivos identificados pela ferramenta."],
            ["Diferenca pendente", pending_difference_total, "Valor ainda sem explicacao. Deve ser zero para considerar conciliado."],
            ["Status", reconciliation_status, "CONCILIADO quando a diferenca pendente for 0,00."],
        ]
        simple_explanation_text = (
            f"Sua diferenca total e {format_decimal_sped(total_difference)}. "
            f"Ela vem de {format_decimal_sped(xml_sped_difference_total)} porque existem {len(xml_sped_difference_rows)} nota(s) em que o valor do XML nao bate com o valor do documento no SPED; "
            f"dentro desse grupo, {format_decimal_sped(xml_sped_st_total)} tem indicio de ST no SPED e {format_decimal_sped(xml_sped_ipi_total)} tem indicio de IPI do XML nao levado ao VL_DOC. "
            f"Depois, dentro do proprio SPED, o resumo fiscal C190 difere do C100 em {format_decimal_sped(c100_c190_total)} "
            f"({format_decimal_sped(c100_c190_positive_total)} por descontos/resumo maior que documento e {format_decimal_sped(c100_c190_negative_total)} por acrescimos/impostos/resumo menor que documento). "
            f"Por fim, o total fiscal tambem inclui {format_decimal_sped(c500_c590_total)} de C500/C590 e {format_decimal_sped(d500_d590_total)} de D500/D590. "
            f"Somando esses motivos, a diferenca explicada e {format_decimal_sped(explained_difference_total)} "
            f"e a diferenca pendente e {format_decimal_sped(pending_difference_total)}. Status: {reconciliation_status}."
        )
        fiscal_component_rows = [
            [
                "1",
                "C100 x C190",
                "Soma de cada C190 dentro do documento menos o VL_DOC do C100.",
                "C190 VL_OPR - C100 VL_DOC",
                c100_c190_total,
                len(c100_c190_rows),
                "Abrir aba Ajuste fiscal detalhado e filtrar Origem = C100 x C190.",
            ],
            [
                "2",
                "C500/C590",
                "Documentos de energia/comunicacao que entram no total fiscal pelos resumos C590.",
                "C590 VL_OPR",
                c500_c590_total,
                len(c500_c590_rows),
                "Abrir aba Ajuste fiscal detalhado e filtrar Origem = C500/C590.",
            ],
            [
                "3",
                "D500/D590",
                "Documentos de transporte/comunicacao que entram no total fiscal pelos resumos D590.",
                "D590 VL_OPR",
                d500_d590_total,
                len(d500_d590_rows),
                "Abrir aba Ajuste fiscal detalhado e filtrar Origem = D500/D590.",
            ],
            [
                "",
                "Total ajuste fiscal",
                "Total que transforma documentos C100/A100 no total fiscal da tela.",
                "-",
                fiscal_document_difference,
                len(c100_c190_rows) + len(c500_c590_rows) + len(d500_d590_rows),
                "Este valor precisa bater com a soma da aba Ajuste fiscal detalhado.",
            ],
        ]
        reconciliation_rows = [
            ["1", "Notas lancadas pelo valor do XML", present_xml_total, "Ponto de partida da comparacao."],
            ["2", "Lancado SPED / Sem XML", sped_missing_total, "Somado porque ja existe no SPED, mas nao existe XML correspondente na pasta."],
            ["3", "Total da comparacao", comparison_total, "1 + 2."],
            ["4", "Troca valor XML pelo valor SPED nas notas lancadas", xml_sped_difference_total, "Diferenca detalhada na aba XML x SPED diferente."],
            ["5", "Total documentos comparaveis pelo SPED", (comparison_total + xml_sped_difference_total).quantize(Decimal("0.01")), "3 + 4."],
            ["6", "Diferenca documentos nao listados", document_difference, "Normalmente zero. Se houver valor, falta classificar documento na conciliacao."],
            ["7", "Total documentos C100/A100", document_total, "5 + 6."],
            ["8", "C100 x C190", c100_c190_total, "Parte do ajuste fiscal detalhada por documento."],
            ["9", "C500/C590", c500_c590_total, "Parte do ajuste fiscal detalhada por documento."],
            ["10", "D500/D590", d500_d590_total, "Parte do ajuste fiscal detalhada por documento."],
            ["11", "Total fiscal da tela", fiscal_total, "7 + 8 + 9 + 10."],
            ["12", "Diferenca total entre comparacao e tela fiscal", total_difference, "11 - 3."],
        ]
        for row in c100_c190_rows:
            difference = Decimal(row[8])
            xml_value = xml_value_by_key.get(str(row[2]), "")
            line_label = f"L{row[9]}" if str(row[9]).strip() else ""
            vl_merc = Decimal(row[11]) if len(row) > 11 else Decimal("0.00")
            vl_desc = Decimal(row[12]) if len(row) > 12 else Decimal("0.00")
            vl_frt = Decimal(row[13]) if len(row) > 13 else Decimal("0.00")
            vl_seg = Decimal(row[14]) if len(row) > 14 else Decimal("0.00")
            vl_out = Decimal(row[15]) if len(row) > 15 else Decimal("0.00")
            vl_st = Decimal(row[16]) if len(row) > 16 else Decimal("0.00")
            vl_ipi = Decimal(row[17]) if len(row) > 17 else Decimal("0.00")
            additions = (vl_frt + vl_seg + vl_out + vl_st + vl_ipi).quantize(Decimal("0.01"))
            expected_difference = (vl_desc - additions).quantize(Decimal("0.01"))
            if abs(difference - vl_desc) <= Decimal("0.01") and vl_desc:
                reason = f"Diferenca bate com desconto do C100: C190 {format_decimal_sped(Decimal(row[7]))} - VL_DOC {format_decimal_sped(Decimal(row[6]))} = desconto {format_decimal_sped(vl_desc)}."
            elif abs(difference - expected_difference) <= Decimal("0.01"):
                reason = (
                    f"Diferenca bate com desconto menos acrescimos: desconto {format_decimal_sped(vl_desc)} - "
                    f"(frete {format_decimal_sped(vl_frt)} + seguro {format_decimal_sped(vl_seg)} + outros {format_decimal_sped(vl_out)} + "
                    f"ST {format_decimal_sped(vl_st)} + IPI {format_decimal_sped(vl_ipi)}) = {format_decimal_sped(expected_difference)}."
                )
            elif difference < 0:
                reason = "C190 menor que C100. Conferir acrescimos/impostos no C100 que nao compoem o VL_OPR do C190."
            else:
                reason = "C190 maior que C100. Conferir desconto/abatimento no C100 ou resumo C190 somando produtos antes do desconto."
            fiscal_difference_rows.append(
                [
                    "C100 x C190",
                    *row[:6],
                    xml_value,
                    row[6],
                    row[7],
                    row[8],
                    line_label,
                    row[10],
                    vl_merc,
                    vl_desc,
                    vl_frt,
                    vl_seg,
                    vl_out,
                    vl_st,
                    vl_ipi,
                    expected_difference,
                    reason,
                ]
            )
        for row in c500_c590_rows:
            fiscal_difference_rows.append(["C500/C590", row[0], row[1], "", row[2], row[3], row[4], "", Decimal("0.00"), row[5], row[5], f"L{row[6]}", row[7], Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), row[5], "Valor vem de documento C500 resumido em C590; entra no total fiscal, mas nao entra na comparacao de XML C100."])
        for row in d500_d590_rows:
            fiscal_difference_rows.append(["D500/D590", row[0], row[1], "", row[2], row[3], row[4], "", Decimal("0.00"), row[5], row[5], f"L{row[6]}", row[7], Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), row[5], "Valor vem de documento D500 resumido em D590; entra no total fiscal, mas nao entra na comparacao de XML C100."])
        fiscal_difference_rows.sort(
            key=lambda values: abs(Decimal(values[9] if len(values) > 9 else Decimal("0.00"))),
            reverse=True,
        )
        return {
            "summary_rows": summary_rows,
            "xml_sped_difference_rows": xml_sped_difference_rows,
            "sped_missing_rows": sped_missing_rows,
            "xml_missing_rows": xml_missing_rows,
            "fiscal_rows": fiscal_rows,
            "fiscal_component_rows": fiscal_component_rows,
            "reconciliation_rows": reconciliation_rows,
            "simple_diagnostic_rows": simple_diagnostic_rows,
            "simple_composition_rows": simple_composition_rows,
            "reconciliation_check_rows": reconciliation_check_rows,
            "explained_difference_total": explained_difference_total,
            "pending_difference_total": pending_difference_total,
            "reconciliation_status": reconciliation_status,
            "simple_explanation_text": simple_explanation_text,
            "fiscal_difference_rows": fiscal_difference_rows,
            "comparison_total": comparison_total,
            "fiscal_total": fiscal_total,
            "total_difference": total_difference,
        }

    def open_compare_difference_investigation(self) -> None:
        try:
            dataset = self.build_compare_investigation_dataset()
        except Exception as exc:
            messagebox.showinfo("Investigacao", str(exc))
            return
        popup = Toplevel(self.root)
        popup.title("Investigacao da diferenca XML x SPED")
        popup.geometry("1180x720")
        popup.minsize(980, 560)
        popup.transient(self.root)

        container = ttk.Frame(popup, padding=10)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        header_text = (
            f"Total da comparacao: {format_decimal_sped(Decimal(dataset['comparison_total']))}   |   "
            f"Total fiscal da tela: {format_decimal_sped(Decimal(dataset['fiscal_total']))}   |   "
            f"Diferenca: {format_decimal_sped(Decimal(dataset['total_difference']))}"
        )
        ttk.Label(container, text=header_text, style="Strong.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew")

        simple_tab = ttk.Frame(notebook, padding=10)
        simple_tab.columnconfigure(0, weight=1)
        simple_tab.rowconfigure(3, weight=1)
        simple_tab.rowconfigure(4, weight=1)
        simple_tab.rowconfigure(5, weight=1)
        simple_tab.rowconfigure(6, weight=0)
        notebook.add(simple_tab, text="Diagnostico simples")
        ttk.Label(
            simple_tab,
            text=(
                f"Sua diferenca e {format_decimal_sped(Decimal(dataset['total_difference']))}. "
                "Ela foi explicada pelos motivos abaixo."
            ),
            style="Strong.TLabel",
            wraplength=1080,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))
        conclusion_box = ttk.LabelFrame(simple_tab, text="Conclusao para o usuario", padding=8)
        conclusion_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        conclusion_box.columnconfigure(0, weight=1)
        conclusion_box.columnconfigure(1, weight=0)
        conclusion_text = str(dataset.get("simple_explanation_text", ""))
        ttk.Label(
            conclusion_box,
            text=conclusion_text,
            wraplength=1080,
            justify=LEFT,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            conclusion_box,
            text="Copiar conclusao",
            style="Secondary.TButton",
            command=lambda text=conclusion_text: self.copy_text_to_clipboard(text, "Conclusao copiada para a area de transferencia."),
        ).grid(row=0, column=1, sticky="ne")
        ttk.Label(
            simple_tab,
            text=(
                "Leia esta tela de cima para baixo. Se precisar conferir uma nota ou linha do SPED, use a coluna Onde conferir."
            ),
            wraplength=1080,
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))

        add_tree_frame = ttk.Frame(simple_tab)
        add_tree_frame.grid(row=3, column=0, sticky="nsew")
        add_tree_frame.columnconfigure(0, weight=1)
        add_tree_frame.rowconfigure(0, weight=1)

        def build_investigation_footer_text(headers: list[str], rows: list[list[object]]) -> str:
            parts = [f"Linhas: {len(rows)}"]
            for column_index, header in enumerate(headers):
                total = Decimal("0.00")
                has_decimal = False
                for row in rows:
                    if column_index >= len(row):
                        continue
                    value = row[column_index]
                    if isinstance(value, Decimal):
                        total += value
                        has_decimal = True
                if has_decimal:
                    parts.append(f"{header}: {format_decimal_sped(total.quantize(Decimal('0.01')))}")
            return "   |   ".join(parts)

        def add_tree_tab(title: str, headers: list[str], rows: list[list[object]], widths: list[int] | None = None) -> ttk.Treeview:
            tab = ttk.Frame(notebook, padding=6)
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)
            notebook.add(tab, text=title)
            columns = [f"c{index}" for index in range(len(headers))]
            tree = ttk.Treeview(tab, columns=columns, show="headings", height=14, selectmode="extended")
            for index, column_id in enumerate(columns):
                tree.heading(column_id, text=headers[index])
                tree.column(column_id, width=(widths[index] if widths and index < len(widths) else 130), anchor="center")
            for row in rows:
                tree.insert("", END, values=[format_decimal_sped(value) if isinstance(value, Decimal) else value for value in row])
            tree.grid(row=0, column=0, sticky="nsew")
            scroll_y = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            tree.bind("<Control-c>", lambda event, current_tree=tree: self.copy_treeview_selection(current_tree))
            tree.bind("<Button-3>", lambda event, current_tree=tree: self.show_treeview_copy_menu(event, current_tree))
            ttk.Label(
                tab,
                text=build_investigation_footer_text(headers, rows),
                style="Strong.TLabel",
                wraplength=1120,
            ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
            return tree

        def fill_tree(parent_frame: ttk.Frame, headers: list[str], rows: list[list[object]], widths: list[int]) -> ttk.Treeview:
            columns = [f"c{index}" for index in range(len(headers))]
            tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=7, selectmode="extended")
            for index, column_id in enumerate(columns):
                tree.heading(column_id, text=headers[index])
                tree.column(column_id, width=widths[index] if index < len(widths) else 130, anchor="center")
            for row in rows:
                tree.insert("", END, values=[format_decimal_sped(value) if isinstance(value, Decimal) else value for value in row])
            tree.grid(row=0, column=0, sticky="nsew")
            scroll_y = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x = ttk.Scrollbar(parent_frame, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            tree.bind("<Control-c>", lambda event, current_tree=tree: self.copy_treeview_selection(current_tree))
            tree.bind("<Button-3>", lambda event, current_tree=tree: self.show_treeview_copy_menu(event, current_tree))
            ttk.Label(
                parent_frame,
                text=build_investigation_footer_text(headers, rows),
                style="Strong.TLabel",
                wraplength=1120,
            ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
            return tree

        fill_tree(
            add_tree_frame,
            ["Motivo", "Valor", "Qtd", "Explicacao", "Onde conferir"],
            dataset["simple_diagnostic_rows"],
            [320, 130, 70, 430, 330],
        )

        check_frame = ttk.LabelFrame(simple_tab, text="Verificacao da conciliacao", padding=6)
        check_frame.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        check_frame.columnconfigure(0, weight=1)
        check_frame.rowconfigure(0, weight=1)
        fill_tree(
            check_frame,
            ["Item", "Valor/Status", "Leitura"],
            dataset["reconciliation_check_rows"],
            [260, 150, 760],
        )

        composition_frame = ttk.LabelFrame(simple_tab, text="Composicao resumida", padding=6)
        composition_frame.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        composition_frame.columnconfigure(0, weight=1)
        composition_frame.rowconfigure(0, weight=1)
        fill_tree(
            composition_frame,
            ["Componente", "Valor", "Leitura"],
            dataset["simple_composition_rows"],
            [260, 130, 760],
        )

        simple_footer = ttk.Frame(simple_tab)
        simple_footer.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        explained_total = sum(
            (
                Decimal(row[1])
                for row in dataset["simple_composition_rows"]
                if len(row) > 1 and isinstance(row[1], Decimal)
            ),
            Decimal("0.00"),
        ).quantize(Decimal("0.01"))
        footer_items = [
            ("Total comparacao", Decimal(dataset["comparison_total"])),
            ("Total fiscal", Decimal(dataset["fiscal_total"])),
            ("Diferenca", Decimal(dataset["total_difference"])),
            ("Diferenca explicada", Decimal(dataset.get("explained_difference_total", explained_total))),
            ("Diferenca pendente", Decimal(dataset.get("pending_difference_total", Decimal("0.00")))),
        ]
        for label, value in footer_items:
            ttk.Label(
                simple_footer,
                text=f"{label}: {format_decimal_sped(value)}",
                style="Strong.TLabel",
            ).pack(side=LEFT, padx=(0, 18))

        add_tree_tab(
            "Fechamento",
            ["Passo", "Conta", "Valor", "Como conferir"],
            dataset["reconciliation_rows"],
            [70, 330, 130, 620],
        )
        add_tree_tab(
            "Resumo tecnico",
            ["Etapa", "Valor", "Explicacao"],
            dataset["summary_rows"],
            [260, 130, 620],
        )
        add_tree_tab(
            "XML x SPED diferente",
            ["Chave", "Numero", "Serie", "Modelo", "CNPJ Emitente", "Data", "Valor XML", "Valor SPED", "Diferenca", "Casamento", "Motivo provavel", "Arquivo XML", "Chave SPED"],
            dataset["xml_sped_difference_rows"],
            [270, 90, 70, 70, 140, 100, 100, 100, 100, 95, 420, 420, 270],
        )
        add_tree_tab(
            "Lancado SPED / Sem XML",
            ["Status", "Operacao", "Modelo", "Chave/Id", "Numero", "Serie", "Data", "CNPJ Participante", "Valor SPED", "Observacao"],
            dataset["sped_missing_rows"],
            [140, 90, 70, 270, 90, 70, 100, 150, 110, 220],
        )
        add_tree_tab(
            "XML nao lancado",
            ["Status", "Operacao", "Modelo", "Chave", "Numero", "Serie", "Data", "CNPJ Emitente", "Valor XML", "Arquivo XML"],
            dataset["xml_missing_rows"],
            [140, 90, 70, 270, 90, 70, 100, 150, 110, 460],
        )
        add_tree_tab(
            "Bases fiscais",
            ["Base", "Linhas", "Valor"],
            dataset["fiscal_rows"],
            [260, 90, 130],
        )
        add_tree_tab(
            "Ajuste fiscal resumo",
            ["Ordem", "Origem", "O que significa", "Formula", "Valor", "Linhas", "Onde conferir"],
            dataset["fiscal_component_rows"],
            [70, 140, 360, 180, 130, 80, 420],
        )
        add_tree_tab(
            "Ajuste fiscal detalhado",
            [
                "Origem", "Operacao", "Modelo", "Chave/Id", "Numero", "Serie", "Data", "Valor XML", "Valor Doc SPED C100",
                "Valor Resumo SPED", "Diferenca Resumo - Doc", "Linha Doc SPED", "Linhas Resumo SPED", "VL_MERC C100", "VL_DESC C100", "VL_FRT C100", "VL_SEG C100",
                "VL_OUT", "VL_ST", "VL_IPI", "Formula", "Motivo provavel",
            ],
            dataset["fiscal_difference_rows"],
            [120, 90, 70, 270, 90, 70, 100, 110, 140, 140, 150, 110, 520, 110, 110, 110, 110, 100, 100, 100, 110, 620],
        )

        footer = ttk.Frame(container)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            footer,
            text="A diferenca nao significa necessariamente erro: a comparacao usa documentos/XML, enquanto o total fiscal da consulta usa registros de resumo do SPED.",
            wraplength=820,
        ).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(footer, text="Exportar investigacao", style="Secondary.TButton", command=lambda: self.export_compare_investigation_dataset(dataset)).pack(side=RIGHT, padx=(8, 0))
        ttk.Button(footer, text="Fechar", style="Secondary.TButton", command=popup.destroy).pack(side=RIGHT)

    def export_compare_investigation_dataset(self, dataset: dict[str, object] | None = None) -> None:
        try:
            current_dataset = dataset or self.build_compare_investigation_dataset()
        except Exception as exc:
            messagebox.showinfo("Investigacao", str(exc))
            return
        selected = filedialog.asksaveasfilename(
            title="Exportar investigacao",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile="investigacao_diferenca_sped.xlsx",
        )
        if not selected:
            return
        try:
            write_simple_excel_workbook(
                Path(selected),
                [
                    ("Diagnostico simples", ["Motivo", "Valor", "Qtd", "Explicacao", "Onde conferir"], current_dataset["simple_diagnostic_rows"]),
                    ("Conclusao", ["Texto"], [[current_dataset.get("simple_explanation_text", "")]]),
                    ("Verificacao", ["Item", "Valor/Status", "Leitura"], current_dataset["reconciliation_check_rows"]),
                    ("Composicao resumida", ["Componente", "Valor", "Leitura"], current_dataset["simple_composition_rows"]),
                    ("Fechamento", ["Passo", "Conta", "Valor", "Como conferir"], current_dataset["reconciliation_rows"]),
                    ("Resumo tecnico", ["Etapa", "Valor", "Explicacao"], current_dataset["summary_rows"]),
                    (
                        "XML x SPED diferente",
                        ["Chave", "Numero", "Serie", "Modelo", "CNPJ Emitente", "Data", "Valor XML", "Valor SPED", "Diferenca", "Casamento", "Motivo provavel", "Arquivo XML", "Chave SPED"],
                        current_dataset["xml_sped_difference_rows"],
                    ),
                    (
                        "Lancado SPED Sem XML",
                        ["Status", "Operacao", "Modelo", "Chave/Id", "Numero", "Serie", "Data", "CNPJ Participante", "Valor SPED", "Observacao"],
                        current_dataset["sped_missing_rows"],
                    ),
                    (
                        "XML nao lancado",
                        ["Status", "Operacao", "Modelo", "Chave", "Numero", "Serie", "Data", "CNPJ Emitente", "Valor XML", "Arquivo XML"],
                        current_dataset["xml_missing_rows"],
                    ),
                    ("Bases fiscais", ["Base", "Linhas", "Valor"], current_dataset["fiscal_rows"]),
                    (
                        "Ajuste fiscal resumo",
                        ["Ordem", "Origem", "O que significa", "Formula", "Valor", "Linhas", "Onde conferir"],
                        current_dataset["fiscal_component_rows"],
                    ),
                    (
                        "Ajuste fiscal detalhado",
                        [
                            "Origem", "Operacao", "Modelo", "Chave/Id", "Numero", "Serie", "Data", "Valor XML", "Valor Doc SPED C100",
                            "Valor Resumo SPED", "Diferenca Resumo - Doc", "Linha Doc SPED", "Linhas Resumo SPED", "VL_MERC C100", "VL_DESC C100", "VL_FRT C100", "VL_SEG C100",
                            "VL_OUT", "VL_ST", "VL_IPI", "Formula", "Motivo provavel",
                        ],
                        current_dataset["fiscal_difference_rows"],
                    ),
                ],
            )
            messagebox.showinfo("Investigacao", f"Arquivo gerado com sucesso:\n{selected}")
        except Exception as exc:
            messagebox.showerror("Investigacao", str(exc))

    def toggle_compare_invoice_mark(self, event: object = None) -> str | None:
        if not hasattr(self, "compare_tree") or self.compare_last_result_origin != "xml":
            return None
        row_id = self.compare_tree.identify_row(getattr(event, "y", 0))
        column_id = self.compare_tree.identify_column(getattr(event, "x", 0))
        mark_column_id = "#1"
        if not row_id or column_id != mark_column_id:
            return None
        values = list(self.compare_tree.item(row_id, "values"))
        if len(values) < 2 or str(values[1]).strip() != "Nao lancada":
            return "break"
        while len(values) <= 10:
            values.append("")
        values[0] = COMPARE_MARK_CHECKED if str(values[0]).strip() != COMPARE_MARK_CHECKED else COMPARE_MARK_UNCHECKED
        self.compare_tree.item(row_id, values=values)
        return "break"

    def mark_all_missing_compare_invoices(self) -> None:
        if not hasattr(self, "compare_tree") or self.compare_last_result_origin != "xml":
            messagebox.showinfo("Marcacao", "A marcacao em lote funciona apenas para comparacao gerada a partir de XML.")
            return
        marked = 0
        for row_id in self.compare_tree.get_children():
            values = list(self.compare_tree.item(row_id, "values"))
            if len(values) > 1 and str(values[1]).strip() == "Nao lancada":
                while len(values) <= 10:
                    values.append("")
                values[0] = COMPARE_MARK_CHECKED
                self.compare_tree.item(row_id, values=values)
                marked += 1
        self.compare_status_var.set(f"{marked} nota(s) nao lancada(s) marcada(s) para lancamento.")

    def clear_compare_invoice_marks(self) -> None:
        if not hasattr(self, "compare_tree"):
            return
        for row_id in self.compare_tree.get_children():
            values = list(self.compare_tree.item(row_id, "values"))
            if values and str(values[0]).strip():
                values[0] = COMPARE_MARK_UNCHECKED if len(values) > 1 and str(values[1]).strip() == "Nao lancada" else ""
                self.compare_tree.item(row_id, values=values)
        self.compare_status_var.set("Marcacoes limpas.")

    def get_marked_compare_invoice_rows(self) -> list[str]:
        if not hasattr(self, "compare_tree"):
            return []
        marked_rows: list[str] = []
        for row_id in self.compare_tree.get_children():
            values = self.compare_tree.item(row_id, "values")
            if values and str(values[0]).strip() == COMPARE_MARK_CHECKED:
                marked_rows.append(str(row_id))
        return marked_rows

    def parse_compare_invoices_from_rows(self, row_ids: Iterable[str]) -> tuple[list[CompareXmlInvoice], list[str]]:
        invoices: list[CompareXmlInvoice] = []
        errors: list[str] = []
        seen_paths: set[str] = set()
        for row_id in row_ids:
            values = self.compare_tree.item(row_id, "values")
            if len(values) < 2 or str(values[1]).strip() != "Nao lancada":
                continue
            source_path = str(values[10]).strip() if len(values) > 10 else ""
            if not source_path or source_path in seen_paths:
                continue
            seen_paths.add(source_path)
            invoice = parse_compare_xml_file(Path(source_path))
            if invoice is None:
                errors.append(f"Nao foi possivel ler o XML: {source_path}")
                continue
            if str(invoice.model or "").strip() == "65":
                errors.append(f"NFC-e modelo 65 ignorada no lancamento de entradas: {source_path}")
                continue
            if not invoice.items:
                errors.append(f"XML sem itens: {source_path}")
                continue
            invoices.append(invoice)
        return invoices, errors

    def launch_marked_compare_invoices(self) -> None:
        if self.compare_last_result_origin != "xml":
            messagebox.showinfo("Lancamento", "O lancamento em lote funciona apenas para resultados gerados a partir de XML.")
            return
        marked_rows = self.get_marked_compare_invoice_rows()
        if not marked_rows:
            selected = list(getattr(self, "compare_tree", None).selection()) if hasattr(self, "compare_tree") else []
            marked_rows = [
                row_id
                for row_id in selected
                if len(self.compare_tree.item(row_id, "values")) > 1
                and str(self.compare_tree.item(row_id, "values")[1]).strip() == "Nao lancada"
            ]
        if not marked_rows:
            messagebox.showinfo("Lancamento", "Marque uma ou mais notas nao lancadas.")
            return
        sped_icms_text = self.compare_sped_icms_path_var.get().strip()
        sped_contrib_text = self.compare_sped_contrib_path_var.get().strip()
        sped_icms = Path(sped_icms_text)
        sped_contrib = Path(sped_contrib_text)
        if not sped_icms.exists():
            messagebox.showerror("Lancamento", "Selecione um SPED ICMS valido.")
            return
        if sped_contrib_text and not sped_contrib.exists():
            messagebox.showerror("Lancamento", "O SPED Contribuicoes informado nao existe.")
            return
        invoices, errors = self.parse_compare_invoices_from_rows(marked_rows)
        if not invoices:
            messagebox.showerror("Lancamento", "\n".join(errors) if errors else "Nenhuma nota valida para lancar.")
            return
        if not messagebox.askyesno("Lancamento em lote", f"Lancar {len(invoices)} nota(s) em um unico arquivo SPED?"):
            return
        has_sped_contrib = bool(sped_contrib_text) and sped_contrib.exists()
        try:
            launched_icms, output_icms, message_icms = launch_compare_invoices_in_sped(sped_icms, invoices)
            launched_contrib, output_contrib, message_contrib = 0, Path(""), "SPED Contribuicoes nao informado."
            if has_sped_contrib:
                launched_contrib, output_contrib, message_contrib = launch_compare_invoices_in_sped(sped_contrib, invoices)
            if launched_icms:
                for row_id in marked_rows:
                    values = list(self.compare_tree.item(row_id, "values"))
                    if len(values) > 1 and str(values[1]).strip() == "Nao lancada":
                        values[0] = ""
                        values[1] = "Lancada (gerada)"
                        self.compare_tree.item(row_id, values=values)
                self.compare_status_var.set(f"Lancamento em lote concluido: {launched_icms} nota(s) no SPED ICMS.")
                message = f"SPED ICMS: {output_icms}\nNotas lancadas: {launched_icms}"
                if has_sped_contrib:
                    message += f"\n\nSPED Contribuicoes: {output_contrib}\nNotas lancadas: {launched_contrib}"
                if errors:
                    message += "\n\nAvisos:\n" + "\n".join(errors[:10])
                messagebox.showinfo("Lancamento concluido", message)
            else:
                messagebox.showwarning("Lancamento", f"Nenhuma nota nova foi lancada.\n\n{message_icms}")
        except Exception as exc:
            messagebox.showerror("Lancamento", str(exc))

    def launch_selected_compare_invoice(self) -> None:
        if self.compare_last_result_origin != "xml":
            messagebox.showinfo("Lancamento", "O lancamento de nota funciona apenas para resultados gerados a partir de XML.")
            return
        selected = getattr(self, "compare_tree", None).selection() if hasattr(self, "compare_tree") else ()
        if not selected:
            messagebox.showinfo("Lancamento", "Selecione uma nota na tabela.")
            return
        values = self.compare_tree.item(selected[0], "values")
        if not values:
            return
        selected_status = str(values[1]).strip() if len(values) > 1 else ""
        if selected_status.upper().startswith("LANCADA"):
            messagebox.showinfo("Lancamento", "Essa nota ja esta lancada no SPED selecionado.")
            return
        if selected_status != "Nao lancada":
            messagebox.showinfo("Lancamento", "O lancamento automatico usa apenas XML existente na pasta e ainda nao lancado no SPED.")
            return
        sped_icms_text = self.compare_sped_icms_path_var.get().strip()
        sped_contrib_text = self.compare_sped_contrib_path_var.get().strip()
        sped_icms = Path(sped_icms_text)
        sped_contrib = Path(sped_contrib_text)
        if not sped_icms.exists():
            messagebox.showerror("Lancamento", "Selecione um SPED ICMS valido.")
            return
        if sped_contrib_text and not sped_contrib.exists():
            messagebox.showerror("Lancamento", "O SPED Contribuicoes informado nao existe.")
            return
        has_sped_contrib = bool(sped_contrib_text) and sped_contrib.exists()
        invoice = parse_compare_xml_file(Path(values[10]))
        if invoice is None:
            messagebox.showerror("Lancamento", "Nao foi possivel ler o XML selecionado.")
            return
        if str(invoice.model or "").strip() == "65":
            messagebox.showinfo("Lancamento", "NFC-e modelo 65 nao deve ser lancada como entrada por XML.")
            return
        if not invoice.items:
            messagebox.showerror("Lancamento", "XML sem itens para gerar C170.")
            return
        try:
            ok_icms, output_icms, message_icms = launch_compare_invoice_in_sped(sped_icms, invoice)
            ok_contrib, output_contrib, message_contrib = True, Path(""), "SPED Contribuicoes nao informado."
            if has_sped_contrib:
                ok_contrib, output_contrib, message_contrib = launch_compare_invoice_in_sped(sped_contrib, invoice)
            if ok_icms and ok_contrib:
                self.compare_tree.set(selected[0], "status", "Lancada (gerada)")
                self.compare_tree.set(selected[0], "launch_mark", "")
                if has_sped_contrib:
                    self.compare_status_var.set("Lancamento concluido nos dois SPEDs.")
                    messagebox.showinfo("Lancamento concluido", f"SPED ICMS: {output_icms}\nSPED Contribuicoes: {output_contrib}")
                else:
                    self.compare_status_var.set("Lancamento concluido no SPED ICMS.")
                    messagebox.showinfo("Lancamento concluido", f"SPED ICMS: {output_icms}")
            else:
                self.compare_status_var.set("Lancamento parcial ou ja existente.")
                messagebox.showwarning("Lancamento", f"ICMS: {message_icms} ({output_icms})\nContribuicoes: {message_contrib} ({output_contrib})")
        except Exception as exc:
            messagebox.showerror("Lancamento", str(exc))

    def open_compare_invoice_mirror_popup(self, event: object = None) -> None:
        if not hasattr(self, "compare_tree"):
            return
        row_id = ""
        if event is not None:
            row_id = self.compare_tree.identify_row(getattr(event, "y", 0))
        selected = (row_id,) if row_id else self.compare_tree.selection()
        if not selected:
            return
        values = self.compare_tree.item(selected[0], "values")
        if not values:
            return
        source_path = str(values[10]).strip() if len(values) > 10 else ""
        if not source_path.lower().endswith(".xml"):
            self.compare_status_var.set("Linha sem XML disponivel para abrir o espelho da nota.")
            return
        xml_path = Path(source_path)
        if not xml_path.exists():
            messagebox.showwarning("Espelho da Nota", "XML de origem nao encontrado no disco.")
            return
        invoice = parse_compare_xml_file(xml_path)
        if invoice is None:
            messagebox.showwarning("Espelho da Nota", "Nao foi possivel ler o XML selecionado.")
            return
        if not invoice.items:
            self.compare_status_var.set("XML selecionado nao possui itens para montar o espelho da nota.")
            return
        self.open_compare_xml_invoice_mirror(invoice)

    def open_compare_xml_invoice_mirror(self, invoice: CompareXmlInvoice) -> None:
        if not invoice.items:
            messagebox.showwarning("Espelho XML", "XML sem itens para montar o espelho da nota.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Espelho XML - {invoice.number}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1360, 720, 1100, 600, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        header_box = ttk.LabelFrame(container, text="Cabecalho do XML", padding=10)
        header_box.grid(row=0, column=0, sticky="ew")
        header_box.columnconfigure(0, weight=1)
        header_box.columnconfigure(1, weight=1)
        header_left = ttk.Frame(header_box)
        header_left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        header_right = ttk.Frame(header_box)
        header_right.grid(row=0, column=1, sticky="nw")

        header_rows_left = [
            ("Chave/Id", invoice.key),
            ("Numero", invoice.number),
            ("Serie", invoice.series),
            ("Modelo", invoice.model),
            ("Emissao", invoice.issue_date),
            ("Arquivo", invoice.file_path),
        ]
        header_rows_right = [
            ("Emitente", invoice.issuer_name),
            ("CNPJ Emitente", normalize_document_key(invoice.issuer_cnpj)),
            ("Destinatario", invoice.recipient_name),
            ("CNPJ/CPF Dest.", normalize_document_key(invoice.recipient_cnpj)),
            ("Valor XML", invoice.total_value),
            ("Itens", len(invoice.items)),
        ]
        for label_text, value_text in header_rows_left:
            ttk.Label(header_left, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)
        for label_text, value_text in header_rows_right:
            ttk.Label(header_right, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=1, column=0, sticky="w", pady=(10, 8))
        ttk.Button(
            export_actions,
            text="Exportar Espelho Excel",
            style="Secondary.TButton",
            command=lambda: self.export_compare_xml_invoice_mirror(invoice, "xlsx"),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Espelho CSV",
            style="Secondary.TButton",
            command=lambda: self.export_compare_xml_invoice_mirror(invoice, "csv"),
        ).pack(side=LEFT, padx=(8, 0))

        items_box = ttk.Frame(container)
        items_box.grid(row=2, column=0, sticky="nsew")
        items_box.columnconfigure(0, weight=1)
        items_box.rowconfigure(0, weight=1)

        columns = (
            "item_no",
            "code",
            "description",
            "ncm",
            "cfop",
            "quantity",
            "operation_value",
            "discount",
            "base_icms",
            "icms_rate",
            "icms_value",
            "base_icms_st",
            "icms_st_value",
            "ipi_value",
            "base_pis",
            "pis_value",
            "base_cofins",
            "cofins_value",
        )
        headings = {
            "item_no": "Item",
            "code": "Codigo",
            "description": "Descricao",
            "ncm": "NCM",
            "cfop": "CFOP",
            "quantity": "Quantidade",
            "operation_value": "Valor Operacao",
            "discount": "Desconto",
            "base_icms": "Base ICMS",
            "icms_rate": "Aliq ICMS",
            "icms_value": "Valor ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "ipi_value": "Valor IPI",
            "base_pis": "Base PIS",
            "pis_value": "Valor PIS",
            "base_cofins": "Base COFINS",
            "cofins_value": "Valor COFINS",
        }
        widths = {
            "item_no": 60,
            "code": 100,
            "description": 320,
            "ncm": 100,
            "cfop": 85,
            "quantity": 95,
            "operation_value": 120,
            "discount": 100,
            "base_icms": 110,
            "icms_rate": 90,
            "icms_value": 110,
            "base_icms_st": 120,
            "icms_st_value": 120,
            "ipi_value": 100,
            "base_pis": 110,
            "pis_value": 100,
            "base_cofins": 110,
            "cofins_value": 110,
        }
        tree = ttk.Treeview(items_box, columns=columns, show="headings", height=15, selectmode="extended")
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))

        scroll_y = ttk.Scrollbar(items_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(items_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        totals = {
            "quantity": Decimal("0"),
            "operation": Decimal("0"),
            "base_icms": Decimal("0"),
            "icms": Decimal("0"),
            "ipi": Decimal("0"),
            "pis": Decimal("0"),
            "cofins": Decimal("0"),
        }
        for item in invoice.items:
            quantity = Decimal(str(item.quantity or 0))
            operation_value = Decimal(str(item.value or 0))
            base_icms = Decimal(str(item.vl_bc_icms or 0))
            icms_value = Decimal(str(item.vl_icms or 0))
            ipi_value = Decimal(str(item.vl_ipi or 0))
            pis_value = Decimal(str(item.vl_pis or 0))
            cofins_value = Decimal(str(item.vl_cofins or 0))
            tree.insert(
                "",
                END,
                values=(
                    item.item_no,
                    item.code,
                    item.description,
                    item.ncm,
                    item.cfop,
                    format_decimal_sped(quantity),
                    format_decimal_sped(operation_value),
                    format_decimal_sped(Decimal(str(item.discount or 0))),
                    format_decimal_sped(base_icms),
                    format_decimal_sped(Decimal(str(item.aliq_icms or 0))),
                    format_decimal_sped(icms_value),
                    format_decimal_sped(Decimal(str(item.vl_bc_icms_st or 0))),
                    format_decimal_sped(Decimal(str(item.vl_icms_st or 0))),
                    format_decimal_sped(ipi_value),
                    format_decimal_sped(Decimal(str(item.vl_bc_pis or 0))),
                    format_decimal_sped(pis_value),
                    format_decimal_sped(Decimal(str(item.vl_bc_cofins or 0))),
                    format_decimal_sped(cofins_value),
                ),
            )
            totals["quantity"] += quantity
            totals["operation"] += operation_value
            totals["base_icms"] += base_icms
            totals["icms"] += icms_value
            totals["ipi"] += ipi_value
            totals["pis"] += pis_value
            totals["cofins"] += cofins_value

        footer = ttk.Frame(container)
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Itens: {len(invoice.items)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Quantidade: {format_decimal_sped(totals['quantity'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(totals['operation'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(totals['base_icms'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(totals['icms'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor IPI: {format_decimal_sped(totals['ipi'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"PIS: {format_decimal_sped(totals['pis'])}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"COFINS: {format_decimal_sped(totals['cofins'])}", style="Strong.TLabel").pack(side=LEFT)

    def export_compare_xml_invoice_mirror(self, invoice: CompareXmlInvoice, output_type: str) -> None:
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Salvar espelho XML",
            defaultextension=suffix,
            initialfile=f"espelho_xml_{invoice.number or 'sem_numero'}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return
        header_headers = ["Campo", "Valor"]
        header_rows = [
            ["Chave/Id", invoice.key],
            ["Numero", invoice.number],
            ["Serie", invoice.series],
            ["Modelo", invoice.model],
            ["Emissao", invoice.issue_date],
            ["Emitente", invoice.issuer_name],
            ["CNPJ Emitente", normalize_document_key(invoice.issuer_cnpj)],
            ["Destinatario", invoice.recipient_name],
            ["CNPJ/CPF Destinatario", normalize_document_key(invoice.recipient_cnpj)],
            ["Valor XML", invoice.total_value],
            ["Arquivo", invoice.file_path],
        ]
        item_headers = ["Item", "Codigo", "Descricao", "NCM", "CFOP", "Quantidade", "Valor Operacao", "Desconto", "Base ICMS", "Aliq ICMS", "Valor ICMS", "Base ICMS ST", "Valor ICMS ST", "Valor IPI", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS"]
        item_rows = [
            [
                item.item_no,
                item.code,
                item.description,
                item.ncm,
                item.cfop,
                Decimal(str(item.quantity or 0)),
                Decimal(str(item.value or 0)),
                Decimal(str(item.discount or 0)),
                Decimal(str(item.vl_bc_icms or 0)),
                Decimal(str(item.aliq_icms or 0)),
                Decimal(str(item.vl_icms or 0)),
                Decimal(str(item.vl_bc_icms_st or 0)),
                Decimal(str(item.vl_icms_st or 0)),
                Decimal(str(item.vl_ipi or 0)),
                Decimal(str(item.vl_bc_pis or 0)),
                Decimal(str(item.vl_pis or 0)),
                Decimal(str(item.vl_bc_cofins or 0)),
                Decimal(str(item.vl_cofins or 0)),
            ]
            for item in invoice.items
        ]
        try:
            output_path = Path(selected)
            if output_path.suffix.lower() == ".csv":
                header_output = output_path.with_name(f"{output_path.stem}_cabecalho.csv")
                write_simple_csv_file(header_output, header_headers, header_rows)
                write_simple_csv_file(output_path, item_headers, item_rows)
            else:
                write_simple_excel_workbook(output_path, [("Cabecalho XML", header_headers, header_rows), ("Itens XML", item_headers, item_rows)])
            messagebox.showinfo("Espelho XML", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Espelho XML", str(exc))

    def open_compare_invoice_source_location(self, _event: object = None) -> None:
        if not hasattr(self, "compare_tree"):
            return
        selected = self.compare_tree.selection()
        if not selected:
            return
        values = self.compare_tree.item(selected[0], "values")
        if not values:
            return
        source_path = str(values[10]).strip() if len(values) > 10 else ""
        if not source_path or not os.path.exists(source_path):
            messagebox.showwarning("Arquivo origem", "Arquivo de origem nao encontrado no disco.")
            return
        target_path = source_path if os.path.isdir(source_path) else os.path.dirname(source_path)
        try:
            os.startfile(target_path)
        except OSError:
            messagebox.showwarning("Arquivo origem", "Nao foi possivel abrir a pasta do arquivo.")

    def log_message(self, message: str) -> None:
        if hasattr(self, "log"):
            self.log.insert("", END, values=(message,))
            self.log.yview_moveto(1.0)
        self.status_var.set(message)
        self.write_audit_log("EVENTO", message)

    def write_audit_log(self, event_type: str, message: str) -> None:
        timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        event = str(event_type or "EVENTO").strip().upper()
        text = str(message or "").replace("\r", " ").replace("\n", " | ").strip()
        try:
            self.audit_log_path.parent.mkdir(parents=True, exist_ok=True)
            with self.audit_log_path.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{timestamp}\t{self.audit_session_id}\t{event}\t{text}\n")
        except Exception:
            pass

    def format_audit_paths(self, paths: Iterable[Path] | Iterable[str]) -> str:
        formatted = [str(path) for path in paths if str(path).strip()]
        return "; ".join(formatted) if formatted else "Nao informado"

    def audit_decimal_from_row(self, row: dict[str, object], *keys: str) -> Decimal:
        for key in keys:
            value = row.get(key)
            if isinstance(value, Decimal):
                return value
            if value not in (None, ""):
                try:
                    return parse_decimal(str(value))
                except Exception:
                    continue
        return Decimal("0.00")

    def summarize_audit_detail_rows(self, rows: list[dict[str, object]]) -> dict[str, Decimal | int]:
        return {
            "linhas": len(rows),
            "valor_operacao": sum((self.audit_decimal_from_row(row, "sale_value", "total_operation_value", "operation_value") for row in rows), Decimal("0.00")).quantize(Decimal("0.01")),
            "base_icms": sum((self.audit_decimal_from_row(row, "base_icms") for row in rows), Decimal("0.00")).quantize(Decimal("0.01")),
            "valor_icms": sum((self.audit_decimal_from_row(row, "icms_value") for row in rows), Decimal("0.00")).quantize(Decimal("0.01")),
        }

    def format_audit_summary(self, summary: dict[str, Decimal | int]) -> str:
        return (
            f"linhas={summary.get('linhas', 0)}; "
            f"valor_operacao={format_decimal_sped(Decimal(summary.get('valor_operacao', Decimal('0.00'))))}; "
            f"base_icms={format_decimal_sped(Decimal(summary.get('base_icms', Decimal('0.00'))))}; "
            f"valor_icms={format_decimal_sped(Decimal(summary.get('valor_icms', Decimal('0.00'))))}"
        )

    def close_application(self) -> None:
        if not self.audit_session_closed:
            self.write_audit_log("FIM_SESSAO", "Sistema fechado pelo usuario.")
            self.audit_session_closed = True
        self.root.destroy()

    def show_front_message(self, kind: str, title: str, message: str) -> None:
        self.root.update_idletasks()
        self.root.lift()
        self.root.attributes("-topmost", True)
        try:
            getattr(messagebox, kind)(title, message, parent=self.root)
        finally:
            self.root.attributes("-topmost", False)
            self.root.lift()
            self.root.focus_force()

    def select_file(self, target_var: StringVar, title: str, filetypes: list[tuple[str, str]]) -> None:
        selected = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
        )
        if selected:
            target_var.set(selected)
            self.log_message(f"Arquivo selecionado: {selected}")

    def select_files_or_directory(self, target_var: StringVar, title: str, filetypes: list[tuple[str, str]]) -> None:
        normalized_title = normalize_text(title)
        prefer_directory = "xml" in normalized_title or "pasta" in normalized_title

        if prefer_directory:
            selected_directory = filedialog.askdirectory(
                title=title,
            )
            current_paths = parse_selected_paths(target_var.get())
            if selected_directory:
                selected_path = Path(selected_directory)
                current_paths = [path for path in current_paths if not path.is_file()]
                if selected_path not in current_paths:
                    current_paths.append(selected_path)
                target_var.set(format_selected_paths(current_paths))
                self.log_message(f"Pasta de XML adicionada: {selected_directory}")
                return

        selected_files = list(
            filedialog.askopenfilenames(
                title=title,
                filetypes=filetypes,
            )
        )
        current_paths = parse_selected_paths(target_var.get())
        if selected_files:
            collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(selected_files)
            if collapsed_to_folder:
                current_paths = [path for path in current_paths if not path.is_file()]
            for selected_path in collapsed_paths:
                if selected_path not in current_paths:
                    current_paths.append(selected_path)
            target_var.set(format_selected_paths(current_paths))
            if collapsed_to_folder:
                self.log_message(
                    f"{len(selected_files)} XML(s) detectado(s). Seleção resumida para a pasta: {collapsed_paths[0]}"
                )
            else:
                self.log_message(f"{len(selected_files)} XML(s) selecionado(s).")
            return

        selected_directory = filedialog.askdirectory(
            title=title,
        )
        if selected_directory:
            selected_path = Path(selected_directory)
            if selected_path not in current_paths:
                current_paths.append(selected_path)
            target_var.set(format_selected_paths(current_paths))
            self.log_message(f"Caminho adicionado: {selected_directory}")

    def select_multi_sped_files(self) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecione de 1 a 12 arquivos SPED",
                filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return

        current_paths = parse_selected_paths(self.multi_sped_paths_var.get())
        for selected in selected_files:
            selected_path = Path(selected)
            if selected_path not in current_paths:
                current_paths.append(selected_path)

        if len(current_paths) > 12:
            messagebox.showwarning("Limite excedido", "A analise aceita no maximo 12 arquivos SPED.")
            current_paths = current_paths[:12]

        self.multi_sped_paths_var.set(format_selected_paths(current_paths))
        self.log_message(f"{len(current_paths)} arquivo(s) SPED selecionado(s) para a analise de entradas.")

    def select_consult_sped_files(self) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecione de 1 a 12 arquivos SPED para consulta",
                filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return

        current_paths = parse_selected_paths(self.consult_sped_paths_var.get())
        for selected in selected_files:
            selected_path = Path(selected)
            if selected_path not in current_paths:
                current_paths.append(selected_path)

        if len(current_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            current_paths = current_paths[:12]

        self.consult_sped_paths_var.set(format_selected_paths(current_paths))
        self.log_message(f"{len(current_paths)} arquivo(s) SPED preparado(s) para consulta em tela.")

    def select_sales_consult_sped_files(self) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecione de 1 a 12 arquivos SPED para consulta de saidas",
                filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return

        current_paths = parse_selected_paths(self.sales_consult_sped_paths_var.get())
        for selected in selected_files:
            selected_path = Path(selected)
            if selected_path not in current_paths:
                current_paths.append(selected_path)

        if len(current_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            current_paths = current_paths[:12]

        self.sales_consult_sped_paths_var.set(format_selected_paths(current_paths))
        self.log_message(f"{len(current_paths)} arquivo(s) SPED preparado(s) para consulta de saidas.")

    def select_fiscal_consult_xml_sources(self) -> None:
        current_value = self.consult_xml_paths_var.get().strip() or self.sales_consult_xml_paths_var.get().strip()
        current_var = StringVar(value=current_value)
        self.select_files_or_directory(
            current_var,
            "Selecione XMLs NF-e/NFC-e modelos 55/65 ou uma pasta com XMLs",
            [("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")],
        )
        formatted_paths = current_var.get().strip()
        if not formatted_paths:
            return
        self.consult_xml_paths_var.set(formatted_paths)
        self.sales_consult_xml_paths_var.set(formatted_paths)
        xml_count = len(parse_selected_paths(formatted_paths))
        self.log_message(f"{xml_count} origem(ns) XML preparada(s) para recomposicao de entradas e saidas ICMS.")

    def select_contrib_consult_sped_files(self) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecione de 1 a 12 arquivos SPED Contribuicoes para consulta de entradas",
                filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return
        current_paths = parse_selected_paths(self.contrib_consult_sped_paths_var.get())
        for selected in selected_files:
            selected_path = Path(selected)
            if selected_path not in current_paths:
                current_paths.append(selected_path)
        if len(current_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            current_paths = current_paths[:12]
        formatted_paths = format_selected_paths(current_paths)
        self.contrib_consult_sped_paths_var.set(formatted_paths)
        self.contrib_sales_consult_sped_paths_var.set(formatted_paths)
        self.log_message(f"{len(current_paths)} arquivo(s) SPED Contribuicoes preparado(s) para consulta de entradas e saidas.")

    def select_contrib_sales_consult_sped_files(self) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecione de 1 a 12 arquivos SPED Contribuicoes para consulta de saidas",
                filetypes=[("Arquivos SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return
        current_paths = parse_selected_paths(self.contrib_sales_consult_sped_paths_var.get())
        for selected in selected_files:
            selected_path = Path(selected)
            if selected_path not in current_paths:
                current_paths.append(selected_path)
        if len(current_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            current_paths = current_paths[:12]
        formatted_paths = format_selected_paths(current_paths)
        self.contrib_sales_consult_sped_paths_var.set(formatted_paths)
        self.contrib_consult_sped_paths_var.set(formatted_paths)
        self.log_message(f"{len(current_paths)} arquivo(s) SPED Contribuicoes preparado(s) para consulta de entradas e saidas.")

    def select_contrib_xml_sources(self) -> None:
        current_value = self.contrib_consult_xml_paths_var.get().strip() or self.contrib_sales_consult_xml_paths_var.get().strip()
        current_var = StringVar(value=current_value)
        self.select_files_or_directory(
            current_var,
            "Selecione XMLs NF-e/NFC-e modelos 55/65 ou uma pasta com XMLs",
            [("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")],
        )
        formatted_paths = current_var.get().strip()
        if not formatted_paths:
            return
        self.contrib_consult_xml_paths_var.set(formatted_paths)
        self.contrib_sales_consult_xml_paths_var.set(formatted_paths)
        xml_count = len(parse_selected_paths(formatted_paths))
        self.log_message(f"{xml_count} origem(ns) XML preparada(s) para recomposicao de entradas e saidas de Contribuicoes.")

    def get_current_manual_filters(self) -> tuple[list[str], set[str], set[str]]:
        descriptions = [
            line.strip()
            for line in self.filter_text.get("1.0", END).splitlines()
            if line.strip()
        ]
        cst_filter_values = parse_filter_values(self.cst_filter_var.get())
        cfop_filter_values = parse_filter_values(self.cfop_filter_var.get())
        return descriptions, cst_filter_values, cfop_filter_values

    def load_current_popup_context(self, source_view: str = "default") -> dict[str, object]:
        sped_source = self.sped_path_var.get().strip()
        excel_source = self.excel_path_var.get().strip()
        consult_sped_paths = parse_selected_paths(self.consult_sped_paths_var.get())
        sales_consult_sped_paths = parse_selected_paths(self.sales_consult_sped_paths_var.get())
        selected_consult_periods = self.get_selected_consult_periods() if hasattr(self, "consult_period_check_vars") else set()
        selected_sales_consult_periods = self.get_selected_sales_consult_periods() if hasattr(self, "sales_consult_period_check_vars") else set()

        sped_path = Path(sped_source) if sped_source else None
        excel_path = Path(excel_source) if excel_source else None
        if sped_path and not sped_path.exists():
            raise ValueError("O arquivo SPED selecionado nao existe mais.")
        if excel_path and not excel_path.exists():
            raise ValueError("A planilha selecionada nao existe mais.")
        if source_view == "sales_consult":
            if not sales_consult_sped_paths:
                raise ValueError("Carregue SPEDs na aba Consulta Saidas para abrir este popup.")
        elif source_view == "consult":
            if not consult_sped_paths:
                raise ValueError("Carregue SPEDs na aba Consulta Entradas para abrir este popup.")
        elif not sped_path and not excel_path and not consult_sped_paths:
            raise ValueError("Selecione o arquivo SPED Fiscal, a planilha Excel, ou carregue SPEDs na aba Consulta Entradas.")

        filter_descriptions, cst_filter_values, cfop_filter_values = self.get_current_manual_filters()
        selected_rule_profile = self.adjustment_mode_var.get().strip()
        if selected_rule_profile == "Filtros":
            selected_rule_profile = ""
        runtime_rules = parse_runtime_rule_lines(self.runtime_rules_text.get("1.0", END))

        source_sped_paths: list[Path] = []
        source_mode = "main"
        if source_view == "sales_consult":
            source_mode = "sales_consult"
            if selected_sales_consult_periods:
                source_sped_paths = [
                    path
                    for label in self.sales_consult_period_labels
                    if label in selected_sales_consult_periods
                    for path in [self.sales_consult_period_path_map.get(label)]
                    if path is not None and path.exists()
                ]
            else:
                source_sped_paths = [path for path in sales_consult_sped_paths if path.exists()]
            if not source_sped_paths:
                source_sped_paths = [path for path in sales_consult_sped_paths if path.exists()]
        elif source_view == "consult":
            source_mode = "consult"
            if selected_consult_periods:
                source_sped_paths = [
                    path
                    for label in self.consult_period_labels
                    if label in selected_consult_periods
                    for path in [self.consult_period_path_map.get(label)]
                    if path is not None and path.exists()
                ]
            else:
                source_sped_paths = [path for path in consult_sped_paths if path.exists()]
            if not source_sped_paths:
                source_sped_paths = [path for path in consult_sped_paths if path.exists()]
        elif sped_path or excel_path:
            source_sped_paths = [sped_path] if sped_path else []
        else:
            source_mode = "consult"
            if selected_consult_periods:
                source_sped_paths = [
                    path
                    for label in self.consult_period_labels
                    if label in selected_consult_periods
                    for path in [self.consult_period_path_map.get(label)]
                    if path is not None and path.exists()
                ]
            else:
                source_sped_paths = [path for path in consult_sped_paths if path.exists()]
            if not source_sped_paths:
                source_sped_paths = [path for path in consult_sped_paths if path.exists()]

        detailed_sales: list[dict[str, object]] = []
        summary_register_rows: list[dict[str, object]] = []
        validation_sped_path: Path | None = None

        if source_mode == "main":
            sped_data = read_sped_file(sped_path) if sped_path else None
            excel_data = read_detailed_product_excel(excel_path) if excel_path else None
            _, _, detailed_sales, c190_rows, _ = combine_imported_data(sped_data, excel_data)
            summary_register_rows = read_sped_summary_register_rows(sped_path) if sped_path else list(c190_rows)
            period_label = infer_sped_period_label(sped_path, detailed_sales) if sped_path else ""
            detailed_sales = [{**item, "period": period_label} for item in detailed_sales]
            summary_register_rows = [{**row, "period": period_label} for row in summary_register_rows]
            validation_sped_path = sped_path
        elif source_mode == "sales_consult":
            source_rows = list(getattr(self, "sales_consult_comparison_rows", []))
            selected_period_set = {
                label
                for label in self.sales_consult_period_labels
                if not selected_sales_consult_periods or label in selected_sales_consult_periods
            }
            seen_detail_keys: set[tuple[str, str, str, str, str, str, str]] = set()

            def append_sales_detail(detail: dict[str, object], period_label: str) -> None:
                normalized_detail = {**detail, "period": period_label}
                detail_key = (
                    str(period_label).strip(),
                    normalize_document_key(str(normalized_detail.get("document_key", ""))) or str(normalized_detail.get("document_number", "")).strip(),
                    str(normalized_detail.get("item_number", "")).strip(),
                    str(normalized_detail.get("code", "")).strip(),
                    str(normalized_detail.get("cfop", "")).strip(),
                    str(normalized_detail.get("cst_icms", "")).strip(),
                    format_decimal_sped(Decimal(normalized_detail.get("icms_rate", Decimal("0")))),
                )
                if detail_key in seen_detail_keys:
                    return
                seen_detail_keys.add(detail_key)
                detailed_sales.append(normalized_detail)

            for current_sped_path in source_sped_paths:
                _, _, current_details, current_c190_rows, _ = read_sped_file(current_sped_path)
                period_label = infer_sped_period_label(current_sped_path, current_details)
                for item in current_details:
                    append_sales_detail(item, period_label)
                summary_rows = read_sped_summary_register_rows(current_sped_path)
                source_summary_rows = summary_rows if summary_rows else current_c190_rows
                summary_register_rows.extend([{**row, "period": period_label} for row in source_summary_rows])
            for row in source_rows:
                period_label = str(row.get("period", "")).strip()
                if selected_period_set and period_label not in selected_period_set:
                    continue
                for detail in row.get("launch_details", []):
                    append_sales_detail(detail, period_label)
            if len(source_sped_paths) == 1:
                validation_sped_path = source_sped_paths[0]
        else:
            for current_sped_path in source_sped_paths:
                _, _, current_details, current_c190_rows, _ = read_sped_file(current_sped_path)
                period_label = infer_sped_period_label(current_sped_path, current_details)
                detailed_sales.extend([{**item, "period": period_label} for item in current_details])
                summary_rows = read_sped_summary_register_rows(current_sped_path)
                source_summary_rows = summary_rows if summary_rows else current_c190_rows
                summary_register_rows.extend([{**row, "period": period_label} for row in source_summary_rows])
            if len(source_sped_paths) == 1:
                validation_sped_path = source_sped_paths[0]

        if not detailed_sales and not summary_register_rows:
            raise ValueError("Nenhum dado valido foi encontrado no arquivo informado.")

        detailed_sales = enrich_details_with_note_snapshots(detailed_sales)

        rule_filtered = bool(selected_rule_profile)
        if source_mode == "consult":
            cst_filter_values = parse_filter_values(self.consult_cst_filter_var.get())
            cfop_filter_values = parse_filter_values(self.consult_cfop_filter_var.get())
            if self.consult_selected_product_code:
                filter_descriptions = [self.consult_selected_product_code]
            elif self.consult_search_var.get().strip():
                filter_descriptions = [self.consult_search_var.get().strip()]
            else:
                filter_descriptions = []
        elif source_mode == "sales_consult":
            cst_filter_values = parse_filter_values(self.sales_consult_cst_filter_var.get())
            cfop_filter_values = parse_filter_values(self.sales_consult_cfop_filter_var.get())
            if self.sales_consult_selected_product_code:
                filter_descriptions = [self.sales_consult_selected_product_code]
            elif self.sales_consult_search_var.get().strip():
                filter_descriptions = [self.sales_consult_search_var.get().strip()]
            else:
                filter_descriptions = []

        if rule_filtered:
            filtered_details = [
                item for item in detailed_sales
                if has_configured_icms_rule(selected_rule_profile, runtime_rules, item)
            ]
        else:
            filtered_details = filter_detailed_sales(
                detailed_sales,
                filter_descriptions,
                cst_filter_values,
                cfop_filter_values,
            )

        return {
            "sped_path": validation_sped_path,
            "excel_path": excel_path,
            "source_mode": source_mode,
            "source_sped_paths": source_sped_paths,
            "selected_rule_profile": selected_rule_profile,
            "runtime_rules": runtime_rules,
            "filter_descriptions": filter_descriptions,
            "cst_filter_values": cst_filter_values,
            "cfop_filter_values": cfop_filter_values,
            "all_details": detailed_sales,
            "filtered_details": filtered_details,
            "summary_register_rows": summary_register_rows,
            "can_use_c190_directly": not bool(filter_descriptions) and not bool(selected_rule_profile or runtime_rules),
            "filters_active": rule_filtered or has_active_item_filters(filter_descriptions, cst_filter_values, cfop_filter_values),
        }

    def format_popup_filter_caption(self, context: dict[str, object]) -> str:
        parts: list[str] = []
        descriptions = list(context.get("filter_descriptions", []))
        cst_values = sorted(str(value) for value in context.get("cst_filter_values", set()))
        cfop_values = sorted(str(value) for value in context.get("cfop_filter_values", set()))
        selected_rule_profile = str(context.get("selected_rule_profile", "")).strip()
        runtime_rules = list(context.get("runtime_rules", []))
        if selected_rule_profile:
            parts.append(f"Perfil: {selected_rule_profile}")
        if runtime_rules:
            parts.append(f"Regras dinamicas aplicadas: {len(runtime_rules)} (sem filtrar)")
        if descriptions:
            preview = ", ".join(descriptions[:3])
            if len(descriptions) > 3:
                preview += ", ..."
            parts.append(f"Descricoes: {preview}")
        if cst_values:
            parts.append(f"CST: {', '.join(cst_values)}")
        if cfop_values:
            parts.append(f"CFOP: {', '.join(cfop_values)}")
        return " | ".join(parts) if parts else "Sem filtros ativos."

    def load_current_contrib_popup_context(self, source_view: str = "contrib_consult") -> dict[str, object]:
        is_sales = source_view == "contrib_sales_consult"
        filtered_rows = list(self.contrib_sales_consult_filtered_rows if is_sales else self.contrib_consult_filtered_rows)
        all_rows = list(self.contrib_sales_consult_comparison_rows if is_sales else self.contrib_consult_comparison_rows)
        if not filtered_rows and not all_rows:
            raise ValueError("Processe a consulta de Contribuicoes antes de abrir este popup.")

        if is_sales:
            cst_filter_values = parse_filter_values(self.contrib_sales_consult_cst_filter_var.get())
            cfop_filter_values = parse_filter_values(self.contrib_sales_consult_cfop_filter_var.get())
            status_value = self.contrib_sales_consult_status_filter_var.get().strip()
            search_value = self.contrib_sales_consult_search_var.get().strip()
            selected_code = self.contrib_sales_consult_selected_product_code
            selected_periods = self.get_selected_contrib_sales_consult_periods()
        else:
            cst_filter_values = parse_filter_values(self.contrib_consult_cst_filter_var.get())
            cfop_filter_values = parse_filter_values(self.contrib_consult_cfop_filter_var.get())
            status_value = self.contrib_consult_status_filter_var.get().strip()
            search_value = self.contrib_consult_search_var.get().strip()
            selected_code = self.contrib_consult_selected_product_code
            selected_periods = self.get_selected_contrib_consult_periods()

        details: list[dict[str, object]] = []
        for row in filtered_rows:
            period_label = str(row.get("period", "")).strip()
            for detail in row.get("launch_details", []):
                details.append({**detail, "period": period_label})
        details = enrich_details_with_note_snapshots(details)

        filter_descriptions: list[str] = []
        if selected_code:
            filter_descriptions = [selected_code]
        elif search_value:
            filter_descriptions = [search_value]

        return {
            "source_view": source_view,
            "filtered_rows": filtered_rows,
            "all_rows": all_rows,
            "filtered_details": details,
            "cst_filter_values": cst_filter_values,
            "cfop_filter_values": cfop_filter_values,
            "status_value": status_value,
            "filter_descriptions": filter_descriptions,
            "selected_periods": selected_periods,
        }

    def format_contrib_popup_filter_caption(self, context: dict[str, object]) -> str:
        parts: list[str] = []
        descriptions = list(context.get("filter_descriptions", []))
        cst_values = sorted(str(value) for value in context.get("cst_filter_values", set()))
        cfop_values = sorted(str(value) for value in context.get("cfop_filter_values", set()))
        status_value = str(context.get("status_value", "")).strip()
        selected_periods = sorted((str(value) for value in context.get("selected_periods", set())), key=period_label_sort_key)
        if selected_periods:
            parts.append(f"Periodos: {', '.join(selected_periods)}")
        if descriptions:
            parts.append(f"Busca: {', '.join(descriptions[:3])}")
        if cst_values:
            parts.append(f"CST: {', '.join(cst_values)}")
        if cfop_values:
            parts.append(f"CFOP: {', '.join(cfop_values)}")
        if status_value and status_value.lower() != "todos":
            parts.append(f"Status: {status_value}")
        return " | ".join(parts) if parts else "Sem filtros ativos."

    def open_credit_diagnostic_popup(self, source_view: str = "default", operation_type: str = "Entrada") -> None:
        normalized_operation = normalize_operation_type(operation_type)
        is_saida = normalized_operation == "saida"
        diagnostic_caption = "Diagnostico de debito" if is_saida else "Diagnostico de credito"
        operation_caption = "Saidas" if is_saida else "Entradas"
        title_caption = "Debito" if is_saida else "Credito"
        screen_title = f"Diagnostico {title_caption} ICMS - {operation_caption}"
        header_title = f"Diagnostico da Base de {title_caption} ICMS - {operation_caption}"
        description_text = (
            "Analise fiscal baseada nos registros oficiais do SPED (C190/C590/D190/D590/D730). "
            "Use esta tela para entender por que a base e o debito do ICMS ficaram diferentes do valor da operacao."
            if is_saida
            else
            "Analise fiscal baseada nos registros oficiais do SPED (C190/C590/D190/D590/D730). "
            "Use esta tela para entender por que a base do credito ficou menor que o valor da operacao."
        )
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir {diagnostic_caption.lower()}: {exc}")
            messagebox.showerror(diagnostic_caption, str(exc))
            return

        summary_rows, detail_rows, totals = build_credit_diagnostic_datasets(
            list(context["summary_register_rows"]),
            operation_type,
            set(context["cst_filter_values"]),
            set(context["cfop_filter_values"]),
            list(context.get("filtered_details", [])),
        )
        if not detail_rows:
            messagebox.showwarning(diagnostic_caption, f"Nao ha dados de {operation_caption.lower()} para analisar.")
            return

        dialog = Toplevel(self.root)
        dialog.title(screen_title)
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1380, 760, 1120, 620, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)
        container.rowconfigure(5, weight=2)

        ttk.Label(
            container,
            text=header_title,
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=description_text,
            wraplength=1320,
        ).grid(row=1, column=0, sticky="w", pady=(6, 6))
        ttk.Label(
            container,
            text=self.format_popup_filter_caption(context),
            wraplength=1320,
        ).grid(row=2, column=0, sticky="w", pady=(0, 10))

        summary_box = ttk.LabelFrame(container, text="Resumo dos Motivos", padding=10)
        summary_box.grid(row=3, column=0, sticky="nsew")
        summary_box.columnconfigure(0, weight=1)
        summary_box.rowconfigure(0, weight=1)
        summary_columns = ("period", "reason", "rows", "operation_value", "base_icms", "base_gap", "base_ratio", "icms_value")
        summary_tree = ttk.Treeview(summary_box, columns=summary_columns, show="headings", height=8, selectmode="extended")
        summary_headings = {
            "period": "Periodo",
            "reason": "Motivo",
            "rows": "Linhas",
            "operation_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "base_gap": "Perda de Base",
            "base_ratio": "% Base/Oper",
            "icms_value": "Valor ICMS",
        }
        summary_widths = {
            "period": 90,
            "reason": 360,
            "rows": 70,
            "operation_value": 120,
            "base_icms": 120,
            "base_gap": 120,
            "base_ratio": 100,
            "icms_value": 120,
        }
        for column_id in summary_columns:
            summary_tree.column(column_id, width=summary_widths[column_id], anchor="center")
        self.enable_treeview_sorting(summary_tree, summary_columns, summary_headings)
        summary_tree.grid(row=0, column=0, sticky="nsew")
        summary_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(summary_tree))
        summary_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, summary_tree))
        summary_tree.bind(
            "<Double-1>",
            lambda event: self.handle_credit_diagnostic_summary_double_click(event, summary_tree, detail_rows),
        )
        summary_scroll = ttk.Scrollbar(summary_box, orient="vertical", command=summary_tree.yview)
        summary_scroll.grid(row=0, column=1, sticky="ns")
        summary_tree.configure(yscrollcommand=summary_scroll.set)

        for row_index, row in enumerate(summary_rows):
            summary_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row["period"],
                    row["reason"],
                    row["row_count"],
                    format_decimal_sped(Decimal(row["total_operation_value"])),
                    format_decimal_sped(Decimal(row["base_icms"])),
                    format_decimal_sped(Decimal(row["base_gap"])),
                    format_decimal_sped(Decimal(row["base_ratio"])),
                    format_decimal_sped(Decimal(row["icms_value"])),
                ),
            )

        detail_box = ttk.LabelFrame(container, text="Detalhamento Fiscal", padding=10)
        detail_box.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        detail_box.columnconfigure(0, weight=1)
        detail_box.rowconfigure(0, weight=1)
        detail_columns = ("period", "reason", "register", "cst", "cfop", "rate", "effective_rate", "rows", "operation_value", "base_icms", "base_gap", "base_ratio", "icms_value")
        detail_tree = ttk.Treeview(detail_box, columns=detail_columns, show="headings", height=14, selectmode="extended")
        detail_headings = {
            "period": "Periodo",
            "reason": "Motivo",
            "register": "Registro",
            "cst": "CST",
            "cfop": "CFOP",
            "rate": "Aliq",
            "effective_rate": "Aliq Efetiva",
            "rows": "Linhas",
            "operation_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "base_gap": "Perda de Base",
            "base_ratio": "% Base/Oper",
            "icms_value": "Valor ICMS",
        }
        detail_widths = {
            "period": 90,
            "reason": 320,
            "register": 70,
            "cst": 70,
            "cfop": 80,
            "rate": 80,
            "effective_rate": 95,
            "rows": 70,
            "operation_value": 120,
            "base_icms": 120,
            "base_gap": 120,
            "base_ratio": 100,
            "icms_value": 120,
        }
        for column_id in detail_columns:
            detail_tree.column(column_id, width=detail_widths[column_id], anchor="center")
        self.enable_treeview_sorting(detail_tree, detail_columns, detail_headings)
        detail_tree.grid(row=0, column=0, sticky="nsew")
        detail_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(detail_tree))
        detail_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, detail_tree))
        detail_tree.bind(
            "<Double-1>",
            lambda event: self.handle_credit_diagnostic_detail_double_click(event, detail_tree, detail_rows),
        )
        detail_scroll_y = ttk.Scrollbar(detail_box, orient="vertical", command=detail_tree.yview)
        detail_scroll_y.grid(row=0, column=1, sticky="ns")
        detail_scroll_x = ttk.Scrollbar(detail_box, orient="horizontal", command=detail_tree.xview)
        detail_scroll_x.grid(row=1, column=0, sticky="ew")
        detail_tree.configure(yscrollcommand=detail_scroll_y.set, xscrollcommand=detail_scroll_x.set)

        for row_index, row in enumerate(detail_rows):
            detail_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row["period"],
                    row["reason"],
                    row["source_register"],
                    row["cst_icms"],
                    row["cfop"],
                    format_decimal_sped(Decimal(row["icms_rate"])),
                    format_decimal_sped(compute_display_icms_rate(Decimal(row["icms_rate"]), Decimal(row["total_operation_value"]), Decimal(row["base_icms"]), Decimal(row["icms_value"]))),
                    row["row_count"],
                    format_decimal_sped(Decimal(row["total_operation_value"])),
                    format_decimal_sped(Decimal(row["base_icms"])),
                    format_decimal_sped(Decimal(row["base_gap"])),
                    format_decimal_sped(Decimal(row["base_ratio"])),
                    format_decimal_sped(Decimal(row["icms_value"])),
                ),
            )

        footer = ttk.Frame(container)
        footer.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        total_operation_value = Decimal(totals["total_operation_value"])
        total_base_icms = Decimal(totals["base_icms"])
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(total_operation_value)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base_icms)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Perda de Base: {format_decimal_sped(Decimal(totals['base_gap']))}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"% Base/Oper: {format_decimal_sped(total_base_ratio)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(Decimal(totals['icms_value']))}", style="Strong.TLabel").pack(side=LEFT)

        export_summary_rows = [
            [
                row["period"],
                row["reason"],
                row["row_count"],
                Decimal(row["total_operation_value"]),
                Decimal(row["base_icms"]),
                Decimal(row["base_gap"]),
                Decimal(row["base_ratio"]),
                Decimal(row["icms_value"]),
            ]
            for row in summary_rows
        ]
        export_detail_rows = [
            [
                row["period"],
                row["reason"],
                row["source_register"],
                row["cst_icms"],
                row["cfop"],
                Decimal(row["icms_rate"]),
                compute_display_icms_rate(Decimal(row["icms_rate"]), Decimal(row["total_operation_value"]), Decimal(row["base_icms"]), Decimal(row["icms_value"])),
                row["row_count"],
                Decimal(row["total_operation_value"]),
                Decimal(row["base_icms"]),
                Decimal(row["base_gap"]),
                Decimal(row["base_ratio"]),
                Decimal(row["icms_value"]),
            ]
            for row in detail_rows
        ]
        export_actions = ttk.Frame(container)
        export_actions.grid(row=7, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Diagnostico Excel",
            style="Secondary.TButton",
            command=lambda: self.export_credit_diagnostic_dataset(
                export_summary_rows,
                export_detail_rows,
                "xlsx",
                detail_rows,
                diagnostic_caption,
                f"diagnostico_{'debito' if is_saida else 'credito'}_icms",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Diagnostico CSV",
            style="Secondary.TButton",
            command=lambda: self.export_credit_diagnostic_dataset(
                export_summary_rows,
                export_detail_rows,
                "csv",
                detail_rows,
                diagnostic_caption,
                f"diagnostico_{'debito' if is_saida else 'credito'}_icms",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        self.log_message(f"Popup de {diagnostic_caption.lower()} aberto.")

    def open_credit_diagnostic_period_comparison_popup(self, source_view: str = "consult", operation_type: str = "Entrada") -> None:
        normalized_operation = normalize_operation_type(operation_type)
        is_saida = normalized_operation == "saida"
        diagnostic_caption = "Comparacao diagnostico de debito" if is_saida else "Comparacao diagnostico de credito"
        operation_caption = "Saidas" if is_saida else "Entradas"
        title_caption = "Debito" if is_saida else "Credito"
        source_tab_caption = "Consulta Saidas" if is_saida else "Consulta Entradas"
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir {diagnostic_caption.lower()}: {exc}")
            messagebox.showerror(diagnostic_caption, str(exc))
            return

        source_sped_paths = list(context.get("source_sped_paths", []))
        if len(source_sped_paths) < 2:
            messagebox.showwarning(
                diagnostic_caption,
                f"Selecione ao menos dois SPEDs na aba {source_tab_caption} para comparar periodos.",
            )
            return

        periods, export_headers, display_rows, export_rows, detail_rows = build_credit_diagnostic_period_comparison_dataset(
            list(context["summary_register_rows"]),
            list(context.get("filtered_details", [])),
            set(context["cst_filter_values"]),
            set(context["cfop_filter_values"]),
            operation_type,
        )
        detail_export_headers, detail_display_rows, detail_export_rows = build_credit_diagnostic_period_detail_comparison_dataset(
            detail_rows,
            periods,
        )
        if not display_rows or not periods:
            messagebox.showwarning(diagnostic_caption, "Nao ha dados de diagnostico para os periodos selecionados.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Comparacao Diagnostico {title_caption} ICMS - Periodos")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1560, 760, 1240, 560)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        ttk.Label(
            container,
            text=f"Comparacao do Diagnostico de {title_caption} ICMS por Periodo",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=(
                "Cada motivo aparece em uma linha. Dentro de cada periodo sao exibidas as metricas do diagnostico. "
                "Duplo clique em 'Valor Operacao' ou 'Base ICMS' abre os produtos e lancamentos do periodo clicado."
            ),
            wraplength=1500,
        ).grid(row=1, column=0, sticky="w", pady=(6, 6))
        ttk.Label(
            container,
            text=self.format_popup_filter_caption(context),
            wraplength=1500,
        ).grid(row=2, column=0, sticky="w", pady=(0, 10))

        period_band = ttk.Frame(container)
        period_band.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        if periods:
            ttk.Label(period_band, text="Periodos:", style="Strong.TLabel").pack(side=LEFT, padx=(0, 8))
        selected_period_state = {"period": ""}
        period_labels: dict[str, tk.Label] = {}

        paned = ttk.Panedwindow(container, orient="vertical")
        paned.grid(row=4, column=0, sticky="nsew")

        summary_box = ttk.LabelFrame(paned, text="Resumo por Motivo", padding=10)
        summary_box.columnconfigure(0, weight=1)
        summary_box.rowconfigure(0, weight=1)
        paned.add(summary_box, weight=2)

        detail_box = ttk.LabelFrame(paned, text="Detalhamento Fiscal Comparativo", padding=10)
        detail_box.columnconfigure(0, weight=1)
        detail_box.rowconfigure(0, weight=1)
        paned.add(detail_box, weight=3)

        columns = ["reason"]
        period_metric_columns: dict[str, tuple[str, str, str, str, str, str]] = {}
        for period in periods:
            period_key = period.replace("/", "_")
            metric_columns = (
                f"{period_key}_rows",
                f"{period_key}_operation_value",
                f"{period_key}_base_icms",
                f"{period_key}_base_gap",
                f"{period_key}_base_ratio",
                f"{period_key}_icms_value",
            )
            period_metric_columns[period] = metric_columns
            columns.extend(metric_columns)
        columns.extend(["total_rows", "total_operation_value", "total_base_icms", "total_base_gap", "total_base_ratio", "total_icms_value"])
        columns = tuple(columns)

        tree = ttk.Treeview(summary_box, columns=columns, show="headings", height=9, selectmode="extended")
        headings = {
            "reason": "Motivo",
            "total_rows": "Total Linhas",
            "total_operation_value": "Total Valor Operacao",
            "total_base_icms": "Total Base ICMS",
            "total_base_gap": "Total Perda de Base",
            "total_base_ratio": "Total % Base/Oper",
            "total_icms_value": "Total Valor ICMS",
        }
        widths = {
            "reason": 360,
            "total_rows": 90,
            "total_operation_value": 130,
            "total_base_icms": 130,
            "total_base_gap": 130,
            "total_base_ratio": 110,
            "total_icms_value": 130,
        }
        for period in periods:
            period_key = period.replace("/", "_")
            headings[f"{period_key}_rows"] = f"{period} Linhas"
            headings[f"{period_key}_operation_value"] = f"{period} Valor Oper."
            headings[f"{period_key}_base_icms"] = f"{period} Base ICMS"
            headings[f"{period_key}_base_gap"] = f"{period} Perda Base"
            headings[f"{period_key}_base_ratio"] = f"{period} % Base/Oper"
            headings[f"{period_key}_icms_value"] = f"{period} Valor ICMS"
            widths[f"{period_key}_rows"] = 80
            widths[f"{period_key}_operation_value"] = 120
            widths[f"{period_key}_base_icms"] = 120
            widths[f"{period_key}_base_gap"] = 120
            widths[f"{period_key}_base_ratio"] = 105
            widths[f"{period_key}_icms_value"] = 120

        sort_state = {"column": "total_base_gap", "reverse": True}
        tree_rows_state = list(display_rows)
        summary_total_columns = (
            "total_rows",
            "total_operation_value",
            "total_base_icms",
            "total_base_gap",
            "total_base_ratio",
            "total_icms_value",
        )

        def comparison_sort_value(row: dict[str, object], column_id: str) -> object:
            values = list(row.get("values", []))
            if column_id not in columns:
                return ""
            index = columns.index(column_id)
            value = values[index] if index < len(values) else ""
            if column_id == "reason":
                return normalize_text(str(value))
            if isinstance(value, Decimal):
                return value
            if isinstance(value, (int, float)):
                return value
            return normalize_text(str(value))

        def render_comparison_tree() -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)
            selected_period_filter = str(selected_period_state["period"]).strip()
            visible_rows = list(tree_rows_state)
            if selected_period_filter:
                row_metric_columns = period_metric_columns.get(selected_period_filter, ())
                row_count_index = columns.index(row_metric_columns[0]) if row_metric_columns else -1
                if row_count_index >= 0:
                    visible_rows = [
                        row
                        for row in visible_rows
                        if int(row.get("values", [])[row_count_index] or 0) > 0
                    ]
            sorted_rows = sorted(
                visible_rows,
                key=lambda row: comparison_sort_value(row, str(sort_state["column"])),
                reverse=bool(sort_state["reverse"]),
            )
            for row_index, row in enumerate(sorted_rows):
                tree.insert(
                    "",
                    END,
                    iid=str(row_index),
                    values=tuple(serialize_value_for_clipboard(value) for value in row["values"]),
                )
            tree._comparison_rows = sorted_rows

        def update_summary_display_columns() -> None:
            selected_period_filter = str(selected_period_state["period"]).strip()
            if selected_period_filter and selected_period_filter in period_metric_columns:
                tree.configure(
                    displaycolumns=("reason", *period_metric_columns[selected_period_filter], *summary_total_columns)
                )
            else:
                tree.configure(displaycolumns=columns)

        def sort_comparison_tree_by(column_id: str) -> None:
            if sort_state["column"] == column_id:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["column"] = column_id
                sort_state["reverse"] = False
            render_comparison_tree()

        for column_id in columns:
            tree.heading(column_id, text=headings[column_id], command=lambda current=column_id: sort_comparison_tree_by(current))
            tree.column(column_id, width=widths[column_id], anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))

        def handle_comparison_double_click(event: object) -> None:
            row_id = tree.identify_row(getattr(event, "y", 0))
            column_token = tree.identify_column(getattr(event, "x", 0))
            if not row_id or not column_token.startswith("#"):
                return
            try:
                row_index = int(row_id)
                column_index = int(column_token[1:]) - 1
            except ValueError:
                return
            filtered_rows = getattr(tree, "_comparison_rows", tree_rows_state)
            if row_index < 0 or row_index >= len(filtered_rows) or column_index < 0 or column_index >= len(columns):
                return
            column_id = columns[column_index]
            selected_reason = str(filtered_rows[row_index].get("reason", "")).strip()
            selected_period = next(
                (
                    period
                    for period, metric_columns in period_metric_columns.items()
                    if column_id in {metric_columns[1], metric_columns[2]}
                ),
                "",
            )
            if not selected_period:
                return
            merged_detail = merge_credit_diagnostic_detail_rows(detail_rows, selected_reason, selected_period)
            if not merged_detail:
                messagebox.showwarning(diagnostic_caption, "Nao ha itens para este motivo no periodo selecionado.")
                return
            self.open_credit_diagnostic_grouped_rows_popup(merged_detail, list(merged_detail.get("grouped_rows", [])))

        tree.bind("<Double-1>", handle_comparison_double_click)

        scroll_y = ttk.Scrollbar(summary_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(summary_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        update_summary_display_columns()
        render_comparison_tree()

        detail_columns = ["reason", "register", "cst", "cfop", "rate"]
        detail_period_metric_columns: dict[str, tuple[str, str, str, str, str, str, str]] = {}
        for period in periods:
            period_key = period.replace("/", "_")
            metric_columns = (
                f"{period_key}_detail_rows",
                f"{period_key}_detail_operation_value",
                f"{period_key}_detail_base_icms",
                f"{period_key}_detail_base_gap",
                f"{period_key}_detail_base_ratio",
                f"{period_key}_detail_icms_value",
                f"{period_key}_detail_effective_rate",
            )
            detail_period_metric_columns[period] = metric_columns
            detail_columns.extend(metric_columns)
        detail_columns.extend(
            [
                "detail_total_rows",
                "detail_total_operation_value",
                "detail_total_base_icms",
                "detail_total_base_gap",
                "detail_total_base_ratio",
                "detail_total_icms_value",
                "detail_total_effective_rate",
            ]
        )
        detail_columns = tuple(detail_columns)

        detail_tree = ttk.Treeview(detail_box, columns=detail_columns, show="headings", height=14, selectmode="extended")
        detail_headings = {
            "reason": "Motivo",
            "register": "Registro",
            "cst": "CST",
            "cfop": "CFOP",
            "rate": "Aliq",
            "detail_total_rows": "Total Linhas",
            "detail_total_operation_value": "Total Valor Operacao",
            "detail_total_base_icms": "Total Base ICMS",
            "detail_total_base_gap": "Total Perda de Base",
            "detail_total_base_ratio": "Total % Base/Oper",
            "detail_total_icms_value": "Total Valor ICMS",
            "detail_total_effective_rate": "Total % Aliq Efetiva",
        }
        detail_widths = {
            "reason": 300,
            "register": 80,
            "cst": 70,
            "cfop": 80,
            "rate": 70,
            "detail_total_rows": 90,
            "detail_total_operation_value": 130,
            "detail_total_base_icms": 130,
            "detail_total_base_gap": 130,
            "detail_total_base_ratio": 110,
            "detail_total_icms_value": 130,
            "detail_total_effective_rate": 115,
        }
        for period in periods:
            period_key = period.replace("/", "_")
            detail_headings[f"{period_key}_detail_rows"] = f"{period} Linhas"
            detail_headings[f"{period_key}_detail_operation_value"] = f"{period} Valor Oper."
            detail_headings[f"{period_key}_detail_base_icms"] = f"{period} Base ICMS"
            detail_headings[f"{period_key}_detail_base_gap"] = f"{period} Perda Base"
            detail_headings[f"{period_key}_detail_base_ratio"] = f"{period} % Base/Oper"
            detail_headings[f"{period_key}_detail_icms_value"] = f"{period} Valor ICMS"
            detail_headings[f"{period_key}_detail_effective_rate"] = f"{period} % Aliq Efetiva"
            detail_widths[f"{period_key}_detail_rows"] = 80
            detail_widths[f"{period_key}_detail_operation_value"] = 120
            detail_widths[f"{period_key}_detail_base_icms"] = 120
            detail_widths[f"{period_key}_detail_base_gap"] = 120
            detail_widths[f"{period_key}_detail_base_ratio"] = 105
            detail_widths[f"{period_key}_detail_icms_value"] = 120
            detail_widths[f"{period_key}_detail_effective_rate"] = 110

        detail_sort_state = {"column": "detail_total_base_gap", "reverse": True}
        detail_tree_rows_state = list(detail_display_rows)
        detail_fixed_columns = ("reason", "register", "cst", "cfop", "rate")
        detail_total_columns = (
            "detail_total_rows",
            "detail_total_operation_value",
            "detail_total_base_icms",
            "detail_total_base_gap",
            "detail_total_base_ratio",
            "detail_total_icms_value",
            "detail_total_effective_rate",
        )

        def detail_comparison_sort_value(row: dict[str, object], column_id: str) -> object:
            values = list(row.get("values", []))
            if column_id not in detail_columns:
                return ""
            index = detail_columns.index(column_id)
            value = values[index] if index < len(values) else ""
            if column_id in {"reason", "register", "cst", "cfop"}:
                return normalize_text(str(value))
            if isinstance(value, Decimal):
                return value
            if isinstance(value, (int, float)):
                return value
            return normalize_text(str(value))

        def render_detail_comparison_tree() -> None:
            for item_id in detail_tree.get_children():
                detail_tree.delete(item_id)
            selected_period_filter = str(selected_period_state["period"]).strip()
            visible_rows = list(detail_tree_rows_state)
            if selected_period_filter:
                row_metric_columns = detail_period_metric_columns.get(selected_period_filter, ())
                row_count_index = detail_columns.index(row_metric_columns[0]) if row_metric_columns else -1
                if row_count_index >= 0:
                    visible_rows = [
                        row
                        for row in visible_rows
                        if int(row.get("values", [])[row_count_index] or 0) > 0
                    ]
            sorted_rows = sorted(
                visible_rows,
                key=lambda row: detail_comparison_sort_value(row, str(detail_sort_state["column"])),
                reverse=bool(detail_sort_state["reverse"]),
            )
            for row_index, row in enumerate(sorted_rows):
                detail_tree.insert(
                    "",
                    END,
                    iid=str(row_index),
                    values=tuple(serialize_value_for_clipboard(value) for value in row["values"]),
                )
            detail_tree._comparison_rows = sorted_rows

        def update_detail_display_columns() -> None:
            selected_period_filter = str(selected_period_state["period"]).strip()
            if selected_period_filter and selected_period_filter in detail_period_metric_columns:
                detail_tree.configure(
                    displaycolumns=(
                        *detail_fixed_columns,
                        *detail_period_metric_columns[selected_period_filter],
                        *detail_total_columns,
                    )
                )
            else:
                detail_tree.configure(displaycolumns=detail_columns)

        def sort_detail_comparison_tree_by(column_id: str) -> None:
            if detail_sort_state["column"] == column_id:
                detail_sort_state["reverse"] = not detail_sort_state["reverse"]
            else:
                detail_sort_state["column"] = column_id
                detail_sort_state["reverse"] = False
            render_detail_comparison_tree()

        for column_id in detail_columns:
            detail_tree.heading(column_id, text=detail_headings[column_id], command=lambda current=column_id: sort_detail_comparison_tree_by(current))
            detail_tree.column(column_id, width=detail_widths[column_id], anchor="center")
        detail_tree.grid(row=0, column=0, sticky="nsew")
        detail_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(detail_tree))
        detail_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, detail_tree))

        def handle_detail_comparison_double_click(event: object) -> None:
            row_id = detail_tree.identify_row(getattr(event, "y", 0))
            column_token = detail_tree.identify_column(getattr(event, "x", 0))
            if not row_id or not column_token.startswith("#"):
                return
            try:
                row_index = int(row_id)
                column_index = int(column_token[1:]) - 1
            except ValueError:
                return
            filtered_rows = getattr(detail_tree, "_comparison_rows", detail_tree_rows_state)
            if row_index < 0 or row_index >= len(filtered_rows) or column_index < 0 or column_index >= len(detail_columns):
                return
            column_id = detail_columns[column_index]
            selected_period = next(
                (
                    period
                    for period, metric_columns in detail_period_metric_columns.items()
                    if column_id in {metric_columns[1], metric_columns[2]}
                ),
                "",
            )
            if not selected_period:
                return
            selected_row = filtered_rows[row_index]
            target_detail_row = next(
                (
                    detail_row
                    for detail_row in detail_rows
                    if str(detail_row.get("period", "")).strip() == selected_period
                    and str(detail_row.get("reason", "")).strip() == str(selected_row.get("reason", "")).strip()
                    and str(detail_row.get("source_register", "")).strip() == str(selected_row.get("source_register", "")).strip()
                    and str(detail_row.get("cst_icms", "")).strip() == str(selected_row.get("cst_icms", "")).strip()
                    and str(detail_row.get("cfop", "")).strip() == str(selected_row.get("cfop", "")).strip()
                    and Decimal(detail_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
                    == Decimal(selected_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
                ),
                None,
            )
            if target_detail_row is None:
                messagebox.showwarning(diagnostic_caption, "Nao ha itens para este agrupamento no periodo selecionado.")
                return
            self.open_credit_diagnostic_grouped_rows_popup(target_detail_row, list(target_detail_row.get("grouped_rows", [])))

        detail_tree.bind("<Double-1>", handle_detail_comparison_double_click)

        detail_scroll_y = ttk.Scrollbar(detail_box, orient="vertical", command=detail_tree.yview)
        detail_scroll_y.grid(row=0, column=1, sticky="ns")
        detail_scroll_x = ttk.Scrollbar(detail_box, orient="horizontal", command=detail_tree.xview)
        detail_scroll_x.grid(row=1, column=0, sticky="ew")
        detail_tree.configure(yscrollcommand=detail_scroll_y.set, xscrollcommand=detail_scroll_x.set)

        update_detail_display_columns()
        render_detail_comparison_tree()

        def render_period_band() -> None:
            for label in period_labels.values():
                label.destroy()
            period_labels.clear()
            selected_period_filter = str(selected_period_state["period"]).strip()
            for period_index, period in enumerate(periods):
                is_selected = period == selected_period_filter
                label = tk.Label(
                    period_band,
                    text=period,
                    bg="#7cc6f2" if is_selected else "#d9eaf7" if period_index % 2 == 0 else "#fbe7c6",
                    fg="#08324d" if is_selected else "#184c73",
                    padx=8,
                    pady=3,
                    relief="solid",
                    bd=2 if is_selected else 1,
                    cursor="hand2",
                    font=("Segoe UI", 9, "bold"),
                )
                label.pack(side=LEFT, padx=(0, 6))
                label.bind(
                    "<Button-1>",
                    lambda _event, current_period=period: toggle_selected_period(current_period),
                )
                period_labels[period] = label

        def toggle_selected_period(period: str) -> None:
            current_period = str(selected_period_state["period"]).strip()
            selected_period_state["period"] = "" if current_period == str(period).strip() else str(period).strip()
            render_period_band()
            update_summary_display_columns()
            update_detail_display_columns()
            render_comparison_tree()
            render_detail_comparison_tree()
            refresh_comparison_footer()

        render_period_band()

        footer = ttk.Frame(container)
        footer.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        total_reasons_var = StringVar(value="Motivos: 0")
        total_groups_var = StringVar(value="Agrupamentos fiscais: 0")
        total_rows_var = StringVar(value="Linhas: 0")
        total_operation_var = StringVar(value="Valor Operacao: 0,00")
        total_base_var = StringVar(value="Base ICMS: 0,00")
        total_gap_var = StringVar(value="Perda de Base: 0,00")
        total_ratio_var = StringVar(value="% Base/Oper: 0,00")
        total_icms_var = StringVar(value="Valor ICMS: 0,00")
        ttk.Label(footer, textvariable=total_reasons_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_groups_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_operation_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_gap_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        def refresh_comparison_footer() -> None:
            visible_summary_rows = list(getattr(tree, "_comparison_rows", []))
            visible_detail_rows = list(getattr(detail_tree, "_comparison_rows", []))
            selected_period_filter = str(selected_period_state["period"]).strip()
            if selected_period_filter and selected_period_filter in period_metric_columns:
                metric_columns = period_metric_columns[selected_period_filter]
                row_count_index = columns.index(metric_columns[0])
                operation_index = columns.index(metric_columns[1])
                base_index = columns.index(metric_columns[2])
                gap_index = columns.index(metric_columns[3])
                icms_index = columns.index(metric_columns[5])
            else:
                row_count_index = len(columns) - 6
                operation_index = len(columns) - 5
                base_index = len(columns) - 4
                gap_index = len(columns) - 3
                icms_index = len(columns) - 1

            total_rows = sum((int(row["values"][row_count_index]) for row in visible_summary_rows), 0)
            total_operation_value = sum((Decimal(row["values"][operation_index]) for row in visible_summary_rows), Decimal("0"))
            total_base_icms = sum((Decimal(row["values"][base_index]) for row in visible_summary_rows), Decimal("0"))
            total_base_gap = sum((Decimal(row["values"][gap_index]) for row in visible_summary_rows), Decimal("0"))
            total_icms_value = sum((Decimal(row["values"][icms_index]) for row in visible_summary_rows), Decimal("0"))
            total_base_ratio = (
                (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
                if total_operation_value > 0
                else Decimal("0.00")
            )
            total_reasons_var.set(f"Motivos: {len(visible_summary_rows)}")
            total_groups_var.set(f"Agrupamentos fiscais: {len(visible_detail_rows)}")
            total_rows_var.set(f"Linhas: {total_rows}")
            total_operation_var.set(f"Valor Operacao: {format_decimal_sped(total_operation_value)}")
            total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base_icms)}")
            total_gap_var.set(f"Perda de Base: {format_decimal_sped(total_base_gap)}")
            total_ratio_var.set(f"% Base/Oper: {format_decimal_sped(total_base_ratio)}")
            total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms_value)}")

        refresh_comparison_footer()

        export_actions = ttk.Frame(container)
        export_actions.grid(row=6, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Comparacao Excel",
            style="Secondary.TButton",
            command=lambda: self.export_credit_diagnostic_period_comparison_dataset(
                periods,
                export_headers,
                export_rows,
                detail_export_headers,
                detail_export_rows,
                detail_rows,
                "xlsx",
                diagnostic_caption,
                f"comparacao_diagnostico_{'debito' if is_saida else 'credito'}",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Comparacao CSV",
            style="Secondary.TButton",
            command=lambda: self.export_credit_diagnostic_period_comparison_dataset(
                periods,
                export_headers,
                export_rows,
                detail_export_headers,
                detail_export_rows,
                detail_rows,
                "csv",
                diagnostic_caption,
                f"comparacao_diagnostico_{'debito' if is_saida else 'credito'}",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        self.log_message(f"Popup de {diagnostic_caption.lower()} aberto.")

    def export_credit_diagnostic_period_comparison_dataset(
        self,
        periods: list[str],
        headers: list[str],
        rows: list[list[object]],
        detail_headers: list[str],
        detail_comparison_rows: list[list[object]],
        detail_rows: list[dict[str, object]],
        output_type: str,
        dialog_title: str = "Comparacao diagnostico de credito",
        initial_filename_base: str = "comparacao_diagnostico_credito",
    ) -> None:
        if not rows:
            messagebox.showwarning("Exportar comparacao", "Nao ha dados para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title=f"Exportar {dialog_title.lower()}",
            defaultextension=suffix,
            initialfile=f"{initial_filename_base}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                workbook_sheets: list[tuple[str, list[str], list[list[object]]]] = [
                    ("Resumo Comparativo", headers, rows),
                ]
                if detail_comparison_rows:
                    workbook_sheets.append(("Detalhamento Comparativo", detail_headers, detail_comparison_rows))
                grouped_item_headers = [
                    "Periodo",
                    "Motivo",
                    "Registro",
                    "CST",
                    "CFOP",
                    "Aliq",
                    "Aliq Efetiva",
                    "Linha SPED",
                    "Valor Operacao",
                    "Base ICMS",
                    "Perda de Base",
                    "% Base/Oper",
                    "Valor ICMS",
                    "Linha Original",
                ]
                product_item_headers = [
                    "Periodo",
                    "Motivo",
                    "Registro",
                    "CST",
                    "CFOP",
                    "Codigo",
                    "Produto",
                    "NCM",
                    "Aliq",
                    "Aliq Efetiva",
                    "Linhas",
                    "Valor Operacao",
                    "Base ICMS",
                    "% Base/Oper",
                    "Valor ICMS",
                ]
                all_grouped_rows: list[list[object]] = []
                all_product_rows, product_footer_rows = build_credit_diagnostic_product_sheet_payload(detail_rows)
                for reason in sorted({str(row.get("reason", "")).strip() for row in detail_rows if str(row.get("reason", "")).strip()}, key=normalize_text):
                    reason_sheet_rows: list[list[object]] = []
                    for detail_row in detail_rows:
                        if str(detail_row.get("reason", "")).strip() != reason:
                            continue
                        for grouped_row in detail_row.get("grouped_rows", []):
                            operation_value = Decimal(grouped_row.get("total_operation_value", Decimal("0")))
                            base_icms = Decimal(grouped_row.get("base_icms", Decimal("0")))
                            icms_value = Decimal(grouped_row.get("icms_value", Decimal("0")))
                            base_gap = max(Decimal("0"), operation_value - base_icms)
                            base_ratio = (
                                (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01"))
                                if operation_value > 0
                                else Decimal("0.00")
                            )
                            reason_sheet_rows.append(
                                [
                                    detail_row.get("period", ""),
                                    reason,
                                    detail_row.get("source_register", ""),
                                    detail_row.get("cst_icms", ""),
                                    detail_row.get("cfop", ""),
                                    Decimal(grouped_row.get("icms_rate", Decimal("0"))),
                                    compute_display_icms_rate(
                                        Decimal(grouped_row.get("icms_rate", Decimal("0"))),
                                        operation_value,
                                        base_icms,
                                        icms_value,
                                    ),
                                    grouped_row.get("line_number", ""),
                                    operation_value,
                                    base_icms,
                                    base_gap,
                                    base_ratio,
                                    icms_value,
                                    grouped_row.get("raw_line", ""),
                                ]
                            )
                            all_grouped_rows.append(reason_sheet_rows[-1])
                if all_grouped_rows:
                    workbook_sheets.append(("Itens Detalhados", grouped_item_headers, all_grouped_rows))
                if all_product_rows:
                    workbook_sheets.append(
                        (
                            "Produtos Detalhados",
                            product_item_headers,
                            all_product_rows,
                            {"footer_rows": product_footer_rows},
                        )
                    )
                for reason in sorted({str(row.get("reason", "")).strip() for row in detail_rows if str(row.get("reason", "")).strip()}, key=normalize_text):
                    reason_sheet_rows: list[list[object]] = []
                    for detail_row in detail_rows:
                        if str(detail_row.get("reason", "")).strip() != reason:
                            continue
                        for grouped_row in detail_row.get("grouped_rows", []):
                            operation_value = Decimal(grouped_row.get("total_operation_value", Decimal("0")))
                            base_icms = Decimal(grouped_row.get("base_icms", Decimal("0")))
                            icms_value = Decimal(grouped_row.get("icms_value", Decimal("0")))
                            base_gap = max(Decimal("0"), operation_value - base_icms)
                            base_ratio = (
                                (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01"))
                                if operation_value > 0
                                else Decimal("0.00")
                            )
                            reason_sheet_rows.append(
                                [
                                    detail_row.get("period", ""),
                                    reason,
                                    detail_row.get("source_register", ""),
                                    detail_row.get("cst_icms", ""),
                                    detail_row.get("cfop", ""),
                                    Decimal(grouped_row.get("icms_rate", Decimal("0"))),
                                    compute_display_icms_rate(
                                        Decimal(grouped_row.get("icms_rate", Decimal("0"))),
                                        operation_value,
                                        base_icms,
                                        icms_value,
                                    ),
                                    grouped_row.get("line_number", ""),
                                    operation_value,
                                    base_icms,
                                    base_gap,
                                    base_ratio,
                                    icms_value,
                                    grouped_row.get("raw_line", ""),
                                ]
                            )
                    if reason_sheet_rows:
                        workbook_sheets.append((f"Motivo {reason}", grouped_item_headers, reason_sheet_rows))
                write_simple_excel_workbook(output_path, workbook_sheets)
            self.log_message(f"{dialog_title} exportada em: {output_path}")
            messagebox.showinfo("Exportar comparacao", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar {dialog_title.lower()}: {exc}")
            messagebox.showerror("Exportar comparacao", str(exc))

    def handle_credit_diagnostic_detail_double_click(
        self,
        event: object,
        detail_tree: ttk.Treeview,
        detail_rows: list[dict[str, object]],
    ) -> None:
        row_id = detail_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(detail_rows):
            return
        detail_row = detail_rows[row_index]
        grouped_rows = detail_row.get("grouped_rows", [])
        if not isinstance(grouped_rows, list) or not grouped_rows:
            messagebox.showwarning("Detalhamento fiscal", "Nao ha linhas detalhadas para este agrupamento.")
            return
        self.open_credit_diagnostic_grouped_rows_popup(detail_row, grouped_rows)

    def handle_credit_diagnostic_summary_double_click(
        self,
        event: object,
        summary_tree: ttk.Treeview,
        detail_rows: list[dict[str, object]],
    ) -> None:
        row_id = summary_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        summary_values = summary_tree.item(row_id, "values")
        if not summary_values or len(summary_values) < 2:
            return
        selected_period = str(summary_values[0]).strip()
        selected_reason = str(summary_values[1]).strip()
        merged_detail = merge_credit_diagnostic_detail_rows(detail_rows, selected_reason, selected_period)
        if not merged_detail:
            messagebox.showwarning("Diagnostico de credito", "Nao ha itens para este motivo no periodo selecionado.")
            return
        self.open_credit_diagnostic_grouped_rows_popup(merged_detail, list(merged_detail.get("grouped_rows", [])))

    def handle_credit_diagnostic_product_double_click(
        self,
        event: object,
        product_tree: ttk.Treeview,
        product_rows: list[dict[str, object]],
        detail_row: dict[str, object],
    ) -> None:
        row_id = product_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        row_index = product_tree.index(row_id)
        sorted_rows = getattr(product_tree, "_sorted_product_rows", product_rows)
        if row_index < 0 or row_index >= len(sorted_rows):
            return
        product_row = sorted_rows[row_index]
        self.open_launch_origin_details_popup(
            list(product_row.get("launch_details", [])),
            str(product_row.get("code", "")).strip() or str(detail_row.get("code", "")).strip() or "Produto",
            str(product_row.get("description", "")).strip() or str(detail_row.get("reason", "")).strip(),
            str(product_row.get("period", "")).strip() or str(detail_row.get("period", "")).strip(),
        )

    def build_grouped_popup_product_rows_fallback(
        self,
        detail_row: dict[str, object],
        grouped_rows: list[dict[str, object]] | None = None,
    ) -> tuple[list[dict[str, object]], dict[str, object]]:
        def infer_period_from_value(value: object) -> str:
            digits = "".join(char for char in str(value or "").strip() if char.isdigit())
            if len(digits) >= 6:
                if digits[:4].isdigit() and int(digits[:4]) >= 1900:
                    return f"{digits[4:6]}/{digits[:4]}"
                if len(digits) >= 8:
                    return f"{digits[2:4]}/{digits[4:8]}"
            parsed_date = parse_sped_document_date(value)
            if parsed_date is not None:
                return parsed_date.strftime("%m/%Y")
            return ""

        operation_type = normalize_operation_type(str(detail_row.get("operation_type", "")).strip())
        source_view = "sales_consult" if operation_type == "Saida" else "consult"
        source_rows = (
            list(getattr(self, "sales_consult_filtered_rows", []))
            + list(getattr(self, "sales_consult_comparison_rows", []))
            if operation_type == "Saida"
            else list(getattr(self, "consult_filtered_rows", []))
            + list(getattr(self, "consult_comparison_rows", []))
        )
        source_details: list[dict[str, object]] = []
        diagnostic_info: dict[str, object] = {
            "operation_type": operation_type,
            "period": "",
            "target_cfops": [],
            "target_keys_exact": 0,
            "target_keys_cst_cfop": 0,
            "source_details_count": 0,
            "source_details_period_count": 0,
            "xml_sources_count": 0,
            "xml_documents_count": 0,
            "xml_items_count": 0,
            "matched_level": "",
            "matched_candidates": 0,
            "product_count": 0,
        }
        selected_rule_profile = ""
        runtime_rules: list[dict[str, object]] = []
        try:
            context = self.load_current_popup_context(source_view)
            selected_rule_profile = str(context.get("selected_rule_profile", "")).strip()
            runtime_rules = list(context.get("runtime_rules", []))
            source_details.extend(list(context.get("all_details", [])))
            source_details.extend(list(context.get("filtered_details", [])))
        except Exception:
            pass
        for row in source_rows:
            row_period = str(row.get("period", "")).strip()
            for detail in row.get("launch_details", []):
                source_details.append(
                    {
                        **detail,
                        "period": row_period,
                        "code": str(detail.get("code", "")).strip() or str(row.get("code", "")).strip(),
                        "description": str(detail.get("description", "")).strip() or str(row.get("description", "")).strip(),
                        "ncm": str(detail.get("ncm", "")).strip() or str(row.get("ncm", "")).strip(),
                        "operation_type": str(detail.get("operation_type", "")).strip() or operation_type,
                    }
                )
        period = str(detail_row.get("period", "")).strip()
        diagnostic_info["period"] = period
        target_keys_exact: set[tuple[str, str, Decimal]] = set()
        target_keys_cst_cfop: set[tuple[str, str]] = set()
        target_cfops: set[str] = set()
        grouped_rows = list(grouped_rows or [])
        if grouped_rows:
            for grouped_row in grouped_rows:
                grouped_cst = normalize_cst_icms_for_sped(str(grouped_row.get("cst_icms", "")).strip())
                grouped_cfop = str(grouped_row.get("cfop", "")).strip()
                grouped_rate = Decimal(grouped_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
                if grouped_cfop:
                    target_cfops.add(grouped_cfop)
                if grouped_cst and grouped_cfop:
                    target_keys_cst_cfop.add((grouped_cst, grouped_cfop))
                    target_keys_exact.add((grouped_cst, grouped_cfop, grouped_rate))
        else:
            target_cst = normalize_cst_icms_for_sped(str(detail_row.get("cst_icms", "")).strip())
            target_cfop = str(detail_row.get("cfop", "")).strip()
            target_rate = Decimal(detail_row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
            if target_cfop:
                target_cfops.add(target_cfop)
            if target_cst and target_cfop:
                target_keys_cst_cfop.add((target_cst, target_cfop))
                target_keys_exact.add((target_cst, target_cfop, target_rate))
        diagnostic_info["target_cfops"] = sorted(target_cfops)
        diagnostic_info["target_keys_exact"] = len(target_keys_exact)
        diagnostic_info["target_keys_cst_cfop"] = len(target_keys_cst_cfop)
        diagnostic_info["source_details_count"] = len(source_details)
        fallback_levels = (
            "period_cst_cfop_rate",
            "period_cst_cfop",
            "period_cfop",
            "cst_cfop_rate",
            "cst_cfop",
            "cfop",
        )
        candidate_details_by_level: dict[str, list[dict[str, object]]] = {level: [] for level in fallback_levels}
        seen_detail_keys: dict[str, set[tuple[str, str, str, str, str, str]]] = {level: set() for level in fallback_levels}

        for detail in source_details:
            detail_period = str(detail.get("period", "")).strip()
            if detail_period != period:
                continue
            diagnostic_info["source_details_period_count"] = int(diagnostic_info["source_details_period_count"]) + 1
            if normalize_operation_type(str(detail.get("operation_type", "")).strip()) != operation_type:
                continue
            adjusted_detail = apply_sped_icms_consistency_rules(detail, selected_rule_profile, runtime_rules)
            detail_cst = normalize_cst_icms_for_sped(str(adjusted_detail.get("cst_icms", "")).strip())
            detail_cfop = str(adjusted_detail.get("cfop", "")).strip()
            detail_rate = Decimal(adjusted_detail.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
            detail_identity = (
                normalize_document_key(str(adjusted_detail.get("document_key", ""))) or str(adjusted_detail.get("document_number", "")).strip(),
                str(adjusted_detail.get("item_number", "")).strip(),
                str(adjusted_detail.get("code", "")).strip(),
                detail_cfop,
                detail_cst,
                format_decimal_sped(detail_rate),
            )
            if (detail_cst, detail_cfop, detail_rate) in target_keys_exact:
                if detail_identity not in seen_detail_keys["period_cst_cfop_rate"]:
                    seen_detail_keys["period_cst_cfop_rate"].add(detail_identity)
                    candidate_details_by_level["period_cst_cfop_rate"].append(adjusted_detail)
            if (detail_cst, detail_cfop) in target_keys_cst_cfop:
                if detail_identity not in seen_detail_keys["period_cst_cfop"]:
                    seen_detail_keys["period_cst_cfop"].add(detail_identity)
                    candidate_details_by_level["period_cst_cfop"].append(adjusted_detail)
            if detail_cfop in target_cfops:
                if detail_identity not in seen_detail_keys["period_cfop"]:
                    seen_detail_keys["period_cfop"].add(detail_identity)
                    candidate_details_by_level["period_cfop"].append(adjusted_detail)
            if (detail_cst, detail_cfop, detail_rate) in target_keys_exact:
                if detail_identity not in seen_detail_keys["cst_cfop_rate"]:
                    seen_detail_keys["cst_cfop_rate"].add(detail_identity)
                    candidate_details_by_level["cst_cfop_rate"].append(adjusted_detail)
            if (detail_cst, detail_cfop) in target_keys_cst_cfop:
                if detail_identity not in seen_detail_keys["cst_cfop"]:
                    seen_detail_keys["cst_cfop"].add(detail_identity)
                    candidate_details_by_level["cst_cfop"].append(adjusted_detail)
            if detail_cfop in target_cfops:
                if detail_identity not in seen_detail_keys["cfop"]:
                    seen_detail_keys["cfop"].add(detail_identity)
                    candidate_details_by_level["cfop"].append(adjusted_detail)

        selected_details: list[dict[str, object]] = []
        for level in fallback_levels:
            if candidate_details_by_level[level]:
                selected_details = candidate_details_by_level[level]
                diagnostic_info["matched_level"] = f"consulta:{level}"
                diagnostic_info["matched_candidates"] = len(selected_details)
                break

        if not selected_details and operation_type == "Saida":
            xml_sources = parse_selected_paths(self.sales_consult_xml_paths_var.get())
            diagnostic_info["xml_sources_count"] = len(xml_sources)
            xml_index = build_xml_fiscal_item_index(xml_sources)
            diagnostic_info["xml_documents_count"] = len(xml_index)
            diagnostic_info["xml_items_count"] = sum(len(items) for items in xml_index.values())
            xml_detail_candidates_by_level: dict[str, list[dict[str, object]]] = {level: [] for level in fallback_levels}
            xml_seen_detail_keys: dict[str, set[tuple[str, str, str, str, str, str]]] = {level: set() for level in fallback_levels}
            for document_items in xml_index.values():
                for xml_item in document_items:
                    detail_period = infer_period_from_value(xml_item.get("data_emissao", ""))
                    if detail_period != period:
                        continue
                    raw_detail_cst = compose_xml_icms_cst_for_sped(xml_item)
                    raw_detail_cfop = str(xml_item.get("cfop", "")).strip()
                    raw_detail_rate = Decimal(xml_item.get("aliquota_icms", Decimal("0"))).quantize(Decimal("0.01"))
                    xml_detail = {
                        "operation_type": "Saida",
                        "period": detail_period,
                        "document_number": str(xml_item.get("numero_nfce", "")).strip(),
                        "document_key": normalize_document_key(str(xml_item.get("chave_acesso", ""))),
                        "document_date": str(xml_item.get("data_emissao", "")).strip(),
                        "document_series": str(xml_item.get("serie", "")).strip(),
                        "document_model": str(xml_item.get("modelo", "")).strip() or "55",
                        "participant_code": "",
                        "participant_name": str(xml_item.get("emitente", "")).strip(),
                        "participant_tax_id": str(xml_item.get("cnpj_emitente", "")).strip(),
                        "item_number": str(xml_item.get("item", "")).strip(),
                        "code": str(xml_item.get("codigo", "")).strip() or str(xml_item.get("descricao", "")).strip(),
                        "description": str(xml_item.get("descricao", "")).strip(),
                        "ncm": str(xml_item.get("ncm", "")).strip(),
                        "cst_icms": raw_detail_cst,
                        "cfop": raw_detail_cfop,
                        "quantity": Decimal(xml_item.get("quantidade", Decimal("0"))),
                        "icms_rate": raw_detail_rate,
                        "sale_value": Decimal(xml_item.get("valor_operacao", xml_item.get("valor_produto", Decimal("0")))),
                        "base_icms": Decimal(xml_item.get("base_icms", Decimal("0"))),
                        "icms_value": Decimal(xml_item.get("valor_icms", Decimal("0"))),
                    }
                    xml_detail = apply_sped_icms_consistency_rules(xml_detail, selected_rule_profile, runtime_rules)
                    detail_cst = normalize_cst_icms_for_sped(str(xml_detail.get("cst_icms", "")).strip())
                    detail_cfop = str(xml_detail.get("cfop", "")).strip()
                    detail_rate = Decimal(xml_detail.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
                    detail_identity = (
                        normalize_document_key(str(xml_detail.get("document_key", ""))) or str(xml_detail.get("document_number", "")).strip(),
                        str(xml_detail.get("item_number", "")).strip(),
                        str(xml_detail.get("code", "")).strip(),
                        detail_cfop,
                        detail_cst,
                        format_decimal_sped(detail_rate),
                    )
                    if (detail_cst, detail_cfop, detail_rate) in target_keys_exact:
                        if detail_identity not in xml_seen_detail_keys["period_cst_cfop_rate"]:
                            xml_seen_detail_keys["period_cst_cfop_rate"].add(detail_identity)
                            xml_detail_candidates_by_level["period_cst_cfop_rate"].append(xml_detail)
                    if (detail_cst, detail_cfop) in target_keys_cst_cfop:
                        if detail_identity not in xml_seen_detail_keys["period_cst_cfop"]:
                            xml_seen_detail_keys["period_cst_cfop"].add(detail_identity)
                            xml_detail_candidates_by_level["period_cst_cfop"].append(xml_detail)
                    if detail_cfop in target_cfops:
                        if detail_identity not in xml_seen_detail_keys["period_cfop"]:
                            xml_seen_detail_keys["period_cfop"].add(detail_identity)
                            xml_detail_candidates_by_level["period_cfop"].append(xml_detail)
                    if (detail_cst, detail_cfop, detail_rate) in target_keys_exact:
                        if detail_identity not in xml_seen_detail_keys["cst_cfop_rate"]:
                            xml_seen_detail_keys["cst_cfop_rate"].add(detail_identity)
                            xml_detail_candidates_by_level["cst_cfop_rate"].append(xml_detail)
                    if (detail_cst, detail_cfop) in target_keys_cst_cfop:
                        if detail_identity not in xml_seen_detail_keys["cst_cfop"]:
                            xml_seen_detail_keys["cst_cfop"].add(detail_identity)
                            xml_detail_candidates_by_level["cst_cfop"].append(xml_detail)
                    if detail_cfop in target_cfops:
                        if detail_identity not in xml_seen_detail_keys["cfop"]:
                            xml_seen_detail_keys["cfop"].add(detail_identity)
                            xml_detail_candidates_by_level["cfop"].append(xml_detail)
            for level in fallback_levels:
                if xml_detail_candidates_by_level[level]:
                    selected_details = xml_detail_candidates_by_level[level]
                    diagnostic_info["matched_level"] = f"xml:{level}"
                    diagnostic_info["matched_candidates"] = len(selected_details)
                    break

        product_buckets: dict[tuple[str, str, str], dict[str, object]] = {}
        for detail in selected_details:
            code = str(detail.get("code", "")).strip()
            description = str(detail.get("description", "")).strip()
            ncm = str(detail.get("ncm", "")).strip()
            document_number = str(detail.get("document_number", "")).strip()
            item_number = str(detail.get("item_number", "")).strip()
            document_key = normalize_document_key(str(detail.get("document_key", "")))
            if not code:
                code = document_number or document_key[:12] or "SEM-CODIGO"
            if not description:
                if document_number or item_number:
                    description = f"Documento {document_number or document_key[:12]} Item {item_number or '?'}"
                else:
                    description = "Item sem descricao"
            detail_rate = Decimal(detail.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01"))
            product_key = (code, description, ncm)
            bucket = product_buckets.setdefault(
                product_key,
                {
                    "code": code,
                    "description": description,
                    "ncm": ncm,
                    "period": period,
                    "icms_rate": detail_rate,
                    "row_count": 0,
                    "document_keys": set(),
                    "launch_details": [],
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            sale_value = Decimal(detail.get("sale_value", Decimal("0")))
            base_icms = Decimal(detail.get("base_icms", Decimal("0")))
            icms_value = Decimal(detail.get("icms_value", Decimal("0")))
            bucket["row_count"] += 1
            document_identity = normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            if document_identity:
                bucket["document_keys"].add(document_identity)
            bucket["launch_details"].append(dict(detail))
            bucket["sale_value"] += sale_value
            bucket["base_icms"] += base_icms
            bucket["icms_value"] += icms_value

        product_rows = sorted(
            product_buckets.values(),
            key=lambda row: (Decimal(row.get("sale_value", Decimal("0"))), str(row.get("code", "")), str(row.get("description", ""))),
            reverse=True,
        )
        diagnostic_info["product_count"] = len(product_rows)
        return product_rows, diagnostic_info

    def open_credit_diagnostic_grouped_rows_popup(
        self,
        detail_row: dict[str, object],
        grouped_rows: list[dict[str, object]],
    ) -> None:
        grouped_rows_state = list(grouped_rows)
        product_rows_state = list(detail_row.get("product_rows", []))
        product_diagnostic_info: dict[str, object] | None = None
        if not product_rows_state:
            product_rows_state, product_diagnostic_info = self.build_grouped_popup_product_rows_fallback(detail_row, grouped_rows_state)

        dialog = Toplevel(self.root)
        dialog.title("Linhas do Agrupamento Fiscal")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1540, 820, 1260, 560, margin_y=190)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)
        grouped_sort_state = {"column": "base_gap", "reverse": True}
        product_sort_state = {"column": "sale_value", "reverse": True}

        ttk.Label(
            container,
            text=(
                f"Periodo {detail_row.get('period', '')} | {detail_row.get('reason', '')} | Registro {detail_row.get('source_register', '')} | "
                f"CST {detail_row.get('cst_icms', '')} | CFOP {detail_row.get('cfop', '')} | "
                f"Aliq {format_decimal_sped(Decimal(detail_row.get('icms_rate', Decimal('0'))))}"
            ),
            style="Title.TLabel",
            wraplength=1400,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=(
                f"{len(grouped_rows)} linha(s) que compoem o agrupamento selecionado. "
                "A area superior prioriza os produtos; abaixo ficam as linhas fiscais do SPED."
            ),
            wraplength=1400,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))
        if product_diagnostic_info is not None and not product_rows_state:
            ttk.Label(
                container,
                text=(
                    f"Diagnostico itens: oper={product_diagnostic_info.get('operation_type', '')} "
                    f"periodo={product_diagnostic_info.get('period', '')} "
                    f"cfops={', '.join(product_diagnostic_info.get('target_cfops', [])) or '-'} "
                    f"detalhes={product_diagnostic_info.get('source_details_count', 0)} "
                    f"detalhes_periodo={product_diagnostic_info.get('source_details_period_count', 0)} "
                    f"xml_fontes={product_diagnostic_info.get('xml_sources_count', 0)} "
                    f"xml_docs={product_diagnostic_info.get('xml_documents_count', 0)} "
                    f"xml_itens={product_diagnostic_info.get('xml_items_count', 0)} "
                    f"camada={product_diagnostic_info.get('matched_level', '-') or '-'} "
                    f"candidatos={product_diagnostic_info.get('matched_candidates', 0)}"
                ),
                wraplength=1400,
                foreground="#8a6d3b",
            ).grid(row=2, column=0, sticky="w", pady=(0, 10))

        export_actions = ttk.Frame(container)
        export_actions.grid(row=2 if product_diagnostic_info is None or product_rows_state else 3, column=0, sticky="w", pady=(0, 10))
        ttk.Button(
            export_actions,
            text="Exportar Excel",
            style="Secondary.TButton",
            command=lambda: self.export_credit_diagnostic_grouped_popup(detail_row, grouped_rows_state, product_rows_state),
        ).pack(side=LEFT)

        paned = ttk.Panedwindow(container, orient="vertical")
        paned.grid(row=4 if product_diagnostic_info is not None and not product_rows_state else 3, column=0, sticky="nsew")

        product_box = ttk.LabelFrame(paned, text="Produtos e NCM Relacionados", padding=10)
        product_box.columnconfigure(0, weight=1)
        product_box.rowconfigure(0, weight=1)
        paned.add(product_box, weight=3)

        line_box = ttk.LabelFrame(paned, text="Linhas do SPED do Agrupamento", padding=10)
        line_box.columnconfigure(0, weight=1)
        line_box.rowconfigure(0, weight=1)
        paned.add(line_box, weight=2)

        columns = (
            "period",
            "line_number",
            "register",
            "cst",
            "cfop",
            "rate",
            "effective_rate",
            "operation_value",
            "base_icms",
            "base_gap",
            "base_ratio",
            "icms_value",
            "raw_line",
        )
        tree = ttk.Treeview(line_box, columns=columns, show="headings", height=18, selectmode="extended")
        headings = {
            "period": "Periodo",
            "line_number": "Linha SPED",
            "register": "Registro",
            "cst": "CST",
            "cfop": "CFOP",
            "rate": "Aliq",
            "effective_rate": "% Aliq Efetiva",
            "operation_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "base_gap": "Perda de Base",
            "base_ratio": "% Base/Oper",
            "icms_value": "Valor ICMS",
            "raw_line": "Linha Original",
        }
        widths = {
            "period": 90,
            "line_number": 90,
            "register": 70,
            "cst": 70,
            "cfop": 80,
            "rate": 80,
            "effective_rate": 110,
            "operation_value": 120,
            "base_icms": 120,
            "base_gap": 120,
            "base_ratio": 100,
            "icms_value": 120,
            "raw_line": 900,
        }
        def grouped_sort_value(row: dict[str, object], column_id: str) -> object:
            if column_id in {"line_number"}:
                return int(row.get(column_id, 0) or 0)
            if column_id in {"rate"}:
                return Decimal(row.get("icms_rate", Decimal("0")))
            if column_id in {"effective_rate"}:
                return compute_display_icms_rate(
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    Decimal(row.get("total_operation_value", Decimal("0"))),
                    Decimal(row.get("base_icms", Decimal("0"))),
                    Decimal(row.get("icms_value", Decimal("0"))),
                )
            if column_id in {"operation_value"}:
                return Decimal(row.get("total_operation_value", Decimal("0")))
            if column_id in {"base_icms"}:
                return Decimal(row.get("base_icms", Decimal("0")))
            if column_id in {"base_gap"}:
                return max(Decimal("0"), Decimal(row.get("total_operation_value", Decimal("0"))) - Decimal(row.get("base_icms", Decimal("0"))))
            if column_id in {"base_ratio"}:
                operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
                base_icms = Decimal(row.get("base_icms", Decimal("0")))
                return (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01")) if operation_value > 0 else Decimal("0.00")
            if column_id in {"icms_value"}:
                return Decimal(row.get("icms_value", Decimal("0")))
            if column_id == "period":
                return period_label_sort_key(row.get("period", ""))
            source_map = {
                "register": "source_register",
                "cst": "cst_icms",
                "cfop": "cfop",
                "raw_line": "raw_line",
            }
            return normalize_text(str(row.get(source_map.get(column_id, column_id), "")))

        def sort_grouped_tree_by(column_id: str) -> None:
            if grouped_sort_state["column"] == column_id:
                grouped_sort_state["reverse"] = not grouped_sort_state["reverse"]
            else:
                grouped_sort_state["column"] = column_id
                grouped_sort_state["reverse"] = False
            render_grouped_tree()

        for column_id in columns:
            tree.heading(column_id, text=headings[column_id], command=lambda current=column_id: sort_grouped_tree_by(current))
            tree.column(column_id, width=widths[column_id], anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        tree.bind(
            "<Double-1>",
            lambda _event: self.open_launch_origin_details_popup(
                list(detail_row.get("launch_details", [])),
                str(detail_row.get("source_register", "")).strip() or str(detail_row.get("code", "")).strip() or "Agrupamento Fiscal",
                f"{detail_row.get('reason', '')} | Registro {detail_row.get('source_register', '')} | CST {detail_row.get('cst_icms', '')} | CFOP {detail_row.get('cfop', '')}",
                str(detail_row.get("period", "")).strip(),
            ),
        )

        scroll_y = ttk.Scrollbar(line_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(line_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        def render_grouped_tree() -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)
            sorted_rows = sorted(
                grouped_rows_state,
                key=lambda row: grouped_sort_value(row, str(grouped_sort_state["column"])),
                reverse=bool(grouped_sort_state["reverse"]),
            )
            for row in sorted_rows:
                operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
                base_icms = Decimal(row.get("base_icms", Decimal("0")))
                icms_value = Decimal(row.get("icms_value", Decimal("0")))
                base_gap = max(Decimal("0"), operation_value - base_icms)
                effective_rate = compute_display_icms_rate(
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    operation_value,
                    base_icms,
                    icms_value,
                )
                base_ratio = (
                    (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01"))
                    if operation_value > 0
                    else Decimal("0.00")
                )
                tree.insert(
                    "",
                    END,
                    values=(
                        row.get("period", ""),
                        row.get("line_number", ""),
                        row.get("source_register", ""),
                        row.get("cst_icms", ""),
                        row.get("cfop", ""),
                        format_decimal_sped(Decimal(row.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(effective_rate),
                        format_decimal_sped(operation_value),
                        format_decimal_sped(base_icms),
                        format_decimal_sped(base_gap),
                        format_decimal_sped(base_ratio),
                        format_decimal_sped(icms_value),
                        row.get("raw_line", ""),
                    ),
                )

        product_columns = ("period", "code", "description", "ncm", "rate", "effective_rate", "rows", "sale_value", "base_icms", "base_ratio", "icms_value")
        product_tree = ttk.Treeview(product_box, columns=product_columns, show="headings", height=14, selectmode="extended")
        product_headings = {
            "period": "Periodo",
            "code": "Codigo",
            "description": "Produto",
            "ncm": "NCM",
            "rate": "% Aliq",
            "effective_rate": "% Aliq Efetiva",
            "rows": "Linhas",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "base_ratio": "% Base/Oper",
            "icms_value": "Valor ICMS",
        }
        product_widths = {
            "period": 90,
            "code": 110,
            "description": 320,
            "ncm": 110,
            "rate": 80,
            "effective_rate": 110,
            "rows": 70,
            "sale_value": 120,
            "base_icms": 120,
            "base_ratio": 100,
            "icms_value": 120,
        }
        def product_sort_value(row: dict[str, object], column_id: str) -> object:
            if column_id == "period":
                return period_label_sort_key(row.get("period", ""))
            if column_id == "rate":
                return Decimal(row.get("icms_rate", Decimal("0")))
            if column_id == "effective_rate":
                return compute_display_icms_rate(
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    Decimal(row.get("sale_value", Decimal("0"))),
                    Decimal(row.get("base_icms", Decimal("0"))),
                    Decimal(row.get("icms_value", Decimal("0"))),
                )
            if column_id == "rows":
                return int(row.get("row_count", 0) or 0)
            if column_id == "sale_value":
                return Decimal(row.get("sale_value", Decimal("0")))
            if column_id == "base_icms":
                return Decimal(row.get("base_icms", Decimal("0")))
            if column_id == "base_ratio":
                sale_value = Decimal(row.get("sale_value", Decimal("0")))
                base_icms = Decimal(row.get("base_icms", Decimal("0")))
                return (base_icms * Decimal("100") / sale_value).quantize(Decimal("0.01")) if sale_value > 0 else Decimal("0.00")
            if column_id == "icms_value":
                return Decimal(row.get("icms_value", Decimal("0")))
            source_map = {"code": "code", "description": "description", "ncm": "ncm"}
            return normalize_text(str(row.get(source_map.get(column_id, column_id), "")))

        def sort_product_tree_by(column_id: str) -> None:
            if product_sort_state["column"] == column_id:
                product_sort_state["reverse"] = not product_sort_state["reverse"]
            else:
                product_sort_state["column"] = column_id
                product_sort_state["reverse"] = False
            render_product_tree()

        for column_id in product_columns:
            product_tree.heading(column_id, text=product_headings[column_id], command=lambda current=column_id: sort_product_tree_by(current))
            product_tree.column(column_id, width=product_widths[column_id], anchor="center")
        product_tree.grid(row=0, column=0, sticky="nsew")
        product_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(product_tree))
        product_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, product_tree))
        product_tree.bind("<Double-1>", lambda event: self.handle_credit_diagnostic_product_double_click(event, product_tree, product_rows_state, detail_row))
        product_scroll_y = ttk.Scrollbar(product_box, orient="vertical", command=product_tree.yview)
        product_scroll_y.grid(row=0, column=1, sticky="ns")
        product_scroll_x = ttk.Scrollbar(product_box, orient="horizontal", command=product_tree.xview)
        product_scroll_x.grid(row=1, column=0, sticky="ew")
        product_tree.configure(yscrollcommand=product_scroll_y.set, xscrollcommand=product_scroll_x.set)

        def render_product_tree() -> None:
            for item_id in product_tree.get_children():
                product_tree.delete(item_id)
            if product_rows_state:
                sorted_rows = sorted(
                    product_rows_state,
                    key=lambda row: product_sort_value(row, str(product_sort_state["column"])),
                    reverse=bool(product_sort_state["reverse"]),
                )
                for product in sorted_rows:
                    product_sale_value = Decimal(product.get("sale_value", Decimal("0")))
                    product_base_icms = Decimal(product.get("base_icms", Decimal("0")))
                    product_base_ratio = (
                        (product_base_icms * Decimal("100") / product_sale_value).quantize(Decimal("0.01"))
                        if product_sale_value > 0
                        else Decimal("0.00")
                    )
                    product_effective_rate = compute_display_icms_rate(
                        Decimal(product.get("icms_rate", Decimal("0"))),
                        product_sale_value,
                        product_base_icms,
                        Decimal(product.get("icms_value", Decimal("0"))),
                    )
                    product_tree.insert(
                        "",
                        END,
                        values=(
                            product.get("period", ""),
                            product.get("code", ""),
                            product.get("description", ""),
                            product.get("ncm", ""),
                            format_decimal_sped(Decimal(product.get("icms_rate", Decimal("0")))),
                            format_decimal_sped(product_effective_rate),
                            product.get("row_count", 0),
                            format_decimal_sped(product_sale_value),
                            format_decimal_sped(product_base_icms),
                            format_decimal_sped(product_base_ratio),
                            format_decimal_sped(Decimal(product.get("icms_value", Decimal("0")))),
                        ),
                    )
                product_tree._sorted_product_rows = sorted_rows
            else:
                product_tree._sorted_product_rows = []
                product_tree.insert("", END, values=("", "", "Nenhum produto/NCM identificado para este agrupamento.", "", "", "", "", "", "", "", ""))

        render_product_tree()
        render_grouped_tree()

        total_operation_value = sum((Decimal(row.get("total_operation_value", Decimal("0"))) for row in grouped_rows_state), Decimal("0"))
        total_base_icms = sum((Decimal(row.get("base_icms", Decimal("0"))) for row in grouped_rows_state), Decimal("0"))
        total_base_gap = sum((max(Decimal("0"), Decimal(row.get("total_operation_value", Decimal("0"))) - Decimal(row.get("base_icms", Decimal("0")))) for row in grouped_rows_state), Decimal("0"))
        total_icms_value = sum((Decimal(row.get("icms_value", Decimal("0"))) for row in grouped_rows_state), Decimal("0"))
        total_base_ratio = (
            (total_base_icms * Decimal("100") / total_operation_value).quantize(Decimal("0.01"))
            if total_operation_value > 0
            else Decimal("0.00")
        )
        footer = ttk.Frame(container)
        footer.grid(row=5 if product_diagnostic_info is not None and not product_rows_state else 4, column=0, sticky="w", pady=(10, 0))
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(total_operation_value)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base_icms)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Perda de Base: {format_decimal_sped(total_base_gap)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"% Base/Oper: {format_decimal_sped(total_base_ratio)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(total_icms_value)}", style="Strong.TLabel").pack(side=LEFT)

    def export_credit_diagnostic_grouped_popup(
        self,
        detail_row: dict[str, object],
        grouped_rows: list[dict[str, object]],
        product_rows: list[dict[str, object]],
    ) -> None:
        if not grouped_rows:
            messagebox.showwarning("Exportar popup", "Nao ha linhas detalhadas para exportar.")
            return

        selected = filedialog.asksaveasfilename(
            title="Exportar popup",
            defaultextension=".xlsx",
            initialfile="diagnostico_credito_agrupamento.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        detail_headers = [
            "Periodo",
            "Linha SPED",
            "Registro",
            "CST",
            "CFOP",
            "Aliq",
            "% Aliq Efetiva",
            "Valor Operacao",
            "Base ICMS",
            "Perda de Base",
            "% Base/Oper",
            "Valor ICMS",
            "Linha Original",
        ]
        detail_export_rows: list[list[object]] = []
        for row in grouped_rows:
            operation_value = Decimal(row.get("total_operation_value", Decimal("0")))
            base_icms = Decimal(row.get("base_icms", Decimal("0")))
            icms_value = Decimal(row.get("icms_value", Decimal("0")))
            base_gap = max(Decimal("0"), operation_value - base_icms)
            effective_rate = compute_display_icms_rate(
                Decimal(row.get("icms_rate", Decimal("0"))),
                operation_value,
                base_icms,
                icms_value,
            )
            base_ratio = (
                (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01"))
                if operation_value > 0
                else Decimal("0.00")
            )
            detail_export_rows.append(
                [
                    row.get("period", ""),
                    row.get("line_number", ""),
                    row.get("source_register", ""),
                    row.get("cst_icms", ""),
                    row.get("cfop", ""),
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    effective_rate,
                    operation_value,
                    base_icms,
                    base_gap,
                    base_ratio,
                    icms_value,
                    row.get("raw_line", ""),
                ]
            )

        product_headers = [
            "Periodo",
            "Codigo",
            "Produto",
            "NCM",
            "% Aliq",
            "% Aliq Efetiva",
            "Linhas",
            "Valor Operacao",
            "Base ICMS",
            "% Base/Oper",
            "Valor ICMS",
        ]
        product_export_rows = [
            [
                row.get("period", ""),
                row.get("code", ""),
                row.get("description", ""),
                row.get("ncm", ""),
                Decimal(row.get("icms_rate", Decimal("0"))),
                compute_display_icms_rate(
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    Decimal(row.get("sale_value", Decimal("0"))),
                    Decimal(row.get("base_icms", Decimal("0"))),
                    Decimal(row.get("icms_value", Decimal("0"))),
                ),
                row.get("row_count", 0),
                Decimal(row.get("sale_value", Decimal("0"))),
                Decimal(row.get("base_icms", Decimal("0"))),
                (
                    (
                        Decimal(row.get("base_icms", Decimal("0"))) * Decimal("100")
                        / Decimal(row.get("sale_value", Decimal("0")))
                    ).quantize(Decimal("0.01"))
                    if Decimal(row.get("sale_value", Decimal("0"))) > 0
                    else Decimal("0.00")
                ),
                Decimal(row.get("icms_value", Decimal("0"))),
            ]
            for row in product_rows
        ]
        if not product_export_rows:
            product_export_rows = [["", "", "Nenhum produto/NCM identificado para este agrupamento.", "", "", "", "", "", "", "", ""]]

        resumo_headers = ["Periodo", "Motivo", "Registro", "CST", "CFOP", "Aliq", "Aliq Efetiva", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        resumo_rows = [[
            detail_row.get("period", ""),
            detail_row.get("reason", ""),
            detail_row.get("source_register", ""),
            detail_row.get("cst_icms", ""),
            detail_row.get("cfop", ""),
            Decimal(detail_row.get("icms_rate", Decimal("0"))),
            compute_display_icms_rate(
                Decimal(detail_row.get("icms_rate", Decimal("0"))),
                Decimal(detail_row.get("total_operation_value", Decimal("0"))),
                Decimal(detail_row.get("base_icms", Decimal("0"))),
                Decimal(detail_row.get("icms_value", Decimal("0"))),
            ),
            detail_row.get("row_count", 0),
            Decimal(detail_row.get("total_operation_value", Decimal("0"))),
            Decimal(detail_row.get("base_icms", Decimal("0"))),
            Decimal(detail_row.get("base_gap", Decimal("0"))),
            Decimal(detail_row.get("base_ratio", Decimal("0"))),
            Decimal(detail_row.get("icms_value", Decimal("0"))),
        ]]
        try:
            write_simple_excel_workbook(
                output_path,
                [
                    ("Resumo Agrupamento", resumo_headers, resumo_rows),
                    ("Linhas Agrupadas", detail_headers, detail_export_rows),
                    ("Produtos Relacionados", product_headers, product_export_rows),
                ],
            )
            self.log_message(f"Exportacao do popup concluida em: {output_path}")
            messagebox.showinfo("Exportar popup", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar popup: {exc}")
            messagebox.showerror("Exportar popup", str(exc))

    def open_operation_summary_popup(self, operation_type: str, source_view: str = "default") -> None:
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir popup de {operation_type.lower()}: {exc}")
            messagebox.showerror("Popup fiscal", str(exc))
            return

        filtered_details = list(context["filtered_details"])
        grouped_launch_details = build_operation_launch_details_map(
            filtered_details,
            operation_type,
            str(context["selected_rule_profile"]),
            list(context["runtime_rules"]),
        )
        if bool(context["can_use_c190_directly"]):
            filtered_c190_rows = filter_c190_rows(
                list(context["summary_register_rows"]),
                set(context["cst_filter_values"]),
                set(context["cfop_filter_values"]),
            )
            rows, totals = build_operation_summary_rows_from_c190(filtered_c190_rows, operation_type)
        else:
            rows, totals = build_operation_summary_rows(
                filtered_details,
                operation_type,
                str(context["selected_rule_profile"]),
                list(context["runtime_rules"]),
            )
        if not rows:
            messagebox.showwarning("Popup fiscal", f"Nao ha dados de {operation_type.lower()} para os filtros atuais.")
            return

        enriched_rows: list[dict[str, object]] = []
        for row in rows:
            row_key = (
                normalize_cst_icms_for_sped(str(row.get("cst_icms", ""))),
                str(row.get("cfop", "")).strip(),
                Decimal(row.get("icms_rate", Decimal("0"))).quantize(Decimal("0.01")),
            )
            launch_details = list(
                grouped_launch_details.get(
                    row_key,
                    [],
                )
            )
            if not launch_details and bool(context["can_use_c190_directly"]):
                launch_details = build_synthetic_launch_details_from_c190(
                    list(context["summary_register_rows"]),
                    row_key[0],
                    row_key[1],
                    row_key[2],
                    operation_type,
                )
            document_keys = {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in launch_details
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
            base_ipi = sum((Decimal(detail.get("base_ipi", Decimal("0"))) for detail in launch_details), Decimal("0")).quantize(Decimal("0.01"))
            detail_ipi_value = sum((Decimal(detail.get("ipi_value", Decimal("0"))) for detail in launch_details), Decimal("0")).quantize(Decimal("0.01"))
            row_ipi_value = Decimal(row.get("ipi_value", Decimal("0"))).quantize(Decimal("0.01"))
            if row_ipi_value == Decimal("0") and detail_ipi_value > Decimal("0"):
                row_ipi_value = detail_ipi_value
            ipi_rate = (row_ipi_value * Decimal("100") / base_ipi).quantize(Decimal("0.01")) if base_ipi > 0 else Decimal("0.00")
            enriched_rows.append(
                {
                    **row,
                    "base_ipi": base_ipi,
                    "ipi_rate": ipi_rate,
                    "ipi_value": row_ipi_value,
                    "launch_details": launch_details,
                    "document_count": len(document_keys),
                    "launch_count": len(launch_details),
                }
            )
        rows = enriched_rows

        dialog = Toplevel(self.root)
        dialog.title(f"Resumo {operation_type}s")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1220, 560, 980, 460, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(
            container,
            text=f"Resumo - {operation_type}s por CST, CFOP e Aliquota",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=self.format_popup_filter_caption(context),
            wraplength=1160,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=1)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=2)
        cst_var = StringVar()
        cfop_var = StringVar()
        rate_var = StringVar()
        search_var = StringVar()

        for column_index, (label_text, variable) in enumerate((
            ("Filtro CST", cst_var),
            ("Filtro CFOP", cfop_var),
            ("Filtro Aliquota", rate_var),
            ("Busca", search_var),
        )):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 3 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=3, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = (
            "cst_icms",
            "cfop",
            "icms_rate",
            "effective_icms_rate",
            "base_ipi",
            "ipi_rate",
            "ipi_value",
            "icms_value",
            "base_icms_st",
            "icms_st_value",
            "total_operation_value",
            "base_icms",
            "operation_base_gap",
            "difference_origin",
            "reduction_value",
            "documents",
            "launch_count",
        )
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=14, selectmode="extended")
        headings = {
            "cst_icms": "CST",
            "cfop": "CFOP",
            "icms_rate": "Aliq ICMS",
            "effective_icms_rate": "Aliq Efetiva",
            "base_ipi": "Base IPI",
            "ipi_rate": "Aliq IPI",
            "ipi_value": "Valor IPI",
            "icms_value": "Valor ICMS",
            "base_icms_st": "Base ICMS ST",
            "icms_st_value": "Valor ICMS ST",
            "total_operation_value": "Total Operacao",
            "base_icms": "Base ICMS",
            "operation_base_gap": "Dif. Oper/Base",
            "difference_origin": "Origem Dif.",
            "reduction_value": "Reducao BC",
            "documents": "Docs",
            "launch_count": "Lanc.",
        }
        widths = {
            "cst_icms": 70,
            "cfop": 80,
            "icms_rate": 90,
            "effective_icms_rate": 95,
            "base_ipi": 100,
            "ipi_rate": 90,
            "ipi_value": 100,
            "icms_value": 100,
            "base_icms_st": 120,
            "icms_st_value": 110,
            "total_operation_value": 120,
            "base_icms": 110,
            "operation_base_gap": 110,
            "difference_origin": 110,
            "reduction_value": 100,
            "documents": 70,
            "launch_count": 70,
        }
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.tag_configure("dynamic_rule", background="#eadcff")
        tree.bind("<Double-1>", lambda event: self.handle_operation_summary_tree_double_click(event, tree, rows, operation_type))
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_operation_summary_tree_menu(event, tree, rows, operation_type))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        total_rows_var = StringVar(value="Linhas: 0")
        total_op_var = StringVar(value="Total Operacao: 0,00")
        total_base_var = StringVar(value="Base ICMS: 0,00")
        total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        total_icms_var = StringVar(value="Valor ICMS: 0,00")
        total_base_st_var = StringVar(value="Base ST: 0,00")
        total_st_var = StringVar(value="Valor ICMS ST: 0,00")
        total_ipi_var = StringVar(value="Valor IPI: 0,00")
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_op_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_icms_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_st_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_st_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_ipi_var, style="Strong.TLabel").pack(side=LEFT)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=5, column=0, sticky="w", pady=(10, 0))

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)

            cst_filter = normalize_text(cst_var.get())
            cfop_filter = normalize_text(cfop_var.get())
            rate_filter = normalize_text(rate_var.get())
            search_filter = normalize_text(search_var.get())
            total_operation_value = Decimal("0")
            total_base_icms = Decimal("0")
            total_icms_value = Decimal("0")
            total_base_icms_st = Decimal("0")
            total_icms_st_value = Decimal("0")
            total_ipi_value = Decimal("0")
            filtered_count = 0
            runtime_rules = self.get_current_runtime_rules_for_highlight()

            for row in rows:
                rate_text = format_decimal_sped(Decimal(row["icms_rate"]))
                search_text = normalize_text(
                    f"{row['cst_icms']} {row['cfop']} {rate_text} "
                    f"{format_decimal_sped(Decimal(row['total_operation_value']))}"
                )
                if cst_filter and cst_filter not in normalize_text(str(row["cst_icms"])):
                    continue
                if cfop_filter and cfop_filter not in normalize_text(str(row["cfop"])):
                    continue
                if rate_filter and rate_filter not in normalize_text(rate_text):
                    continue
                if search_filter and search_filter not in search_text:
                    continue

                document_keys = row.get("document_keys", set())
                effective_rate = compute_display_icms_rate(
                    Decimal(row["icms_rate"]),
                    Decimal(row["total_operation_value"]),
                    Decimal(row["base_icms"]),
                    Decimal(row["icms_value"]),
                )
                current_row_index = filtered_count
                operation_value = Decimal(row["total_operation_value"])
                base_icms = Decimal(row["base_icms"])
                ipi_value = Decimal(row["ipi_value"])
                icms_st_value = Decimal(row["icms_st_value"])
                operation_base_gap = (operation_value - base_icms).quantize(Decimal("0.01"))
                difference_origin = describe_operation_base_difference(
                    operation_value,
                    base_icms,
                    ipi_value,
                    icms_st_value,
                    Decimal(row.get("discount_value", Decimal("0"))),
                )
                tree.insert(
                    "",
                    END,
                    iid=str(current_row_index),
                    values=(
                        row["cst_icms"],
                        row["cfop"],
                        rate_text,
                        format_decimal_sped(effective_rate),
                        format_decimal_sped(Decimal(row.get("base_ipi", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("ipi_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row["ipi_value"])),
                        format_decimal_sped(Decimal(row["icms_value"])),
                        format_decimal_sped(Decimal(row["base_icms_st"])),
                        format_decimal_sped(Decimal(row["icms_st_value"])),
                        format_decimal_sped(operation_value),
                        format_decimal_sped(base_icms),
                        format_decimal_sped(operation_base_gap),
                        difference_origin,
                        format_decimal_sped(Decimal(row["reduction_value"])),
                        len(document_keys),
                        row["launch_count"],
                    ),
                    tags=(
                        ("dynamic_rule",)
                        if self.consultation_row_matches_runtime_rules(row, runtime_rules)
                        else ()
                    ),
                )
                total_operation_value += Decimal(row["total_operation_value"])
                total_base_icms += Decimal(row["base_icms"])
                total_icms_value += Decimal(row["icms_value"])
                total_base_icms_st += Decimal(row["base_icms_st"])
                total_icms_st_value += Decimal(row["icms_st_value"])
                total_ipi_value += Decimal(row.get("ipi_value", Decimal("0")))
                filtered_count += 1

            total_rows_var.set(f"Linhas: {filtered_count}")
            total_op_var.set(f"Total Operacao: {format_decimal_sped(total_operation_value)}")
            total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base_icms)}")
            total_base_ratio_var.set(
                f"% Base/Oper: {format_decimal_sped((total_base_icms * Decimal('100') / total_operation_value).quantize(Decimal('0.01')) if total_operation_value > 0 else Decimal('0.00'))}"
            )
            total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms_value)}")
            total_base_st_var.set(f"Base ST: {format_decimal_sped(total_base_icms_st)}")
            total_st_var.set(f"Valor ICMS ST: {format_decimal_sped(total_icms_st_value)}")
            total_ipi_var.set(f"Valor IPI: {format_decimal_sped(total_ipi_value)}")
            tree._filtered_operation_rows = [
                row
                for row in rows
                if (
                    (not cst_filter or cst_filter in normalize_text(str(row["cst_icms"])))
                    and (not cfop_filter or cfop_filter in normalize_text(str(row["cfop"])))
                    and (not rate_filter or rate_filter in normalize_text(format_decimal_sped(Decimal(row["icms_rate"]))))
                    and (
                        not search_filter
                        or search_filter in normalize_text(
                            f"{row['cst_icms']} {row['cfop']} {format_decimal_sped(Decimal(row['icms_rate']))} "
                            f"{format_decimal_sped(Decimal(row['total_operation_value']))}"
                        )
                    )
                )
            ]
            tree._filtered_popup_rows = [
                [
                    row["cst_icms"],
                    row["cfop"],
                    Decimal(row["icms_rate"]),
                    compute_display_icms_rate(
                        Decimal(row["icms_rate"]),
                        Decimal(row["total_operation_value"]),
                        Decimal(row["base_icms"]),
                        Decimal(row["icms_value"]),
                    ),
                    Decimal(row.get("base_ipi", Decimal("0"))),
                    Decimal(row.get("ipi_rate", Decimal("0"))),
                    Decimal(row["ipi_value"]),
                    Decimal(row["icms_value"]),
                    Decimal(row["base_icms_st"]),
                    Decimal(row["icms_st_value"]),
                    Decimal(row["total_operation_value"]),
                    Decimal(row["base_icms"]),
                    (Decimal(row["total_operation_value"]) - Decimal(row["base_icms"])).quantize(Decimal("0.01")),
                    describe_operation_base_difference(
                        Decimal(row["total_operation_value"]),
                        Decimal(row["base_icms"]),
                        Decimal(row["ipi_value"]),
                        Decimal(row["icms_st_value"]),
                        Decimal(row.get("discount_value", Decimal("0"))),
                    ),
                    Decimal(row["reduction_value"]),
                    len(row.get("document_keys", set())),
                    row["launch_count"],
                ]
                for row in rows
                if (
                    (not cst_filter or cst_filter in normalize_text(str(row["cst_icms"])))
                    and (not cfop_filter or cfop_filter in normalize_text(str(row["cfop"])))
                    and (not rate_filter or rate_filter in normalize_text(format_decimal_sped(Decimal(row["icms_rate"]))))
                    and (
                        not search_filter
                        or search_filter in normalize_text(
                            f"{row['cst_icms']} {row['cfop']} {format_decimal_sped(Decimal(row['icms_rate']))} "
                            f"{format_decimal_sped(Decimal(row['total_operation_value']))}"
                        )
                    )
                )
            ]

        for variable in (cst_var, cfop_var, rate_var, search_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"popup_{operation_type.lower()}",
                f"Resumo {operation_type}s",
                ["CST", "CFOP", "Aliq ICMS", "Aliq Efetiva", "Base IPI", "Aliq IPI", "Valor IPI", "Valor ICMS", "Base ICMS ST", "Valor ICMS ST", "Total Operacao", "Base ICMS", "Dif. Oper/Base", "Origem Dif.", "Reducao BC", "Docs", "Lanc."],
                list(getattr(tree, "_filtered_popup_rows", [])),
                "xlsx",
            ),
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"popup_{operation_type.lower()}",
                f"Resumo {operation_type}s",
                ["CST", "CFOP", "Aliq ICMS", "Aliq Efetiva", "Base IPI", "Aliq IPI", "Valor IPI", "Valor ICMS", "Base ICMS ST", "Valor ICMS ST", "Total Operacao", "Base ICMS", "Dif. Oper/Base", "Origem Dif.", "Reducao BC", "Docs", "Lanc."],
                list(getattr(tree, "_filtered_popup_rows", [])),
                "csv",
            ),
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.log_message(f"Popup de {operation_type.lower()} aberto com {len(rows)} agrupamento(s).")

    def open_apuracao_popup(self, source_view: str = "default") -> None:
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir popup de apuracao: {exc}")
            messagebox.showerror("Popup fiscal", str(exc))
            return

        filtered_details = list(context["filtered_details"])
        if bool(context["can_use_c190_directly"]):
            filtered_c190_rows = filter_c190_rows(
                list(context["summary_register_rows"]),
                set(context["cst_filter_values"]),
                set(context["cfop_filter_values"]),
            )
            entry_rows, entry_totals = build_operation_summary_rows_from_c190(filtered_c190_rows, "Entrada")
            exit_rows, exit_totals = build_operation_summary_rows_from_c190(filtered_c190_rows, "Saida")
        else:
            entry_rows, entry_totals = build_operation_summary_rows(
                filtered_details,
                "Entrada",
                str(context["selected_rule_profile"]),
                list(context["runtime_rules"]),
            )
            exit_rows, exit_totals = build_operation_summary_rows(
                filtered_details,
                "Saida",
                str(context["selected_rule_profile"]),
                list(context["runtime_rules"]),
            )
        if not entry_rows and not exit_rows:
            messagebox.showwarning("Popup fiscal", "Nao ha dados de apuracao para os filtros atuais.")
            return

        sped_path = context.get("sped_path")
        source_sped_paths = [
            path
            for path in context.get("source_sped_paths", [])
            if isinstance(path, Path) and path.exists() and path.is_file()
        ]
        if isinstance(sped_path, Path) and sped_path.exists() and sped_path.is_file():
            original_summary = read_sped_e110_summary(sped_path)
        elif source_sped_paths:
            original_summary = read_combined_e110_summary(source_sped_paths)
        else:
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
        preserve_original_adjustments = not bool(context["filters_active"])
        apuracao_rows = build_filtered_apuracao_rows(
            entry_totals,
            exit_totals,
            original_summary,
            preserve_original_adjustments,
        )
        validation_ok = (
            not bool(context["filters_active"])
            and apuracao_rows[0][1].quantize(Decimal("0.01")) == original_summary["debitos"].quantize(Decimal("0.01"))
            and apuracao_rows[4][1].quantize(Decimal("0.01")) == original_summary["creditos"].quantize(Decimal("0.01"))
            and apuracao_rows[9][1].quantize(Decimal("0.01")) == original_summary["saldo_devedor"].quantize(Decimal("0.01"))
            and apuracao_rows[11][1].quantize(Decimal("0.01")) == original_summary["icms_recolher"].quantize(Decimal("0.01"))
        )

        dialog = Toplevel(self.root)
        dialog.title("Apuracao do ICMS")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1020, 520, 860, 420, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(container, text="Apuracao do ICMS - Operacoes Proprias", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        caption = self.format_popup_filter_caption(context)
        if validation_ok:
            caption = f"{caption} | Validacao SPED: OK"
        elif not bool(context["filters_active"]) and source_sped_paths:
            caption = f"{caption} | Validacao SPED: divergencia encontrada"
        else:
            caption = f"{caption} | Valores recalculados com base no filtro atual"
        ttk.Label(container, text=caption, wraplength=960).grid(row=1, column=0, sticky="w", pady=(6, 10))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        tree = ttk.Treeview(tree_box, columns=("descricao", "valor"), show="headings", height=14, selectmode="extended")
        tree.column("descricao", width=760, anchor="center")
        tree.column("valor", width=160, anchor="center")
        self.enable_treeview_sorting(
            tree,
            ("descricao", "valor"),
            {"descricao": "Descricao", "valor": "Valor R$"},
        )
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll_y.set)

        for description, value in apuracao_rows:
            tree.insert("", END, values=(description, format_decimal_sped(value)))

        export_rows = [[description, value] for description, value in apuracao_rows]
        export_actions = ttk.Frame(container)
        export_actions.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                "popup_apuracao",
                "Apuracao",
                ["Descricao", "Valor R$"],
                export_rows,
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                "popup_apuracao",
                "Apuracao",
                ["Descricao", "Valor R$"],
                export_rows,
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            footer,
            text=f"Entradas (credito): {format_decimal_sped(entry_totals['icms_value'])}",
            style="Strong.TLabel",
        ).pack(side=LEFT, padx=(0, 18))
        ttk.Label(
            footer,
            text=f"Saidas (debito): {format_decimal_sped(exit_totals['icms_value'])}",
            style="Strong.TLabel",
        ).pack(side=LEFT, padx=(0, 18))
        ttk.Label(
            footer,
            text=f"ICMS a Recolher: {format_decimal_sped(apuracao_rows[11][1])}",
            style="Strong.TLabel",
        ).pack(side=LEFT)
        self.log_message("Popup de apuracao aberto.")

    def open_reduction_items_popup(self, source_view: str = "default") -> None:
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir popup de reducao de base: {exc}")
            messagebox.showerror("Popup fiscal", str(exc))
            return

        filtered_details = list(context["filtered_details"])
        rows = build_reduction_launch_rows(
            filtered_details,
            str(context["selected_rule_profile"]),
            list(context["runtime_rules"]),
        )
        if not rows and bool(context["can_use_c190_directly"]):
            filtered_c190_rows = filter_c190_rows(
                list(context["summary_register_rows"]),
                set(context["cst_filter_values"]),
                set(context["cfop_filter_values"]),
            )
            rows = build_reduction_rows_from_c190(filtered_c190_rows)
        if not rows:
            messagebox.showwarning("Popup fiscal", "Nao ha itens com reducao de base para os filtros atuais.")
            return

        dialog = Toplevel(self.root)
        dialog.title("Itens com Reducao de Base")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1320, 720, 1080, 560, margin_y=190)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(
            container,
            text="Itens com Reducao de Base de Calculo",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text=self.format_popup_filter_caption(context),
            wraplength=1260,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=2)
        doc_var = StringVar()
        cfop_var = StringVar()
        cst_var = StringVar()
        search_var = StringVar()

        for column_index, (label_text, variable) in enumerate((
            ("Documento", doc_var),
            ("Filtro CFOP", cfop_var),
            ("Filtro CST", cst_var),
            ("Busca", search_var),
        )):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 3 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=3, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = (
            "operation_type",
            "document_date",
            "document_number",
            "participant_name",
            "document_key",
            "item_number",
            "code",
            "description",
            "cfop",
            "cst_icms",
            "icms_rate",
            "effective_icms_rate",
            "sale_value",
            "base_icms",
            "base_ratio",
            "reduction_value",
            "icms_value",
        )
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=12, selectmode="extended")
        headings = {
            "operation_type": "Tipo",
            "document_date": "Data",
            "document_number": "Documento",
            "participant_name": "Fornecedor/Participante",
            "document_key": "Chave",
            "item_number": "Item",
            "code": "Codigo",
            "description": "Descricao",
            "cfop": "CFOP",
            "cst_icms": "CST",
            "icms_rate": "Aliq ICMS",
            "effective_icms_rate": "Aliq Efetiva",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "base_ratio": "% Base/Oper",
            "reduction_value": "Reducao BC",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "operation_type": 70,
            "document_date": 90,
            "document_number": 90,
            "participant_name": 200,
            "document_key": 250,
            "item_number": 60,
            "code": 100,
            "description": 260,
            "cfop": 80,
            "cst_icms": 70,
            "icms_rate": 85,
            "effective_icms_rate": 95,
            "sale_value": 110,
            "base_icms": 110,
            "base_ratio": 100,
            "reduction_value": 110,
            "icms_value": 100,
        }
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.tag_configure("reduced", background="#e8f2ff")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        tree.bind("<Double-1>", lambda event: self.handle_reduction_popup_double_click(event, tree, rows))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=4, column=0, sticky="w", pady=(10, 0))

        footer = ttk.Frame(container)
        footer.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        total_rows_var = StringVar(value="Linhas: 0")
        total_sale_var = StringVar(value="Valor Operacao: 0,00")
        total_base_var = StringVar(value="Base ICMS: 0,00")
        total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        total_reduction_var = StringVar(value="Reducao BC: 0,00")
        total_icms_var = StringVar(value="Valor ICMS: 0,00")
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_reduction_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)

            doc_filter = normalize_text(doc_var.get())
            cfop_filter = normalize_text(cfop_var.get())
            cst_filter = normalize_text(cst_var.get())
            search_filter = normalize_text(search_var.get())
            total_sale = Decimal("0")
            total_base = Decimal("0")
            total_reduction = Decimal("0")
            total_icms = Decimal("0")
            filtered_count = 0
            filtered_rows: list[dict[str, object]] = []

            for row in rows:
                search_text = normalize_text(
                    f"{row.get('document_number', '')} {row.get('document_key', '')} {row.get('item_number', '')} "
                    f"{row.get('code', '')} {row.get('description', '')} {row.get('cfop', '')} {row.get('cst_icms', '')}"
                )
                if doc_filter and doc_filter not in normalize_text(str(row.get("document_number", ""))):
                    continue
                if cfop_filter and cfop_filter not in normalize_text(str(row.get("cfop", ""))):
                    continue
                if cst_filter and cst_filter not in normalize_text(str(row.get("cst_icms", ""))):
                    continue
                if search_filter and search_filter not in search_text:
                    continue

                current_row_index = filtered_count
                tree.insert(
                    "",
                    END,
                    iid=str(current_row_index),
                    values=(
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
                        format_decimal_sped(Decimal(row.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(
                            compute_display_icms_rate(
                                Decimal(row.get("icms_rate", Decimal("0"))),
                                Decimal(row.get("sale_value", Decimal("0"))),
                                Decimal(row.get("base_icms", Decimal("0"))),
                                Decimal(row.get("icms_value", Decimal("0"))),
                            )
                        ),
                        format_decimal_sped(Decimal(row.get("sale_value", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("base_icms", Decimal("0")))),
                        format_decimal_sped(
                            (
                                (
                                    Decimal(row.get("base_icms", Decimal("0"))) * Decimal("100")
                                    / Decimal(row.get("sale_value", Decimal("0")))
                                ).quantize(Decimal("0.01"))
                                if Decimal(row.get("sale_value", Decimal("0"))) > 0
                                else Decimal("0.00")
                            )
                        ),
                        format_decimal_sped(Decimal(row.get("reduction_value", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("icms_value", Decimal("0")))),
                    ),
                    tags=("reduced",),
                )
                filtered_rows.append(row)
                total_sale += Decimal(row.get("sale_value", Decimal("0")))
                total_base += Decimal(row.get("base_icms", Decimal("0")))
                total_reduction += Decimal(row.get("reduction_value", Decimal("0")))
                total_icms += Decimal(row.get("icms_value", Decimal("0")))
                filtered_count += 1

            tree._filtered_reduction_rows = filtered_rows
            tree._filtered_popup_rows = [
                [
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
                    Decimal(row.get("icms_rate", Decimal("0"))),
                    compute_display_icms_rate(
                        Decimal(row.get("icms_rate", Decimal("0"))),
                        Decimal(row.get("sale_value", Decimal("0"))),
                        Decimal(row.get("base_icms", Decimal("0"))),
                        Decimal(row.get("icms_value", Decimal("0"))),
                    ),
                    Decimal(row.get("sale_value", Decimal("0"))),
                    Decimal(row.get("base_icms", Decimal("0"))),
                    (
                        (
                            Decimal(row.get("base_icms", Decimal("0"))) * Decimal("100")
                            / Decimal(row.get("sale_value", Decimal("0")))
                        ).quantize(Decimal("0.01"))
                        if Decimal(row.get("sale_value", Decimal("0"))) > 0
                        else Decimal("0.00")
                    ),
                    Decimal(row.get("reduction_value", Decimal("0"))),
                    Decimal(row.get("icms_value", Decimal("0"))),
                ]
                for row in filtered_rows
            ]
            total_rows_var.set(f"Linhas: {filtered_count}")
            total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
            total_base_ratio_var.set(
                f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}"
            )
            total_reduction_var.set(f"Reducao BC: {format_decimal_sped(total_reduction)}")
            total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

        for variable in (doc_var, cfop_var, cst_var, search_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()

        headers = ["Tipo", "Data", "Documento", "Fornecedor/Participante", "Chave", "Item", "Codigo", "Descricao", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Valor Operacao", "Base ICMS", "% Base/Oper", "Reducao BC", "Valor ICMS"]
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                "popup_reducao_base",
                "Reducao Base",
                headers,
                list(getattr(tree, "_filtered_popup_rows", [])),
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                "popup_reducao_base",
                "Reducao Base",
                headers,
                list(getattr(tree, "_filtered_popup_rows", [])),
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))
        self.log_message(f"Popup de reducao de base aberto com {len(rows)} linha(s).")

    def handle_reduction_popup_double_click(
        self,
        event: object,
        tree: ttk.Treeview,
        rows: list[dict[str, object]],
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        filtered_rows = getattr(tree, "_filtered_reduction_rows", rows)
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(filtered_rows):
            return
        row = filtered_rows[row_index]
        self.open_launch_origin_details_popup(
            [dict(row)],
            str(row.get("code", "")).strip() or "Item",
            str(row.get("description", "")).strip() or "Origem do lancamento",
            str(row.get("document_date", "")).strip() or str(row.get("operation_type", "")).strip(),
        )

    def handle_operation_summary_tree_double_click(
        self,
        event: object,
        tree: ttk.Treeview,
        rows: list[dict[str, object]],
        operation_type: str,
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        filtered_rows = getattr(tree, "_filtered_operation_rows", rows)
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(filtered_rows):
            return
        row = filtered_rows[row_index]
        launch_details = list(row.get("launch_details", []))
        if not launch_details:
            messagebox.showwarning("Detalhamento", f"Nao ha lancamentos detalhados para este agrupamento de {operation_type.lower()}.")
            return
        detail_row = {
            "period": f"{operation_type}s",
            "code": f"{row.get('cst_icms', '')}/{row.get('cfop', '')}",
            "description": f"Aliq {format_decimal_sped(Decimal(row.get('icms_rate', Decimal('0'))))}",
            "document_count": int(row.get("document_count", 0)),
            "launch_count": int(row.get("launch_count", 0)),
            "launch_details": launch_details,
        }
        self.open_consultation_launch_details(detail_row)

    def show_operation_summary_tree_menu(
        self,
        event: object,
        tree: ttk.Treeview,
        rows: list[dict[str, object]],
        operation_type: str,
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            self.show_treeview_copy_menu(event, tree)
            return
        tree.selection_set(row_id)
        filtered_rows = getattr(tree, "_filtered_operation_rows", rows)
        try:
            row_index = int(row_id)
        except ValueError:
            self.show_treeview_copy_menu(event, tree)
            return
        if row_index < 0 or row_index >= len(filtered_rows):
            self.show_treeview_copy_menu(event, tree)
            return
        row = filtered_rows[row_index]
        menu = Menu(tree, tearoff=0)
        has_matching_rule = self.add_matching_runtime_rule_menu_items(menu, row)
        if not has_matching_rule:
            menu.add_command(
                label="Criar Regra Dinamica",
                command=lambda current_row=row, current_row_id=row_id: self.open_runtime_rule_builder_for_operation_summary(
                    current_row,
                    operation_type,
                    tree,
                    current_row_id,
                ),
            )
        menu.add_command(
            label="Mostrar Produtos do Consolidado",
            command=lambda current_row=row: self.open_operation_summary_products_popup(current_row, operation_type),
        )
        menu.add_separator()
        menu.add_command(label="Copiar Selecao", command=lambda: self.copy_treeview_selection(tree))
        menu.tk_popup(getattr(event, "x_root", 0), getattr(event, "y_root", 0))

    def open_operation_summary_products_popup(self, summary_row: dict[str, object], operation_type: str) -> None:
        launch_details = list(summary_row.get("launch_details", []))
        if not launch_details:
            messagebox.showwarning("Produtos do consolidado", "Nao ha produtos detalhados para este consolidado.")
            return

        grouped_products: dict[tuple[str, str, str], dict[str, object]] = {}
        for detail in launch_details:
            code = str(detail.get("code", "")).strip()
            description = str(detail.get("description", "")).strip()
            ncm = str(detail.get("ncm", "")).strip()
            product_key = (code, description, ncm)
            bucket = grouped_products.setdefault(
                product_key,
                {
                    "code": code,
                    "description": description,
                    "ncm": ncm,
                    "row_count": 0,
                    "document_keys": set(),
                    "quantity": Decimal("0"),
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                    "launch_details": [],
                },
            )
            bucket["row_count"] += 1
            document_identity = normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            if document_identity:
                bucket["document_keys"].add(document_identity)
            bucket["quantity"] += Decimal(detail.get("quantity", Decimal("0")))
            bucket["sale_value"] += Decimal(detail.get("sale_value", Decimal("0")))
            bucket["base_icms"] += Decimal(detail.get("base_icms", Decimal("0")))
            bucket["icms_value"] += Decimal(detail.get("icms_value", Decimal("0")))
            bucket["launch_details"].append(dict(detail))

        product_rows = sorted(
            grouped_products.values(),
            key=lambda row: (Decimal(row.get("sale_value", Decimal("0"))), str(row.get("code", "")), str(row.get("description", ""))),
            reverse=True,
        )

        dialog = Toplevel(self.root)
        dialog.title(f"Produtos do Consolidado - {operation_type}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1180, 560, 980, 460, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        ttk.Label(
            container,
            text=(
                f"Consolidado {operation_type}: CST {summary_row.get('cst_icms', '')} | "
                f"CFOP {summary_row.get('cfop', '')} | "
                f"Aliq {format_decimal_sped(Decimal(summary_row.get('icms_rate', Decimal('0'))))}"
            ),
            style="Title.TLabel",
            wraplength=1120,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text="Duplo clique em um produto abre os lancamentos/notas desse produto dentro do consolidado.",
            wraplength=1120,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        export_headers = ["Codigo", "Produto", "NCM", "Docs", "Lanc.", "Quantidade", "Valor Operacao", "Base ICMS", "Valor ICMS"]
        export_rows = [
            [
                product_row.get("code", ""),
                product_row.get("description", ""),
                product_row.get("ncm", ""),
                len(product_row.get("document_keys", set())),
                product_row.get("row_count", 0),
                Decimal(product_row.get("quantity", Decimal("0"))),
                Decimal(product_row.get("sale_value", Decimal("0"))),
                Decimal(product_row.get("base_icms", Decimal("0"))),
                Decimal(product_row.get("icms_value", Decimal("0"))),
            ]
            for product_row in product_rows
        ]

        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = ("code", "description", "ncm", "documents", "launches", "quantity", "sale_value", "base_icms", "icms_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=16, selectmode="extended")
        headings = {
            "code": "Codigo",
            "description": "Produto",
            "ncm": "NCM",
            "documents": "Docs",
            "launches": "Lanc.",
            "quantity": "Quantidade",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "code": 120,
            "description": 360,
            "ncm": 110,
            "documents": 70,
            "launches": 70,
            "quantity": 100,
            "sale_value": 130,
            "base_icms": 130,
            "icms_value": 130,
        }
        for column_id in columns:
            tree.heading(column_id, text=headings[column_id])
            tree.column(column_id, width=widths[column_id], anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        tree.tag_configure("dynamic_rule", background="#eadcff")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind(
            "<Button-3>",
            lambda event: self.show_operation_summary_products_menu(event, tree, product_rows, summary_row, operation_type),
        )
        tree.bind(
            "<Double-1>",
            lambda event: self.handle_operation_summary_products_double_click(event, tree, product_rows, summary_row, operation_type),
        )

        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        total_quantity = Decimal("0")
        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for product_row in product_rows:
            total_quantity += Decimal(product_row.get("quantity", Decimal("0")))
            total_sale += Decimal(product_row.get("sale_value", Decimal("0")))
            total_base += Decimal(product_row.get("base_icms", Decimal("0")))
            total_icms += Decimal(product_row.get("icms_value", Decimal("0")))
            product_tags: tuple[str, ...] = (
                ("dynamic_rule",)
                if self.consultation_row_matches_runtime_rules(product_row, self.get_current_runtime_rules_for_highlight())
                else ()
            )
            tree.insert(
                "",
                END,
                iid=str(len(tree.get_children())),
                values=(
                    product_row.get("code", ""),
                    product_row.get("description", ""),
                    product_row.get("ncm", ""),
                    len(product_row.get("document_keys", set())),
                    product_row.get("row_count", 0),
                    format_decimal_sped(Decimal(product_row.get("quantity", Decimal("0")))),
                    format_decimal_sped(Decimal(product_row.get("sale_value", Decimal("0")))),
                    format_decimal_sped(Decimal(product_row.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(product_row.get("icms_value", Decimal("0")))),
                ),
                tags=product_tags,
            )

        footer = ttk.Frame(container)
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Produtos: {len(product_rows)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Quantidade: {format_decimal_sped(total_quantity)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(total_sale)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(total_icms)}", style="Strong.TLabel").pack(side=LEFT)
        footer_actions = ttk.Frame(footer)
        footer_actions.pack(side=RIGHT)
        ttk.Button(
            footer_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"produtos_consolidado_{operation_type.lower()}",
                f"Produtos Consolidado {operation_type}",
                export_headers,
                export_rows,
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            footer_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"produtos_consolidado_{operation_type.lower()}",
                f"Produtos Consolidado {operation_type}",
                export_headers,
                export_rows,
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))

    def handle_operation_summary_products_double_click(
        self,
        event: object,
        tree: ttk.Treeview,
        product_rows: list[dict[str, object]],
        summary_row: dict[str, object],
        operation_type: str,
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(product_rows):
            return
        product_row = product_rows[row_index]
        detail_row = {
            "period": f"{operation_type}s",
            "code": str(product_row.get("code", "")).strip() or f"{summary_row.get('cst_icms', '')}/{summary_row.get('cfop', '')}",
            "description": str(product_row.get("description", "")).strip() or "Produto consolidado",
            "document_count": len(product_row.get("document_keys", set())),
            "launch_count": int(product_row.get("row_count", 0)),
            "launch_details": list(product_row.get("launch_details", [])),
        }
        self.open_consultation_launch_details(detail_row)

    def show_operation_summary_products_menu(
        self,
        event: object,
        tree: ttk.Treeview,
        product_rows: list[dict[str, object]],
        summary_row: dict[str, object],
        operation_type: str,
    ) -> None:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            self.show_treeview_copy_menu(event, tree)
            return
        tree.selection_set(row_id)
        try:
            row_index = int(row_id)
        except ValueError:
            self.show_treeview_copy_menu(event, tree)
            return
        if row_index < 0 or row_index >= len(product_rows):
            self.show_treeview_copy_menu(event, tree)
            return

        product_row = product_rows[row_index]
        menu = Menu(tree, tearoff=0)
        has_matching_rule = self.add_matching_runtime_rule_menu_items(menu, product_row)
        if not has_matching_rule:
            menu.add_command(
                label="Criar Regra Dinamica",
                command=lambda current_product=product_row: self.open_runtime_rule_builder_for_consolidated_product(
                    current_product,
                    summary_row,
                    operation_type,
                    tree,
                    row_id,
                ),
            )
        menu.add_command(
            label="Mostrar Lancamentos do Produto",
            command=lambda current_product=product_row: self.open_operation_summary_product_details(
                current_product,
                summary_row,
                operation_type,
            ),
        )
        menu.add_separator()
        menu.add_command(label="Copiar Selecao", command=lambda: self.copy_treeview_selection(tree))
        menu.tk_popup(getattr(event, "x_root", 0), getattr(event, "y_root", 0))

    def open_operation_summary_product_details(
        self,
        product_row: dict[str, object],
        summary_row: dict[str, object],
        operation_type: str,
    ) -> None:
        detail_row = {
            "period": f"{operation_type}s",
            "code": str(product_row.get("code", "")).strip() or f"{summary_row.get('cst_icms', '')}/{summary_row.get('cfop', '')}",
            "description": str(product_row.get("description", "")).strip() or "Produto consolidado",
            "document_count": len(product_row.get("document_keys", set())),
            "launch_count": int(product_row.get("row_count", 0)),
            "launch_details": list(product_row.get("launch_details", [])),
        }
        self.open_consultation_launch_details(detail_row)

    def open_runtime_rule_builder_for_consolidated_product(
        self,
        product_row: dict[str, object],
        summary_row: dict[str, object],
        operation_type: str,
        source_tree: ttk.Treeview | None = None,
        source_row_id: str | None = None,
    ) -> None:
        launch_details = list(product_row.get("launch_details", []))
        first_detail = next((detail for detail in launch_details if isinstance(detail, dict)), {})
        cst_value = str(summary_row.get("cst_icms", "") or first_detail.get("cst_icms", "")).strip()
        cfop_value = str(summary_row.get("cfop", "") or first_detail.get("cfop", "")).strip()
        rate_value = summary_row.get("icms_rate", first_detail.get("icms_rate", Decimal("0")))
        try:
            rate_text = format_decimal_sped(Decimal(rate_value))
        except Exception:
            rate_text = str(rate_value or "").strip()

        def mark_source_row(_rule_line: str) -> None:
            if source_tree is not None and source_row_id:
                try:
                    current_tags = tuple(source_tree.item(source_row_id, "tags") or ())
                    if "dynamic_rule" not in current_tags:
                        source_tree.item(source_row_id, tags=(*current_tags, "dynamic_rule"))
                except Exception:
                    pass
            if normalize_operation_type(operation_type) == "Entrada":
                self.refresh_consultation_tree()
            elif normalize_operation_type(operation_type) == "Saida":
                self.refresh_sales_consultation_tree()

        self.open_runtime_rule_builder(
            {
                "operation": normalize_operation_type(operation_type).lower() or str(operation_type).strip().lower(),
                "codes": str(product_row.get("code", "")).strip(),
                "cst": cst_value.split("|", 1)[0].strip(),
                "cfop": cfop_value.split("|", 1)[0].strip(),
                "rate": rate_text,
            },
            on_rule_added=mark_source_row,
        )

    def open_runtime_rule_builder_for_operation_summary(
        self,
        summary_row: dict[str, object],
        operation_type: str,
        source_tree: ttk.Treeview | None = None,
        source_row_id: str | None = None,
    ) -> None:
        cst_value = str(summary_row.get("cst_icms", "")).strip()
        cfop_value = str(summary_row.get("cfop", "")).strip()
        rate_value = summary_row.get("icms_rate", Decimal("0"))
        try:
            rate_text = format_decimal_sped(Decimal(rate_value))
        except Exception:
            rate_text = str(rate_value or "").strip()

        def mark_source_row(_rule_line: str) -> None:
            if source_tree is not None and source_row_id:
                try:
                    current_tags = tuple(source_tree.item(source_row_id, "tags") or ())
                    if "dynamic_rule" not in current_tags:
                        source_tree.item(source_row_id, tags=(*current_tags, "dynamic_rule"))
                except Exception:
                    pass
            if normalize_operation_type(operation_type) == "Entrada":
                self.refresh_consultation_tree()
            elif normalize_operation_type(operation_type) == "Saida":
                self.refresh_sales_consultation_tree()

        self.open_runtime_rule_builder(
            {
                "operation": normalize_operation_type(operation_type).lower() or str(operation_type).strip().lower(),
                "cst": cst_value.split("|", 1)[0].strip(),
                "cfop": cfop_value.split("|", 1)[0].strip(),
                "rate": rate_text,
            },
            on_rule_added=mark_source_row,
        )

    def process_consultation_speds(self, sync_sales: bool = True, show_success: bool = True) -> None:
        sped_paths = parse_selected_paths(self.consult_sped_paths_var.get())
        xml_sources = parse_selected_paths(self.consult_xml_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione de 1 a 12 arquivos SPED para processar.")
            return
        if len(sped_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            return

        progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
            "Processar Consultas",
            "Preparando processamento de entradas...",
        )
        try:
            self.log_message("Processando SPEDs para consulta comparativa...")
            self.write_audit_log(
                "CONSULTA_ENTRADAS_INICIO",
                f"speds={self.format_audit_paths(sped_paths)}; xmls={self.format_audit_paths(xml_sources)}",
            )
            period_labels, comparison_rows = build_entry_period_comparison_rows(
                sped_paths,
                xml_sources,
                progress_callback=lambda current, total, message: self.update_progress_dialog(
                    progress_dialog,
                    progress_message_var,
                    progress_percent_var,
                    current,
                    total,
                    message,
                ),
            )
            self.consult_period_labels = period_labels
            self.consult_period_path_map = {label: path for label, path in zip(period_labels, sped_paths)}
            self.consult_comparison_rows = comparison_rows
            self.rebuild_consultation_periods()
            self.refresh_consultation_tree()
            self.sales_consult_sped_paths_var.set(format_selected_paths(sped_paths))
            self.log_message(
                f"Consulta carregada com sucesso: {len(period_labels)} periodo(s) e {len(comparison_rows)} linha(s) comparativas."
            )
            self.write_audit_log(
                "CONSULTA_ENTRADAS_FIM",
                f"periodos={len(period_labels)}; linhas={len(comparison_rows)}; speds={self.format_audit_paths(sped_paths)}",
            )
            if sync_sales:
                self.log_message("Sincronizando a mesma lista de SPEDs na Consulta Saidas...")
                if xml_sources:
                    self.sales_consult_xml_paths_var.set(format_selected_paths(xml_sources))
                progress_message_var.set("Entradas concluídas. Iniciando processamento de saídas...")
                progress_percent_var.set("0%")
                progress_dialog._progressbar["value"] = 0
                progress_dialog.update_idletasks()
                self.process_sales_consultation_speds(show_success=False, external_progress=(progress_dialog, progress_message_var, progress_percent_var))
            if show_success:
                messagebox.showinfo("Processar Consultas", "Processado com sucesso.")
        except Exception as exc:
            self.log_message(f"Falha ao processar SPEDs da consulta: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))
        finally:
            if progress_dialog.winfo_exists():
                progress_dialog.destroy()

    def process_sales_consultation_speds(
        self,
        show_success: bool = True,
        external_progress: tuple[Toplevel, StringVar, StringVar] | None = None,
    ) -> None:
        sped_paths = parse_selected_paths(self.sales_consult_sped_paths_var.get())
        xml_sources = parse_selected_paths(self.sales_consult_xml_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione de 1 a 12 arquivos SPED para processar.")
            return
        if len(sped_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos SPED.")
            return

        owns_progress = external_progress is None
        if external_progress is None:
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Processar Consultas",
                "Preparando processamento de saídas...",
            )
        else:
            progress_dialog, progress_message_var, progress_percent_var = external_progress
        try:
            self.log_message("Processando SPEDs para consulta comparativa de saidas...")
            self.write_audit_log(
                "CONSULTA_SAIDAS_INICIO",
                f"speds={self.format_audit_paths(sped_paths)}; xmls={self.format_audit_paths(xml_sources)}",
            )
            period_labels, comparison_rows = build_sale_period_comparison_rows(
                sped_paths,
                xml_sources,
                progress_callback=lambda current, total, message: self.update_progress_dialog(
                    progress_dialog,
                    progress_message_var,
                    progress_percent_var,
                    current,
                    total,
                    message,
                ),
            )
            self.sales_consult_period_labels = period_labels
            self.sales_consult_period_path_map = {label: path for label, path in zip(period_labels, sped_paths)}
            self.sales_consult_comparison_rows = comparison_rows
            self.rebuild_sales_consultation_periods()
            self.refresh_sales_consultation_tree()
            self.log_message(
                f"Consulta de saidas carregada com sucesso: {len(period_labels)} periodo(s) e {len(comparison_rows)} linha(s) comparativas."
            )
            self.write_audit_log(
                "CONSULTA_SAIDAS_FIM",
                f"periodos={len(period_labels)}; linhas={len(comparison_rows)}; speds={self.format_audit_paths(sped_paths)}",
            )
            if show_success:
                messagebox.showinfo("Processar Consultas", "Processado com sucesso.")
        except Exception as exc:
            self.log_message(f"Falha ao processar SPEDs da consulta de saidas: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))
        finally:
            if owns_progress and progress_dialog.winfo_exists():
                progress_dialog.destroy()

    def process_all_consultations(self) -> None:
        self.process_consultation_speds(sync_sales=True, show_success=True)

    def process_contrib_consultation_speds(
        self,
        show_success: bool = True,
        sync_sales: bool = False,
        external_progress: tuple[Toplevel, StringVar, StringVar] | None = None,
    ) -> None:
        sped_paths = parse_selected_paths(self.contrib_consult_sped_paths_var.get())
        xml_sources = parse_selected_paths(self.contrib_consult_xml_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione de 1 a 12 arquivos do SPED Contribuicoes para processar.")
            return
        if len(sped_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos do SPED Contribuicoes.")
            return
        if sync_sales:
            self.contrib_sales_consult_sped_paths_var.set(format_selected_paths(sped_paths))
            if xml_sources:
                self.contrib_sales_consult_xml_paths_var.set(format_selected_paths(xml_sources))

        owns_progress = external_progress is None
        if external_progress is None:
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Processar Consultas PIS/COFINS",
                "Preparando processamento de entradas PIS/COFINS...",
            )
        else:
            progress_dialog, progress_message_var, progress_percent_var = external_progress
        try:
            self.write_audit_log(
                "CONSULTA_CONTRIB_ENTRADAS_INICIO",
                f"speds={self.format_audit_paths(sped_paths)}; xmls={self.format_audit_paths(xml_sources)}",
            )
            period_labels, comparison_rows = build_pis_cofins_period_comparison_rows(
                sped_paths,
                "Entrada",
                xml_sources=xml_sources,
                progress_callback=lambda current, total, message: self.update_progress_dialog(
                    progress_dialog, progress_message_var, progress_percent_var, current, total, message
                ),
            )
            self.contrib_consult_period_labels = period_labels
            self.contrib_consult_period_path_map = {label: path for label, path in zip(period_labels, sped_paths)}
            self.contrib_consult_comparison_rows = comparison_rows
            self.rebuild_contrib_consultation_periods()
            self.refresh_contrib_consultation_tree()
            self.log_message(f"Consulta de entradas PIS/COFINS carregada: {len(period_labels)} periodo(s), {len(comparison_rows)} linha(s).")
            self.write_audit_log(
                "CONSULTA_CONTRIB_ENTRADAS_FIM",
                f"periodos={len(period_labels)}; linhas={len(comparison_rows)}; speds={self.format_audit_paths(sped_paths)}",
            )
            if sync_sales:
                self.log_message("Sincronizando o mesmo conjunto de SPEDs/XMLs para saidas de PIS/COFINS...")
                progress_message_var.set("Entradas PIS/COFINS concluidas. Iniciando processamento de saidas...")
                progress_percent_var.set("0%")
                progress_dialog._progressbar["value"] = 0
                progress_dialog.update_idletasks()
                self.process_contrib_sales_consultation_speds(show_success=False, external_progress=(progress_dialog, progress_message_var, progress_percent_var))
            if show_success:
                messagebox.showinfo("Processar Consultas", "Consultas de entradas e saidas PIS/COFINS processadas com sucesso." if sync_sales else "Consulta de entradas PIS/COFINS processada com sucesso.")
        except Exception as exc:
            self.log_message(f"Falha ao processar SPED Contribuicoes de entradas: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))
        finally:
            if owns_progress:
                self.close_progress_dialog(progress_dialog)

    def process_contrib_sales_consultation_speds(
        self,
        show_success: bool = True,
        external_progress: tuple[Toplevel, StringVar, StringVar] | None = None,
    ) -> None:
        sped_paths = parse_selected_paths(self.contrib_sales_consult_sped_paths_var.get())
        xml_sources = parse_selected_paths(self.contrib_sales_consult_xml_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione de 1 a 12 arquivos do SPED Contribuicoes para processar.")
            return
        if len(sped_paths) > 12:
            messagebox.showwarning("Limite excedido", "A consulta aceita no maximo 12 arquivos do SPED Contribuicoes.")
            return
        owns_progress = external_progress is None
        if external_progress is None:
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Processar Consultas PIS/COFINS",
                "Preparando processamento de saidas PIS/COFINS...",
            )
        else:
            progress_dialog, progress_message_var, progress_percent_var = external_progress
        try:
            self.write_audit_log(
                "CONSULTA_CONTRIB_SAIDAS_INICIO",
                f"speds={self.format_audit_paths(sped_paths)}; xmls={self.format_audit_paths(xml_sources)}",
            )
            period_labels, comparison_rows = build_pis_cofins_period_comparison_rows(
                sped_paths,
                "Saida",
                xml_sources=xml_sources,
                progress_callback=lambda current, total, message: self.update_progress_dialog(
                    progress_dialog, progress_message_var, progress_percent_var, current, total, message
                ),
            )
            self.contrib_sales_consult_period_labels = period_labels
            self.contrib_sales_consult_period_path_map = {label: path for label, path in zip(period_labels, sped_paths)}
            self.contrib_sales_consult_comparison_rows = comparison_rows
            self.rebuild_contrib_sales_consultation_periods()
            self.refresh_contrib_sales_consultation_tree()
            self.log_message(f"Consulta de saidas PIS/COFINS carregada: {len(period_labels)} periodo(s), {len(comparison_rows)} linha(s).")
            self.write_audit_log(
                "CONSULTA_CONTRIB_SAIDAS_FIM",
                f"periodos={len(period_labels)}; linhas={len(comparison_rows)}; speds={self.format_audit_paths(sped_paths)}",
            )
            if show_success:
                messagebox.showinfo("Processar Consultas", "Consulta de saidas PIS/COFINS processada com sucesso.")
        except Exception as exc:
            self.log_message(f"Falha ao processar SPED Contribuicoes de saidas: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))
        finally:
            if owns_progress:
                self.close_progress_dialog(progress_dialog)

    def process_all_contrib_consultations(self) -> None:
        self.process_contrib_consultation_speds(sync_sales=True, show_success=True)

    def rebuild_consultation_periods(self) -> None:
        if not hasattr(self, "consult_periods_checks_box"):
            return
        for child in self.consult_periods_checks_box.winfo_children():
            child.destroy()
        self.consult_period_check_vars = {}
        for label in self.consult_period_labels:
            variable = BooleanVar(value=True)
            self.consult_period_check_vars[label] = variable
            ttk.Checkbutton(
                self.consult_periods_checks_box,
                text=label,
                variable=variable,
                command=self.handle_consult_period_selection_change,
            ).pack(anchor="w")

    def rebuild_sales_consultation_periods(self) -> None:
        if not hasattr(self, "sales_consult_periods_checks_box"):
            return
        for child in self.sales_consult_periods_checks_box.winfo_children():
            child.destroy()
        self.sales_consult_period_check_vars = {}
        for label in self.sales_consult_period_labels:
            variable = BooleanVar(value=True)
            self.sales_consult_period_check_vars[label] = variable
            ttk.Checkbutton(
                self.sales_consult_periods_checks_box,
                text=label,
                variable=variable,
                command=self.handle_sales_consult_period_selection_change,
            ).pack(anchor="w")

    def rebuild_contrib_consultation_periods(self) -> None:
        if not hasattr(self, "contrib_consult_periods_checks_box"):
            return
        for child in self.contrib_consult_periods_checks_box.winfo_children():
            child.destroy()
        self.contrib_consult_period_check_vars = {}
        for label in self.contrib_consult_period_labels:
            variable = BooleanVar(value=True)
            self.contrib_consult_period_check_vars[label] = variable
            ttk.Checkbutton(
                self.contrib_consult_periods_checks_box,
                text=label,
                variable=variable,
                command=self.handle_contrib_consult_period_selection_change,
            ).pack(anchor="w")

    def rebuild_contrib_sales_consultation_periods(self) -> None:
        if not hasattr(self, "contrib_sales_consult_periods_checks_box"):
            return
        for child in self.contrib_sales_consult_periods_checks_box.winfo_children():
            child.destroy()
        self.contrib_sales_consult_period_check_vars = {}
        for label in self.contrib_sales_consult_period_labels:
            variable = BooleanVar(value=True)
            self.contrib_sales_consult_period_check_vars[label] = variable
            ttk.Checkbutton(
                self.contrib_sales_consult_periods_checks_box,
                text=label,
                variable=variable,
                command=self.handle_contrib_sales_consult_period_selection_change,
            ).pack(anchor="w")

    def handle_consult_period_selection_change(self, _event: object | None = None) -> None:
        pending_job = getattr(self, "_consult_period_refresh_job", None)
        if pending_job:
            try:
                self.root.after_cancel(pending_job)
            except Exception:
                pass
        self._consult_period_refresh_job = self.root.after(250, self.refresh_consultation_tree)

    def handle_sales_consult_period_selection_change(self, _event: object | None = None) -> None:
        pending_job = getattr(self, "_sales_consult_period_refresh_job", None)
        if pending_job:
            try:
                self.root.after_cancel(pending_job)
            except Exception:
                pass
        self._sales_consult_period_refresh_job = self.root.after(250, self.refresh_sales_consultation_tree)

    def handle_contrib_consult_period_selection_change(self, _event: object | None = None) -> None:
        pending_job = getattr(self, "_contrib_consult_period_refresh_job", None)
        if pending_job:
            try:
                self.root.after_cancel(pending_job)
            except Exception:
                pass
        self._contrib_consult_period_refresh_job = self.root.after(250, self.refresh_contrib_consultation_tree)

    def handle_contrib_sales_consult_period_selection_change(self, _event: object | None = None) -> None:
        pending_job = getattr(self, "_contrib_sales_consult_period_refresh_job", None)
        if pending_job:
            try:
                self.root.after_cancel(pending_job)
            except Exception:
                pass
        self._contrib_sales_consult_period_refresh_job = self.root.after(250, self.refresh_contrib_sales_consultation_tree)

    def get_selected_consult_periods(self) -> set[str]:
        if not hasattr(self, "consult_period_check_vars"):
            return set()
        selected_labels = {
            label
            for label, variable in self.consult_period_check_vars.items()
            if bool(variable.get())
        }
        if not selected_labels:
            return set(self.consult_period_labels)
        return selected_labels

    def get_selected_sales_consult_periods(self) -> set[str]:
        if not hasattr(self, "sales_consult_period_check_vars"):
            return set()
        selected_labels = {
            label
            for label, variable in self.sales_consult_period_check_vars.items()
            if bool(variable.get())
        }
        if not selected_labels:
            return set(self.sales_consult_period_labels)
        return selected_labels

    def get_selected_contrib_consult_periods(self) -> set[str]:
        if not hasattr(self, "contrib_consult_period_check_vars"):
            return set()
        selected_labels = {label for label, variable in self.contrib_consult_period_check_vars.items() if bool(variable.get())}
        return selected_labels or set(self.contrib_consult_period_labels)

    def get_selected_contrib_sales_consult_periods(self) -> set[str]:
        if not hasattr(self, "contrib_sales_consult_period_check_vars"):
            return set()
        selected_labels = {label for label, variable in self.contrib_sales_consult_period_check_vars.items() if bool(variable.get())}
        return selected_labels or set(self.contrib_sales_consult_period_labels)

    def get_consultation_filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.get_selected_consult_periods()
        cst_filters = parse_filter_values(self.consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.consult_cfop_filter_var.get())
        status_filter = self.consult_status_filter_var.get().strip()
        search_text = normalize_text(self.consult_search_var.get())

        filtered_rows: list[dict[str, object]] = []
        for row in self.consult_comparison_rows:
            if selected_periods and str(row["period"]) not in selected_periods:
                continue
            if cst_filters:
                row_csts = {part.strip() for part in str(row["cst_icms"]).split("|") if part.strip()}
                if not row_csts.intersection(cst_filters):
                    continue
            if cfop_filters:
                row_cfops = {part.strip() for part in str(row["cfop"]).split("|") if part.strip()}
                if not row_cfops.intersection(cfop_filters):
                    continue
            status_value = str(row["status"]).strip()
            if status_filter == "Ok" and status_value != "Ok":
                continue
            if status_filter == "Sem credito" and "Sem credito" not in status_value:
                continue
            if status_filter == "Sem entrada" and status_value != "Sem entrada":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem entrada"} and "Sem credito" not in status_value:
                continue
            if search_text:
                searchable = normalize_text(f"{row['code']} {row['description']} {row.get('cest', '')} {row.get('suppliers', '')}")
                if search_text not in searchable:
                    continue
            if self.consult_selected_product_code and str(row["code"]).strip() != self.consult_selected_product_code:
                continue
            filtered_rows.append(row)
        return filtered_rows

    def get_sales_consultation_filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.get_selected_sales_consult_periods()
        cst_filters = parse_filter_values(self.sales_consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.sales_consult_cfop_filter_var.get())
        status_filter = self.sales_consult_status_filter_var.get().strip()
        search_text = normalize_text(self.sales_consult_search_var.get())

        filtered_rows: list[dict[str, object]] = []
        for row in self.sales_consult_comparison_rows:
            if selected_periods and str(row["period"]) not in selected_periods:
                continue
            if cst_filters:
                row_csts = {part.strip() for part in str(row["cst_icms"]).split("|") if part.strip()}
                if not row_csts.intersection(cst_filters):
                    continue
            if cfop_filters:
                row_cfops = {part.strip() for part in str(row["cfop"]).split("|") if part.strip()}
                if not row_cfops.intersection(cfop_filters):
                    continue
            status_value = str(row["status"]).strip()
            if status_filter == "Ok" and status_value != "Ok":
                continue
            if status_filter == "Sem debito" and "Sem debito" not in status_value:
                continue
            if status_filter == "Sem saida" and status_value != "Sem saida":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem saida"} and "Sem debito" not in status_value:
                continue
            if search_text:
                searchable = normalize_text(f"{row['code']} {row['description']} {row.get('cest', '')} {row.get('suppliers', '')}")
                if search_text not in searchable:
                    continue
            if self.sales_consult_selected_product_code and str(row["code"]).strip() != self.sales_consult_selected_product_code:
                continue
            filtered_rows.append(row)
        return filtered_rows

    def get_contrib_consultation_filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.get_selected_contrib_consult_periods()
        cst_filters = parse_filter_values(self.contrib_consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.contrib_consult_cfop_filter_var.get())
        status_filter = self.contrib_consult_status_filter_var.get().strip()
        search_text = normalize_text(self.contrib_consult_search_var.get())

        filtered_rows: list[dict[str, object]] = []
        for row in self.contrib_consult_comparison_rows:
            if selected_periods and str(row.get("period", "")) not in selected_periods:
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
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem entrada"} and "Sem credito" not in status_value:
                continue
            if search_text:
                searchable = normalize_text(f"{row.get('code', '')} {row.get('description', '')} {row.get('cest', '')} {row.get('suppliers', '')}")
                if search_text not in searchable:
                    continue
            if self.contrib_consult_selected_product_code and str(row.get("code", "")).strip() != self.contrib_consult_selected_product_code:
                continue
            filtered_rows.append(row)
        return filtered_rows

    def get_contrib_sales_consultation_filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.get_selected_contrib_sales_consult_periods()
        cst_filters = parse_filter_values(self.contrib_sales_consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.contrib_sales_consult_cfop_filter_var.get())
        status_filter = self.contrib_sales_consult_status_filter_var.get().strip()
        search_text = normalize_text(self.contrib_sales_consult_search_var.get())

        filtered_rows: list[dict[str, object]] = []
        for row in self.contrib_sales_consult_comparison_rows:
            if selected_periods and str(row.get("period", "")) not in selected_periods:
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
            if status_filter == "Sem debito" and "Sem debito" not in status_value:
                continue
            if status_filter == "Sem saida" and status_value != "Sem saida":
                continue
            if status_filter == "Com divergencia" and status_value in {"Ok", "Sem saida"} and "Sem debito" not in status_value:
                continue
            if search_text:
                searchable = normalize_text(f"{row.get('code', '')} {row.get('description', '')} {row.get('cest', '')} {row.get('suppliers', '')}")
                if search_text not in searchable:
                    continue
            if self.contrib_sales_consult_selected_product_code and str(row.get("code", "")).strip() != self.contrib_sales_consult_selected_product_code:
                continue
            filtered_rows.append(row)
        return filtered_rows

    def build_consultation_summary_rows(self, filtered_rows: list[dict[str, object]]) -> list[dict[str, object]]:
        grouped_rows: dict[str, dict[str, object]] = {}
        totals_by_code: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
        for row in filtered_rows:
            code = str(row["code"])
            bucket = grouped_rows.setdefault(
                code,
                {
                    "code": code,
                    "description": str(row["description"]),
                    "periods": set(),
                    "cfops": set(),
                    "csts": set(),
                    "cests": set(),
                    "suppliers": {},
                    "nominal_icms_rate": Decimal("0"),
                    "sale_value": Decimal("0"),
                    "base_icms": Decimal("0"),
                    "icms_value": Decimal("0"),
                },
            )
            if str(row["description"]).strip() and not bucket["description"]:
                bucket["description"] = str(row["description"])
            bucket["periods"].add(str(row["period"]))
            for cfop in str(row["cfop"]).split("|"):
                cfop_text = cfop.strip()
                if cfop_text:
                    bucket["cfops"].add(cfop_text)
            for cst in str(row["cst_icms"]).split("|"):
                cst_text = cst.strip()
                if cst_text:
                    bucket["csts"].add(cst_text)
            for cest in str(row.get("cest", "")).split("|"):
                cest_text = cest.strip()
                if cest_text:
                    bucket["cests"].add(cest_text)
            for detail in row.get("launch_details", []):
                supplier_name = str(detail.get("participant_name", "")).strip()
                supplier_tax_id = str(detail.get("participant_tax_id", "")).strip()
                supplier_key = supplier_name or supplier_tax_id
                if not supplier_key:
                    continue
                supplier_bucket = bucket["suppliers"].setdefault(
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
                supplier_bucket["periods"].add(str(row["period"]))
                document_identity = normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                if document_identity:
                    supplier_bucket["document_keys"].add(document_identity)
                supplier_bucket["launch_count"] += 1
                supplier_bucket["sale_value"] += Decimal(detail.get("sale_value", Decimal("0")))
                supplier_bucket["base_icms"] += Decimal(detail.get("base_icms", Decimal("0")))
                supplier_bucket["icms_value"] += Decimal(detail.get("icms_value", Decimal("0")))
            if Decimal(bucket["nominal_icms_rate"]) <= 0:
                bucket["nominal_icms_rate"] = extract_representative_icms_rate(row)
            bucket["sale_value"] += Decimal(row["sale_value"])
            bucket["base_icms"] += Decimal(row["base_icms"])
            bucket["icms_value"] += Decimal(row["icms_value"])
            totals_by_code[code] += Decimal(row["sale_value"])

        curve_labels = calculate_abc_curve_labels(totals_by_code)
        summary_rows: list[dict[str, object]] = []
        for item in grouped_rows.values():
            supplier_items = sorted(
                item["suppliers"].values(),
                key=lambda current: (normalize_text(str(current["name"])), normalize_text(str(current["tax_id"]))),
            )
            supplier_names = [
                str(current["name"]).strip() or str(current["tax_id"]).strip()
                for current in supplier_items
                if str(current["name"]).strip() or str(current["tax_id"]).strip()
            ]
            summary_rows.append(
                {
                    "code": item["code"],
                    "description": item["description"],
                    "curve_abc": curve_labels.get(str(item["code"]).strip(), "C"),
                    "suppliers": " | ".join(supplier_names),
                    "supplier_count": len(supplier_items),
                    "supplier_details": supplier_items,
                    "periods": len(item["periods"]),
                    "cfops": " | ".join(sorted(item["cfops"])),
                    "csts": " | ".join(sorted(item["csts"])),
                    "cests": " | ".join(sorted(item["cests"])),
                    "nominal_icms_rate": Decimal(item["nominal_icms_rate"]),
                    "icms_rate": compute_display_icms_rate(
                        Decimal("0"),
                        item["sale_value"],
                        item["base_icms"],
                        item["icms_value"],
                    ),
                    "has_icms_reduction": any(
                        bool(row.get("has_icms_reduction"))
                        for row in filtered_rows
                        if str(row["code"]) == str(item["code"])
                    ),
                    "sale_value": item["sale_value"],
                    "base_icms": item["base_icms"],
                    "icms_value": item["icms_value"],
                }
            )
        return summary_rows

    def build_effective_consultation_row(
        self,
        row: dict[str, object],
        cst_filters: set[str],
        cfop_filters: set[str],
    ) -> dict[str, object] | None:
        launch_details = list(row.get("launch_details", []))
        if not launch_details:
            return dict(row)

        if not cst_filters and not cfop_filters:
            return dict(row)

        filtered_launch_details: list[dict[str, object]] = []
        for detail in launch_details:
            detail_cst = str(detail.get("cst_icms", "")).strip()
            detail_cfop = str(detail.get("cfop", "")).strip()
            if cst_filters and detail_cst not in cst_filters:
                continue
            if cfop_filters and detail_cfop not in cfop_filters:
                continue
            filtered_launch_details.append(detail)

        if not filtered_launch_details:
            return None

        cfops = sorted({str(detail.get("cfop", "")).strip() for detail in filtered_launch_details if str(detail.get("cfop", "")).strip()})
        csts = sorted({str(detail.get("cst_icms", "")).strip() for detail in filtered_launch_details if str(detail.get("cst_icms", "")).strip()})
        document_keys = {
            normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            for detail in filtered_launch_details
            if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
        }
        adjusted_row = dict(row)
        adjusted_row["cfop"] = " | ".join(cfops)
        adjusted_row["cst_icms"] = " | ".join(csts)
        adjusted_row["cest"] = format_cest_values(filtered_launch_details)
        adjusted_row["suppliers"] = " | ".join(
            sorted(
                {
                    str(detail.get("participant_name", "")).strip()
                    for detail in filtered_launch_details
                    if str(detail.get("participant_name", "")).strip()
                }
            )
        )
        adjusted_row["supplier_count"] = len(
            {
                (
                    str(detail.get("participant_name", "")).strip(),
                    str(detail.get("participant_tax_id", "")).strip(),
                )
                for detail in filtered_launch_details
                if str(detail.get("participant_name", "")).strip() or str(detail.get("participant_tax_id", "")).strip()
            }
        )
        adjusted_row["quantity"] = sum((Decimal(detail.get("quantity", Decimal("0"))) for detail in filtered_launch_details), Decimal("0"))
        adjusted_row["sale_value"] = sum((Decimal(detail.get("sale_value", Decimal("0"))) for detail in filtered_launch_details), Decimal("0"))
        adjusted_row["base_icms"] = sum((Decimal(detail.get("base_icms", Decimal("0"))) for detail in filtered_launch_details), Decimal("0"))
        adjusted_row["icms_value"] = sum((Decimal(detail.get("icms_value", Decimal("0"))) for detail in filtered_launch_details), Decimal("0"))
        adjusted_row["display_icms_rate"] = compute_display_icms_rate(
            next((Decimal(detail.get("icms_rate", Decimal("0"))) for detail in filtered_launch_details if Decimal(detail.get("icms_rate", Decimal("0"))) > 0), Decimal("0")),
            adjusted_row["sale_value"],
            adjusted_row["base_icms"],
            adjusted_row["icms_value"],
        )
        adjusted_row["document_count"] = len(document_keys)
        adjusted_row["launch_count"] = len(filtered_launch_details)
        adjusted_row["launch_details"] = filtered_launch_details
        adjusted_row["status"] = summarize_entry_analysis(filtered_launch_details)["status"]
        adjusted_row["has_icms_reduction"] = any(
            has_icms_reduction(
                detail.get("icms_rate", Decimal("0")),
                detail.get("sale_value", Decimal("0")),
                detail.get("base_icms", Decimal("0")),
                detail.get("icms_value", Decimal("0")),
            )
            for detail in filtered_launch_details
        )
        return adjusted_row

    def build_effective_sales_consultation_row(
        self,
        row: dict[str, object],
        cst_filters: set[str],
        cfop_filters: set[str],
    ) -> dict[str, object] | None:
        adjusted_row = self.build_effective_consultation_row(row, cst_filters, cfop_filters)
        if adjusted_row is None:
            return None
        adjusted_row["status"] = summarize_sale_analysis(list(adjusted_row.get("launch_details", [])))["status"]
        return adjusted_row

    def get_consultation_sort_value(self, row: dict[str, object], column_id: str) -> object:
        value = row.get(column_id, "")
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return value
        return normalize_text(str(value))

    def copy_treeview_selection(self, tree: ttk.Treeview) -> str:
        selected = tree.selection()
        if not selected:
            return "break"
        lines: list[str] = []
        for item_id in selected:
            values = tree.item(item_id, "values")
            lines.append("\t".join(str(value) for value in values))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.log_message(f"{len(selected)} linha(s) copiada(s) para a area de transferencia.")
        return "break"

    def copy_text_to_clipboard(self, text: str, message: str = "Texto copiado para a area de transferencia.") -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.log_message(message)
        if hasattr(self, "compare_status_var"):
            self.compare_status_var.set(message)

    def show_treeview_copy_menu(self, event: object, tree: ttk.Treeview) -> str:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if row_id and row_id not in tree.selection():
            tree.selection_set(row_id)
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: self.copy_treeview_selection(tree))
        try:
            menu.tk_popup(getattr(event, "x_root", 0), getattr(event, "y_root", 0))
        finally:
            menu.grab_release()
        return "break"

    def show_consult_runtime_rule_menu(
        self,
        event: object,
        tree: ttk.Treeview,
        rows: list[dict[str, object]],
    ) -> str:
        row_id = tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return self.show_treeview_copy_menu(event, tree)
        tree.selection_set(row_id)
        try:
            row_index = int(row_id)
        except ValueError:
            return self.show_treeview_copy_menu(event, tree)
        if row_index < 0 or row_index >= len(rows):
            return self.show_treeview_copy_menu(event, tree)

        row = rows[row_index]
        menu = Menu(tree, tearoff=0)
        has_matching_rule = self.add_matching_runtime_rule_menu_items(menu, row)
        if has_matching_rule:
            menu.add_separator()
        menu.add_command(label="Copiar Selecao", command=lambda: self.copy_treeview_selection(tree))
        try:
            menu.tk_popup(getattr(event, "x_root", 0), getattr(event, "y_root", 0))
        finally:
            menu.grab_release()
        return "break"

    def treeview_sort_key(self, value: object) -> object:
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        text = str(value or "").strip()
        if not text:
            return (3, "")
        if re.fullmatch(r"\d{2}/\d{4}", text):
            return (0, *period_label_sort_key(text))
        parsed_date = parse_sped_document_date(text)
        if parsed_date is not None:
            return (1, parsed_date.year, parsed_date.month, parsed_date.day)
        compact_text = text.replace(" ", "")
        if re.fullmatch(r"-?[\d\.,]+", compact_text):
            return (2, parse_decimal(compact_text))
        return (3, normalize_text(text))

    def enable_treeview_sorting(
        self,
        tree: ttk.Treeview,
        columns: tuple[str, ...] | list[str],
        headings: dict[str, str],
    ) -> None:
        sort_state = {"column": "", "reverse": False}

        def sort_by(column_id: str) -> None:
            if sort_state["column"] == column_id:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["column"] = column_id
                sort_state["reverse"] = False
            ordered_items = sorted(
                tree.get_children(""),
                key=lambda item_id: self.treeview_sort_key(tree.set(item_id, column_id)),
                reverse=bool(sort_state["reverse"]),
            )
            for row_index, item_id in enumerate(ordered_items):
                tree.move(item_id, "", row_index)

        for column_id in columns:
            tree.heading(
                column_id,
                text=headings[column_id],
                command=lambda current=column_id: sort_by(current),
            )

    def bind_escape_to_close(self, dialog: Toplevel) -> None:
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def set_dialog_screen_geometry(
        self,
        dialog: Toplevel,
        preferred_width: int,
        preferred_height: int,
        min_width: int,
        min_height: int,
        margin_x: int = 60,
        margin_y: int = 140,
    ) -> None:
        dialog.update_idletasks()
        screen_width = max(dialog.winfo_screenwidth(), min_width)
        screen_height = max(dialog.winfo_screenheight(), min_height)
        width = max(min_width, min(preferred_width, screen_width - margin_x))
        height = max(min_height, min(preferred_height, screen_height - margin_y))
        pos_x = max(0, (screen_width - width) // 2)
        pos_y = max(0, (screen_height - height) // 2)
        dialog.minsize(min_width, min_height)
        dialog.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

    def open_progress_dialog(self, title: str, initial_message: str = "Aguarde...") -> tuple[Toplevel, StringVar, StringVar]:
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        self.set_dialog_screen_geometry(dialog, 760, 270, 620, 240, margin_x=120, margin_y=160)

        container = ttk.Frame(dialog, padding=(22, 18, 22, 20))
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, minsize=64)

        ttk.Label(
            container,
            text="Aguarde, processando arquivos...",
            style="Title.TLabel",
            anchor="center",
            justify="center",
            font=("Segoe UI", 14, "bold"),
        ).grid(row=0, column=0, sticky="ew")
        message_var = StringVar(value=initial_message)
        percent_var = StringVar(value="0%")
        ttk.Label(
            container,
            textvariable=message_var,
            wraplength=700,
            anchor="center",
            justify="center",
            font=("Segoe UI", 11, "bold"),
        ).grid(row=1, column=0, sticky="nsew", pady=(14, 6))
        progress = ttk.Progressbar(container, orient="horizontal", mode="determinate", maximum=100)
        progress.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        ttk.Label(
            container,
            textvariable=percent_var,
            style="Strong.TLabel",
            anchor="center",
            justify="center",
            font=("Segoe UI", 12, "bold"),
        ).grid(row=3, column=0, sticky="ew", pady=(12, 0))
        dialog._progressbar = progress
        dialog.update_idletasks()
        return dialog, message_var, percent_var

    def update_progress_dialog(
        self,
        dialog: Toplevel,
        message_var: StringVar,
        percent_var: StringVar,
        current: int,
        total: int,
        message: str,
    ) -> None:
        total_safe = max(total, 1)
        percentage = int((max(0, min(current, total_safe)) * 100) / total_safe)
        message_var.set(message)
        percent_var.set(f"{percentage}%")
        progress = getattr(dialog, "_progressbar", None)
        if progress is not None:
            progress["value"] = percentage
        dialog.update_idletasks()
        self.root.update_idletasks()

    def close_progress_dialog(self, dialog: Toplevel | None) -> None:
        if dialog is None:
            return
        try:
            dialog.grab_release()
        except Exception:
            pass
        if dialog.winfo_exists():
            dialog.destroy()

    def setup_consultation_filter_traces(self) -> None:
        for variable in (
            self.consult_cst_filter_var,
            self.consult_cfop_filter_var,
            self.consult_status_filter_var,
            self.consult_search_var,
        ):
            variable.trace_add("write", lambda *_args: self.refresh_consultation_tree())

    def setup_sales_consultation_filter_traces(self) -> None:
        for variable in (
            self.sales_consult_cst_filter_var,
            self.sales_consult_cfop_filter_var,
            self.sales_consult_status_filter_var,
            self.sales_consult_search_var,
        ):
            variable.trace_add("write", lambda *_args: self.refresh_sales_consultation_tree())

    def setup_contrib_consultation_filter_traces(self) -> None:
        for variable in (
            self.contrib_consult_cst_filter_var,
            self.contrib_consult_cfop_filter_var,
            self.contrib_consult_status_filter_var,
            self.contrib_consult_search_var,
        ):
            variable.trace_add("write", lambda *_args: self.refresh_contrib_consultation_tree())

    def setup_contrib_sales_consultation_filter_traces(self) -> None:
        for variable in (
            self.contrib_sales_consult_cst_filter_var,
            self.contrib_sales_consult_cfop_filter_var,
            self.contrib_sales_consult_status_filter_var,
            self.contrib_sales_consult_search_var,
        ):
            variable.trace_add("write", lambda *_args: self.refresh_contrib_sales_consultation_tree())

    def sort_consultation_tree_by(self, column_id: str) -> None:
        if self.consult_tree_sort_column == column_id:
            self.consult_tree_sort_reverse = not self.consult_tree_sort_reverse
        else:
            self.consult_tree_sort_column = column_id
            self.consult_tree_sort_reverse = False
        self.refresh_consultation_tree()

    def sort_consultation_summary_by(self, column_id: str) -> None:
        if self.consult_summary_sort_column == column_id:
            self.consult_summary_sort_reverse = not self.consult_summary_sort_reverse
        else:
            self.consult_summary_sort_column = column_id
            self.consult_summary_sort_reverse = False
        self.refresh_consultation_tree()

    def sort_sales_consultation_tree_by(self, column_id: str) -> None:
        if self.sales_consult_tree_sort_column == column_id:
            self.sales_consult_tree_sort_reverse = not self.sales_consult_tree_sort_reverse
        else:
            self.sales_consult_tree_sort_column = column_id
            self.sales_consult_tree_sort_reverse = False
        self.refresh_sales_consultation_tree()

    def sort_sales_consultation_summary_by(self, column_id: str) -> None:
        if self.sales_consult_summary_sort_column == column_id:
            self.sales_consult_summary_sort_reverse = not self.sales_consult_summary_sort_reverse
        else:
            self.sales_consult_summary_sort_column = column_id
            self.sales_consult_summary_sort_reverse = False
        self.refresh_sales_consultation_tree()

    def sort_contrib_consultation_tree_by(self, column_id: str) -> None:
        if self.contrib_consult_tree_sort_column == column_id:
            self.contrib_consult_tree_sort_reverse = not self.contrib_consult_tree_sort_reverse
        else:
            self.contrib_consult_tree_sort_column = column_id
            self.contrib_consult_tree_sort_reverse = False
        self.refresh_contrib_consultation_tree()

    def sort_contrib_sales_consultation_tree_by(self, column_id: str) -> None:
        if self.contrib_sales_consult_tree_sort_column == column_id:
            self.contrib_sales_consult_tree_sort_reverse = not self.contrib_sales_consult_tree_sort_reverse
        else:
            self.contrib_sales_consult_tree_sort_column = column_id
            self.contrib_sales_consult_tree_sort_reverse = False
        self.refresh_contrib_sales_consultation_tree()

    def focus_selected_summary_product(self) -> None:
        if not hasattr(self, "consult_summary_tree"):
            return
        selected = self.consult_summary_tree.selection()
        if not selected:
            return
        item_values = self.consult_summary_tree.item(selected[0], "values")
        if not item_values:
            return
        self.consult_selected_product_code = str(item_values[0]).strip()
        self.consult_search_var.set(self.consult_selected_product_code)
        self.refresh_consultation_tree()
        self.log_message(f"Consulta focada no produto {self.consult_selected_product_code}.")

    def focus_selected_sales_summary_product(self) -> None:
        if not hasattr(self, "sales_consult_summary_tree"):
            return
        selected = self.sales_consult_summary_tree.selection()
        if not selected:
            return
        item_values = self.sales_consult_summary_tree.item(selected[0], "values")
        if not item_values:
            return
        self.sales_consult_selected_product_code = str(item_values[0]).strip()
        self.sales_consult_search_var.set(self.sales_consult_selected_product_code)
        self.refresh_sales_consultation_tree()
        self.log_message(f"Consulta de saidas focada no produto {self.sales_consult_selected_product_code}.")

    def center_treeview_columns(self, tree: ttk.Treeview, columns: tuple[str, ...] | list[str]) -> None:
        for column_id in columns:
            tree.column(column_id, anchor="center")

    def open_launch_origin_details_popup(
        self,
        launch_details: list[dict[str, object]],
        code: str,
        description: str,
        period: str,
    ) -> None:
        if not launch_details:
            messagebox.showwarning("Origem dos lancamentos", "Nao ha origem de lancamentos para este item.")
            return
        document_keys = {
            normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            for detail in launch_details
            if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
        }
        detail_row = {
            "period": period,
            "code": code,
            "description": description,
            "document_count": len(document_keys),
            "launch_count": len(launch_details),
            "launch_details": [dict(detail) for detail in launch_details],
        }
        self.open_consultation_launch_details(detail_row)

    def open_summary_product_origin_documents(
        self,
        summary_row: dict[str, object],
        source_rows: list[dict[str, object]],
        operation_label: str,
    ) -> None:
        product_code = str(summary_row.get("code", "")).strip()
        if not product_code:
            messagebox.showwarning("Notas de origem", "Nao foi possivel identificar o codigo do produto.")
            return

        aggregated_launch_details: list[dict[str, object]] = []
        period_labels: set[str] = set()
        for row in source_rows:
            if str(row.get("code", "")).strip() != product_code:
                continue
            period_text = str(row.get("period", "")).strip()
            if period_text:
                period_labels.add(period_text)
            aggregated_launch_details.extend(list(row.get("launch_details", [])))

        if not aggregated_launch_details:
            messagebox.showwarning("Notas de origem", "Nao ha documentos de origem disponiveis para este item.")
            return

        aggregated_launch_details.sort(
            key=lambda detail: (
                str(detail.get("document_date", "")).strip(),
                str(detail.get("document_number", "")).strip(),
                str(detail.get("item_number", "")).strip(),
                str(detail.get("cfop", "")).strip(),
            )
        )
        document_keys = {
            normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            for detail in aggregated_launch_details
            if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
        }
        period_caption = " | ".join(sorted(period_labels, key=period_label_sort_key)) if period_labels else operation_label
        detail_row = {
            "period": period_caption,
            "code": product_code,
            "description": str(summary_row.get("description", "")).strip() or f"Produto {product_code}",
            "document_count": len(document_keys),
            "launch_count": len(aggregated_launch_details),
            "launch_details": aggregated_launch_details,
        }
        self.open_consultation_launch_details(detail_row)

    def build_supplier_popup_row_from_detail_row(self, detail_row: dict[str, object]) -> dict[str, object]:
        supplier_buckets: dict[str, dict[str, object]] = {}
        period_text = str(detail_row.get("period", "")).strip()
        for launch_detail in detail_row.get("launch_details", []):
            supplier_name = str(launch_detail.get("participant_name", "")).strip()
            supplier_tax_id = str(launch_detail.get("participant_tax_id", "")).strip()
            supplier_key = supplier_name or supplier_tax_id
            if not supplier_key:
                continue
            supplier_bucket = supplier_buckets.setdefault(
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
            if period_text:
                supplier_bucket["periods"].add(period_text)
            document_identity = normalize_document_key(str(launch_detail.get("document_key", ""))) or str(launch_detail.get("document_number", "")).strip()
            if document_identity:
                supplier_bucket["document_keys"].add(document_identity)
            supplier_bucket["launch_count"] += 1
            supplier_bucket["sale_value"] += Decimal(launch_detail.get("sale_value", Decimal("0")))
            supplier_bucket["base_icms"] += Decimal(launch_detail.get("base_icms", Decimal("0")))
            supplier_bucket["icms_value"] += Decimal(launch_detail.get("icms_value", Decimal("0")))

        return {
            "code": detail_row.get("code", ""),
            "description": detail_row.get("description", ""),
            "supplier_details": sorted(
                supplier_buckets.values(),
                key=lambda current: (normalize_text(str(current.get("name", ""))), normalize_text(str(current.get("tax_id", "")))),
            ),
        }

    def open_contrib_operation_summary_popup(self, operation_type: str, source_view: str = "contrib_consult") -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir popup de contribuicoes: {exc}")
            messagebox.showerror("Popup PIS/COFINS", str(exc))
            return

        filtered_details = list(context["filtered_details"])
        rows, _totals = build_contrib_operation_summary_rows(filtered_details, operation_type)
        grouped_launch_details = build_contrib_operation_launch_details_map(filtered_details, operation_type)
        if not rows:
            messagebox.showwarning("Popup PIS/COFINS", f"Nao ha dados de {operation_type.lower()} para os filtros atuais.")
            return

        enriched_rows: list[dict[str, object]] = []
        for row in rows:
            key = (
                str(row.get("cst_pis", "")).strip(),
                str(row.get("cst_cofins", "")).strip(),
                str(row.get("cfop", "")).strip(),
                Decimal(row.get("aliquota_pis", Decimal("0"))).quantize(Decimal("0.01")),
                Decimal(row.get("aliquota_cofins", Decimal("0"))).quantize(Decimal("0.01")),
            )
            launch_details = list(grouped_launch_details.get(key, []))
            document_keys = {
                normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
                for detail in launch_details
                if normalize_document_key(str(detail.get("document_key", ""))) or str(detail.get("document_number", "")).strip()
            }
            enriched_rows.append({**row, "launch_details": launch_details, "document_count": len(document_keys)})
        rows = enriched_rows

        dialog = Toplevel(self.root)
        dialog.title(f"Resumo {operation_type}s PIS/COFINS")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1280, 560, 1040, 460, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(container, text=f"Resumo - {operation_type}s por CST, CFOP e Aliquota", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text=self.format_contrib_popup_filter_caption(context), wraplength=1220).grid(row=1, column=0, sticky="w", pady=(6, 10))

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=1)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=1)
        filter_box.columnconfigure(4, weight=2)
        cst_var = StringVar()
        cfop_var = StringVar()
        pis_rate_var = StringVar()
        cofins_rate_var = StringVar()
        search_var = StringVar()
        for column_index, (label_text, variable) in enumerate((
            ("Filtro CST", cst_var),
            ("Filtro CFOP", cfop_var),
            ("Aliq PIS", pis_rate_var),
            ("Aliq COFINS", cofins_rate_var),
            ("Busca", search_var),
        )):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 4 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=3, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("cst_pis", "cst_cofins", "cfop", "aliquota_pis", "aliquota_cofins", "sale_value", "base_pis", "pis_value", "base_cofins", "cofins_value", "documents", "launch_count")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=14, selectmode="extended")
        headings = {
            "cst_pis": "CST PIS",
            "cst_cofins": "CST COFINS",
            "cfop": "CFOP",
            "aliquota_pis": "Aliq PIS",
            "aliquota_cofins": "Aliq COFINS",
            "sale_value": "Valor Operacao",
            "base_pis": "Base PIS",
            "pis_value": "Valor PIS",
            "base_cofins": "Base COFINS",
            "cofins_value": "Valor COFINS",
            "documents": "Docs",
            "launch_count": "Lanc.",
        }
        widths = {"cst_pis": 85, "cst_cofins": 95, "cfop": 85, "aliquota_pis": 90, "aliquota_cofins": 100, "sale_value": 120, "base_pis": 110, "pis_value": 110, "base_cofins": 120, "cofins_value": 120, "documents": 70, "launch_count": 70}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        total_rows_var = StringVar(value="Linhas: 0")
        total_sale_var = StringVar(value="Valor Operacao: 0,00")
        total_base_pis_var = StringVar(value="Base PIS: 0,00")
        total_pis_var = StringVar(value="Valor PIS: 0,00")
        total_base_cofins_var = StringVar(value="Base COFINS: 0,00")
        total_cofins_var = StringVar(value="Valor COFINS: 0,00")
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_pis_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_pis_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_cofins_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_cofins_var, style="Strong.TLabel").pack(side=LEFT)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=5, column=0, sticky="w", pady=(10, 0))

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)
            cst_filter = normalize_text(cst_var.get())
            cfop_filter = normalize_text(cfop_var.get())
            pis_rate_filter = normalize_text(pis_rate_var.get())
            cofins_rate_filter = normalize_text(cofins_rate_var.get())
            search_filter = normalize_text(search_var.get())
            filtered_rows: list[dict[str, object]] = []
            total_sale = Decimal("0")
            total_base_pis = Decimal("0")
            total_pis = Decimal("0")
            total_base_cofins = Decimal("0")
            total_cofins = Decimal("0")
            export_rows: list[list[object]] = []
            for row in rows:
                cst_text = f"{row['cst_pis']} {row['cst_cofins']}"
                pis_rate_text = format_decimal_sped(Decimal(row["aliquota_pis"]))
                cofins_rate_text = format_decimal_sped(Decimal(row["aliquota_cofins"]))
                search_text = normalize_text(
                    f"{cst_text} {row['cfop']} {pis_rate_text} {cofins_rate_text} {format_decimal_sped(Decimal(row['sale_value']))}"
                )
                if cst_filter and cst_filter not in normalize_text(cst_text):
                    continue
                if cfop_filter and cfop_filter not in normalize_text(str(row["cfop"])):
                    continue
                if pis_rate_filter and pis_rate_filter not in normalize_text(pis_rate_text):
                    continue
                if cofins_rate_filter and cofins_rate_filter not in normalize_text(cofins_rate_text):
                    continue
                if search_filter and search_filter not in search_text:
                    continue
                row_index = len(filtered_rows)
                tree.insert("", END, iid=str(row_index), values=(row["cst_pis"], row["cst_cofins"], row["cfop"], pis_rate_text, cofins_rate_text, format_decimal_sped(Decimal(row["sale_value"])), format_decimal_sped(Decimal(row["base_pis"])), format_decimal_sped(Decimal(row["pis_value"])), format_decimal_sped(Decimal(row["base_cofins"])), format_decimal_sped(Decimal(row["cofins_value"])), row["document_count"], row["launch_count"]))
                filtered_rows.append(row)
                total_sale += Decimal(row["sale_value"])
                total_base_pis += Decimal(row["base_pis"])
                total_pis += Decimal(row["pis_value"])
                total_base_cofins += Decimal(row["base_cofins"])
                total_cofins += Decimal(row["cofins_value"])
                export_rows.append([row["cst_pis"], row["cst_cofins"], row["cfop"], Decimal(row["aliquota_pis"]), Decimal(row["aliquota_cofins"]), Decimal(row["sale_value"]), Decimal(row["base_pis"]), Decimal(row["pis_value"]), Decimal(row["base_cofins"]), Decimal(row["cofins_value"]), row["document_count"], row["launch_count"]])
            tree._filtered_operation_rows = filtered_rows
            tree._filtered_popup_rows = export_rows
            total_rows_var.set(f"Linhas: {len(filtered_rows)}")
            total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            total_base_pis_var.set(f"Base PIS: {format_decimal_sped(total_base_pis)}")
            total_pis_var.set(f"Valor PIS: {format_decimal_sped(total_pis)}")
            total_base_cofins_var.set(f"Base COFINS: {format_decimal_sped(total_base_cofins)}")
            total_cofins_var.set(f"Valor COFINS: {format_decimal_sped(total_cofins)}")

        def handle_double_click(_event: object) -> None:
            row_id = tree.identify_row(getattr(_event, "y", 0))
            if not row_id:
                return
            filtered_rows = getattr(tree, "_filtered_operation_rows", rows)
            try:
                row_index = int(row_id)
            except ValueError:
                return
            if row_index < 0 or row_index >= len(filtered_rows):
                return
            detail_row = filtered_rows[row_index]
            self.open_contrib_consultation_launch_details(
                {
                    "code": f"{detail_row.get('cst_pis', '')}/{detail_row.get('cst_cofins', '')}/{detail_row.get('cfop', '')}",
                    "description": "Agrupamento fiscal PIS/COFINS",
                    "period": "",
                    "document_count": detail_row.get("document_count", 0),
                    "launch_count": detail_row.get("launch_count", 0),
                    "launch_details": list(detail_row.get("launch_details", [])),
                }
            )

        tree.bind("<Double-1>", handle_double_click)
        for variable in (cst_var, cfop_var, pis_rate_var, cofins_rate_var, search_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()
        ttk.Button(export_actions, text="Exportar Popup Excel", style="Secondary.TButton", command=lambda: self.export_simple_popup_dataset("Exportar popup", f"popup_pis_cofins_{operation_type.lower()}", f"Resumo {operation_type}s PIS COFINS", ["CST PIS", "CST COFINS", "CFOP", "Aliq PIS", "Aliq COFINS", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS", "Docs", "Lanc."], list(getattr(tree, "_filtered_popup_rows", [])), "xlsx")).pack(side=LEFT)
        ttk.Button(export_actions, text="Exportar Popup CSV", style="Secondary.TButton", command=lambda: self.export_simple_popup_dataset("Exportar popup", f"popup_pis_cofins_{operation_type.lower()}", f"Resumo {operation_type}s PIS COFINS", ["CST PIS", "CST COFINS", "CFOP", "Aliq PIS", "Aliq COFINS", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS", "Docs", "Lanc."], list(getattr(tree, "_filtered_popup_rows", [])), "csv")).pack(side=LEFT, padx=(8, 0))

    def open_contrib_diagnostic_popup(self, source_view: str = "contrib_consult", operation_type: str = "Entrada") -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir diagnostico de contribuicoes: {exc}")
            messagebox.showerror("Diagnostico PIS/COFINS", str(exc))
            return

        filtered_rows = list(context["filtered_rows"])
        if not filtered_rows:
            messagebox.showwarning("Diagnostico PIS/COFINS", "Nao ha dados para os filtros atuais.")
            return

        grouped: dict[str, dict[str, object]] = {}
        for row in filtered_rows:
            status = str(row.get("status", "")).strip() or "Sem status"
            bucket = grouped.setdefault(status, {"status": status, "lines": 0, "products": set(), "sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
            bucket["lines"] += 1
            bucket["products"].add(str(row.get("code", "")).strip())
            bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
            bucket["base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
            bucket["pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
            bucket["base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
            bucket["cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

        rows = sorted(grouped.values(), key=lambda item: normalize_text(str(item.get("status", ""))))
        dialog = Toplevel(self.root)
        dialog.title(f"Diagnostico PIS/COFINS - {operation_type}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1160, 520, 960, 420, margin_y=180)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)
        ttk.Label(container, text=f"Diagnostico PIS/COFINS - {operation_type}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text=self.format_contrib_popup_filter_caption(context), wraplength=1120).grid(row=1, column=0, sticky="w", pady=(6, 10))
        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("status", "lines", "products", "sale_value", "base_pis", "pis_value", "base_cofins", "cofins_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=14, selectmode="extended")
        headings = {"status": "Status", "lines": "Linhas", "products": "Produtos", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "pis_value": "Valor PIS", "base_cofins": "Base COFINS", "cofins_value": "Valor COFINS"}
        widths = {"status": 240, "lines": 80, "products": 80, "sale_value": 120, "base_pis": 110, "pis_value": 110, "base_cofins": 120, "cofins_value": 120}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        for row in rows:
            tree.insert("", END, values=(row["status"], row["lines"], len(row["products"]), format_decimal_sped(Decimal(row["sale_value"])), format_decimal_sped(Decimal(row["base_pis"])), format_decimal_sped(Decimal(row["pis_value"])), format_decimal_sped(Decimal(row["base_cofins"])), format_decimal_sped(Decimal(row["cofins_value"]))))

    def open_contrib_diagnostic_period_comparison_popup(self, source_view: str = "contrib_consult", operation_type: str = "Entrada") -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir comparativo diagnostico de contribuicoes: {exc}")
            messagebox.showerror("Diagnostico PIS/COFINS", str(exc))
            return

        filtered_rows = list(context["filtered_rows"])
        if not filtered_rows:
            messagebox.showwarning("Diagnostico PIS/COFINS", "Nao ha dados para os filtros atuais.")
            return

        grouped: dict[tuple[str, str], dict[str, object]] = {}
        for row in filtered_rows:
            key = (str(row.get("period", "")).strip(), str(row.get("status", "")).strip() or "Sem status")
            bucket = grouped.setdefault(key, {"period": key[0], "status": key[1], "lines": 0, "sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
            bucket["lines"] += 1
            bucket["sale_value"] += Decimal(row.get("sale_value", Decimal("0")))
            bucket["base_pis"] += Decimal(row.get("base_pis", Decimal("0")))
            bucket["pis_value"] += Decimal(row.get("pis_value", Decimal("0")))
            bucket["base_cofins"] += Decimal(row.get("base_cofins", Decimal("0")))
            bucket["cofins_value"] += Decimal(row.get("cofins_value", Decimal("0")))

        rows = sorted(grouped.values(), key=lambda item: (period_label_sort_key(str(item.get("period", ""))), normalize_text(str(item.get("status", "")))))
        dialog = Toplevel(self.root)
        dialog.title(f"Comparativo Diagnostico PIS/COFINS - {operation_type}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1220, 540, 980, 430, margin_y=180)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)
        ttk.Label(container, text=f"Comparativo Diagnostico PIS/COFINS - {operation_type}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text=self.format_contrib_popup_filter_caption(context), wraplength=1180).grid(row=1, column=0, sticky="w", pady=(6, 10))
        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("period", "status", "lines", "sale_value", "base_pis", "pis_value", "base_cofins", "cofins_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=15, selectmode="extended")
        headings = {"period": "Periodo", "status": "Status", "lines": "Linhas", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "pis_value": "Valor PIS", "base_cofins": "Base COFINS", "cofins_value": "Valor COFINS"}
        widths = {"period": 90, "status": 240, "lines": 80, "sale_value": 120, "base_pis": 110, "pis_value": 110, "base_cofins": 120, "cofins_value": 120}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        for row in rows:
            tree.insert("", END, values=(row["period"], row["status"], row["lines"], format_decimal_sped(Decimal(row["sale_value"])), format_decimal_sped(Decimal(row["base_pis"])), format_decimal_sped(Decimal(row["pis_value"])), format_decimal_sped(Decimal(row["base_cofins"])), format_decimal_sped(Decimal(row["cofins_value"]))))

    def open_contrib_product_monthly_summary_popup(self, operation_type: str, source_view: str) -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir curva ABC de contribuicoes: {exc}")
            messagebox.showerror("Levantamento por produto", str(exc))
            return
        filtered_rows = list(context["filtered_rows"])
        periods, export_headers, display_rows, export_rows = build_contrib_product_monthly_linear_dataset(filtered_rows, operation_type)
        if not display_rows:
            messagebox.showwarning("Levantamento por produto", "Nao ha dados para os filtros atuais.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Levantamento por Produto - {operation_type}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1360, 620, 1100, 500, margin_y=190)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)
        ttk.Label(container, text=f"Levantamento por Produto/Mes - {operation_type}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text="Resultado linear por produto com totais de PIS e COFINS por periodo.", wraplength=1320).grid(row=1, column=0, sticky="w", pady=(6, 10))
        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=2)
        product_var = StringVar()
        curve_var = StringVar()
        search_var = StringVar()
        for column_index, (label_text, variable) in enumerate((("Produto", product_var), ("Curva ABC", curve_var), ("Busca", search_var))):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 2 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))
        month_band = ttk.Frame(container)
        month_band.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        if periods:
            ttk.Label(month_band, text="Meses:", style="Strong.TLabel").pack(side=LEFT, padx=(0, 8))
        for period_index, period in enumerate(periods):
            tk.Label(month_band, text=f"{period}  Oper | BPIS | VPIS | BCOF | VCOF", bg="#d9eaf7" if period_index % 2 == 0 else "#fbe7c6", fg="#184c73", padx=10, pady=4, relief="solid", bd=1).pack(side=LEFT, padx=(0, 6))
        tree_box = ttk.Frame(container)
        tree_box.grid(row=4, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ["curve_abc", "code", "description", "ncm", "cfops", "csts"]
        widths = {"curve_abc": 80, "code": 95, "description": 260, "ncm": 90, "cfops": 110, "csts": 100}
        headings = {"curve_abc": "Curva ABC", "code": "Codigo", "description": "Descricao", "ncm": "NCM", "cfops": "CFOPs", "csts": "CSTs"}
        for period in periods:
            period_key = period.replace("/", "_")
            columns.extend([f"{period_key}_sale", f"{period_key}_base_pis", f"{period_key}_pis", f"{period_key}_base_cofins", f"{period_key}_cofins"])
            headings[f"{period_key}_sale"] = f"{period} Valor"
            headings[f"{period_key}_base_pis"] = f"{period} Base PIS"
            headings[f"{period_key}_pis"] = f"{period} Valor PIS"
            headings[f"{period_key}_base_cofins"] = f"{period} Base COF"
            headings[f"{period_key}_cofins"] = f"{period} Valor COF"
            widths[f"{period_key}_sale"] = 110
            widths[f"{period_key}_base_pis"] = 110
            widths[f"{period_key}_pis"] = 110
            widths[f"{period_key}_base_cofins"] = 110
            widths[f"{period_key}_cofins"] = 110
        columns.extend(["total_sale", "total_base_pis", "total_pis", "total_base_cofins", "total_cofins"])
        headings.update({"total_sale": "Total Valor Operacao", "total_base_pis": "Total Base PIS", "total_pis": "Total Valor PIS", "total_base_cofins": "Total Base COFINS", "total_cofins": "Total Valor COFINS"})
        widths.update({"total_sale": 130, "total_base_pis": 130, "total_pis": 130, "total_base_cofins": 140, "total_cofins": 140})
        tree = ttk.Treeview(tree_box, columns=tuple(columns), show="headings", height=16, selectmode="extended")
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, tuple(columns), headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.tag_configure("curve_a", background="#f4fff4")
        tree.tag_configure("curve_b", background="#fff8e1")
        tree.tag_configure("curve_c", background="#ffe8e8")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        footer = ttk.Frame(container)
        footer.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        total_rows_var = StringVar(value="Linhas: 0")
        total_products_var = StringVar(value="Produtos: 0")
        total_sale_var = StringVar(value="Valor Operacao: 0,00")
        total_pis_var = StringVar(value="Valor PIS: 0,00")
        total_cofins_var = StringVar(value="Valor COFINS: 0,00")
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_pis_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_cofins_var, style="Strong.TLabel").pack(side=LEFT)
        export_actions = ttk.Frame(container)
        export_actions.grid(row=6, column=0, sticky="w", pady=(10, 0))

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)
            product_filter = normalize_text(product_var.get())
            curve_filter = normalize_text(curve_var.get())
            search_filter = normalize_text(search_var.get())
            filtered_products: set[str] = set()
            filtered_export_rows: list[list[object]] = []
            total_pis = Decimal("0")
            total_cofins = Decimal("0")
            total_sale = Decimal("0")
            for row in display_rows:
                values = list(row["values"])
                search_text = normalize_text(" ".join(serialize_value_for_clipboard(value) for value in values))
                if product_filter and product_filter not in normalize_text(serialize_value_for_clipboard(values[1])):
                    continue
                if curve_filter and curve_filter not in normalize_text(str(row["curve_abc"])):
                    continue
                if search_filter and search_filter not in search_text:
                    continue
                tree.insert("", END, values=tuple(serialize_value_for_clipboard(value) for value in values), tags=(f"curve_{str(row['curve_abc']).strip().lower()}",))
                filtered_products.add(str(values[1]))
                filtered_export_rows.append(values)
            for source_row in filtered_rows:
                if str(source_row.get("code", "")).strip() in filtered_products:
                    total_sale += Decimal(source_row.get("sale_value", Decimal("0")))
                    total_pis += Decimal(source_row.get("pis_value", Decimal("0")))
                    total_cofins += Decimal(source_row.get("cofins_value", Decimal("0")))
            tree._filtered_popup_rows = filtered_export_rows
            total_rows_var.set(f"Linhas: {len(filtered_export_rows)}")
            total_products_var.set(f"Produtos: {len(filtered_products)}")
            total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            total_pis_var.set(f"Valor PIS: {format_decimal_sped(total_pis)}")
            total_cofins_var.set(f"Valor COFINS: {format_decimal_sped(total_cofins)}")

        for variable in (product_var, curve_var, search_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()
        ttk.Button(export_actions, text="Exportar Popup Excel", style="Secondary.TButton", command=lambda: self.export_simple_popup_dataset("Exportar popup", f"levantamento_produto_pis_cofins_{operation_type.lower()}", f"Levantamento {operation_type}", export_headers, list(getattr(tree, "_filtered_popup_rows", [])), "xlsx")).pack(side=LEFT)
        ttk.Button(export_actions, text="Exportar Popup CSV", style="Secondary.TButton", command=lambda: self.export_simple_popup_dataset("Exportar popup", f"levantamento_produto_pis_cofins_{operation_type.lower()}", f"Levantamento {operation_type}", export_headers, list(getattr(tree, "_filtered_popup_rows", [])), "csv")).pack(side=LEFT, padx=(8, 0))

    def open_contrib_apuracao_popup(self, source_view: str = "contrib_consult") -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir apuracao de contribuicoes: {exc}")
            messagebox.showerror("Apuracao PIS/COFINS", str(exc))
            return
        details = list(context["filtered_details"])
        if not details:
            messagebox.showwarning("Apuracao PIS/COFINS", "Nao ha dados para os filtros atuais.")
            return
        grouped: dict[str, dict[str, Decimal]] = defaultdict(lambda: {"sale_value": Decimal("0"), "base_pis": Decimal("0"), "pis_value": Decimal("0"), "base_cofins": Decimal("0"), "cofins_value": Decimal("0")})
        for detail in details:
            period = str(detail.get("period", "")).strip() or "(sem periodo)"
            grouped[period]["sale_value"] += Decimal(detail.get("sale_value", Decimal("0")))
            grouped[period]["base_pis"] += Decimal(detail.get("base_pis", Decimal("0")))
            grouped[period]["pis_value"] += Decimal(detail.get("pis_value", Decimal("0")))
            grouped[period]["base_cofins"] += Decimal(detail.get("base_cofins", Decimal("0")))
            grouped[period]["cofins_value"] += Decimal(detail.get("cofins_value", Decimal("0")))
        dialog = Toplevel(self.root)
        dialog.title("Apuracao PIS/COFINS")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 980, 480, 860, 380, margin_y=180)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)
        ttk.Label(container, text="Apuracao PIS/COFINS", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text=self.format_contrib_popup_filter_caption(context), wraplength=940).grid(row=1, column=0, sticky="w", pady=(6, 10))
        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("period", "sale_value", "base_pis", "pis_value", "aliquota_pis", "base_cofins", "cofins_value", "aliquota_cofins")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=12, selectmode="extended")
        headings = {"period": "Periodo", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "pis_value": "Valor PIS", "aliquota_pis": "Aliq PIS Efetiva", "base_cofins": "Base COFINS", "cofins_value": "Valor COFINS", "aliquota_cofins": "Aliq COFINS Efetiva"}
        widths = {"period": 90, "sale_value": 120, "base_pis": 110, "pis_value": 110, "aliquota_pis": 120, "base_cofins": 120, "cofins_value": 120, "aliquota_cofins": 130}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll_y.set)
        for period in sorted(grouped.keys(), key=period_label_sort_key):
            totals = grouped[period]
            aliquota_pis = (totals["pis_value"] * Decimal("100") / totals["base_pis"]).quantize(Decimal("0.01")) if totals["base_pis"] > 0 else Decimal("0.00")
            aliquota_cofins = (totals["cofins_value"] * Decimal("100") / totals["base_cofins"]).quantize(Decimal("0.01")) if totals["base_cofins"] > 0 else Decimal("0.00")
            tree.insert("", END, values=(period, format_decimal_sped(totals["sale_value"]), format_decimal_sped(totals["base_pis"]), format_decimal_sped(totals["pis_value"]), format_decimal_sped(aliquota_pis), format_decimal_sped(totals["base_cofins"]), format_decimal_sped(totals["cofins_value"]), format_decimal_sped(aliquota_cofins)))

    def open_contrib_fiscal_documents_popup(self, source_view: str, operation_type: str) -> None:
        try:
            context = self.load_current_contrib_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir espelho de documentos de contribuicoes: {exc}")
            messagebox.showerror("Espelho de documentos", str(exc))
            return
        note_snapshots = self.collect_note_snapshots_from_details(list(context.get("filtered_details", [])), operation_type)
        if not note_snapshots:
            messagebox.showwarning("Espelho de documentos", "Nao ha documentos com espelho disponivel para os filtros atuais.")
            return
        normalized_operation = normalize_operation_type(operation_type)
        operation_caption = "Entradas" if normalized_operation == "entrada" else "Saidas"
        dialog = Toplevel(self.root)
        dialog.title(f"Espelho de Documentos - {operation_caption}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1420, 760, 1140, 620, margin_y=180)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)
        ttk.Label(container, text=f"Espelho de Documentos - {operation_caption}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(container, text="Pesquise pelo numero do documento ou pela chave. Duplo clique abre o espelho completo.", wraplength=1360).grid(row=1, column=0, sticky="w", pady=(6, 10))
        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=2)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=2)
        document_filter_var = StringVar()
        key_filter_var = StringVar()
        model_filter_var = StringVar()
        search_filter_var = StringVar()
        for column_index, (label_text, variable) in enumerate((("Numero Documento", document_filter_var), ("Chave", key_filter_var), ("Modelo", model_filter_var), ("Busca", search_filter_var))):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 3 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))
        tree_box = ttk.Frame(container)
        tree_box.grid(row=3, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("period", "document_date", "document_number", "document_series", "document_model", "participant_name", "participant_tax_id", "document_key", "items", "sale_value", "base_pis", "pis_value", "base_cofins", "cofins_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=18, selectmode="extended")
        headings = {"period": "Periodo", "document_date": "Data", "document_number": "Documento", "document_series": "Serie", "document_model": "Modelo", "participant_name": "Participante", "participant_tax_id": "CPF/CNPJ", "document_key": "Chave", "items": "Itens", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "pis_value": "Valor PIS", "base_cofins": "Base COFINS", "cofins_value": "Valor COFINS"}
        widths = {"period": 90, "document_date": 90, "document_number": 100, "document_series": 70, "document_model": 70, "participant_name": 220, "participant_tax_id": 150, "document_key": 300, "items": 60, "sale_value": 120, "base_pis": 120, "pis_value": 120, "base_cofins": 120, "cofins_value": 120}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)
            document_filter = normalize_text(document_filter_var.get())
            key_filter = normalize_text(key_filter_var.get())
            model_filter = normalize_text(model_filter_var.get())
            search_filter = normalize_text(search_filter_var.get())
            filtered_snapshots: list[dict[str, object]] = []
            for snapshot in note_snapshots:
                document_number = str(snapshot.get("document_number", "")).strip()
                document_key = str(snapshot.get("document_key", "")).strip()
                document_model = str(snapshot.get("document_model", "")).strip()
                participant_name = str(snapshot.get("participant_name", "")).strip()
                participant_tax_id = str(snapshot.get("participant_tax_id", "")).strip()
                search_text = normalize_text(f"{document_number} {document_key} {participant_name} {participant_tax_id} {document_model}")
                if document_filter and document_filter not in normalize_text(document_number):
                    continue
                if key_filter and key_filter not in normalize_text(document_key):
                    continue
                if model_filter and model_filter not in normalize_text(document_model):
                    continue
                if search_filter and search_filter not in search_text:
                    continue
                snapshot_items = list(snapshot.get("items", []))
                sale_value = sum((Decimal(item.get("sale_value", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                base_pis = sum((Decimal(item.get("base_pis", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                pis_value = sum((Decimal(item.get("pis_value", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                base_cofins = sum((Decimal(item.get("base_cofins", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                cofins_value = sum((Decimal(item.get("cofins_value", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                row_index = len(filtered_snapshots)
                tree.insert("", END, iid=str(row_index), values=(snapshot.get("period", ""), snapshot.get("document_date", ""), document_number, snapshot.get("document_series", ""), document_model, participant_name, participant_tax_id, document_key, len(snapshot_items), format_decimal_sped(sale_value), format_decimal_sped(base_pis), format_decimal_sped(pis_value), format_decimal_sped(base_cofins), format_decimal_sped(cofins_value)))
                filtered_snapshots.append(snapshot)
            tree._filtered_note_snapshots = filtered_snapshots

        def handle_tree_double_click(event: object) -> None:
            row_id = tree.identify_row(getattr(event, "y", 0))
            if not row_id:
                return
            filtered_snapshots = getattr(tree, "_filtered_note_snapshots", note_snapshots)
            try:
                row_index = int(row_id)
            except ValueError:
                return
            if row_index < 0 or row_index >= len(filtered_snapshots):
                return
            self.open_complete_contrib_note_popup(filtered_snapshots[row_index])

        tree.bind("<Double-1>", handle_tree_double_click)
        for variable in (document_filter_var, key_filter_var, model_filter_var, search_filter_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()

    def open_complete_contrib_note_popup(self, note_snapshot: dict[str, object]) -> None:
        note_items = list(note_snapshot.get("items", []))
        if not note_items:
            messagebox.showwarning("Nota completa", "Nao ha itens disponiveis para esta nota.")
            return
        dialog = Toplevel(self.root)
        dialog.title(f"Espelho da Nota - {note_snapshot.get('document_number', '')}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1420, 760, 1140, 620, margin_y=190)
        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)
        header_box = ttk.LabelFrame(container, text="Cabecalho da Nota", padding=10)
        header_box.grid(row=0, column=0, sticky="ew")
        header_box.columnconfigure(0, weight=1)
        header_box.columnconfigure(1, weight=1)
        header_left = ttk.Frame(header_box)
        header_left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        header_right = ttk.Frame(header_box)
        header_right.grid(row=0, column=1, sticky="nw")
        for label_text, value_text in [("Documento", note_snapshot.get("document_number", "")), ("Chave", note_snapshot.get("document_key", "")), ("Data", note_snapshot.get("document_date", "")), ("Serie", note_snapshot.get("document_series", "")), ("Modelo", note_snapshot.get("document_model", ""))]:
            ttk.Label(header_left, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)
        total_pis = sum((Decimal(item.get("pis_value", Decimal("0"))) for item in note_items), Decimal("0"))
        total_cofins = sum((Decimal(item.get("cofins_value", Decimal("0"))) for item in note_items), Decimal("0"))
        for label_text, value_text in [("Participante", note_snapshot.get("participant_name", "")), ("Cod. Participante", note_snapshot.get("participant_code", "")), ("CPF/CNPJ", note_snapshot.get("participant_tax_id", "")), ("Valor PIS", format_decimal_sped(total_pis)), ("Valor COFINS", format_decimal_sped(total_cofins)), ("Itens", len(note_items))]:
            ttk.Label(header_right, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)
        export_actions = ttk.Frame(container)
        export_actions.grid(row=1, column=0, sticky="w", pady=(10, 8))
        ttk.Button(export_actions, text="Exportar Nota Excel", style="Secondary.TButton", command=lambda: self.export_complete_contrib_note(note_snapshot, "xlsx")).pack(side=LEFT)
        ttk.Button(export_actions, text="Exportar Nota CSV", style="Secondary.TButton", command=lambda: self.export_complete_contrib_note(note_snapshot, "csv")).pack(side=LEFT, padx=(8, 0))
        items_box = ttk.Frame(container)
        items_box.grid(row=3, column=0, sticky="nsew")
        items_box.columnconfigure(0, weight=1)
        items_box.rowconfigure(0, weight=1)
        columns = ("item_number", "code", "description", "ncm", "cest", "cfop", "cst_pis", "cst_cofins", "aliquota_pis", "aliquota_cofins", "quantity", "sale_value", "base_pis", "pis_value", "base_cofins", "cofins_value")
        items_tree = ttk.Treeview(items_box, columns=columns, show="headings", height=18, selectmode="extended")
        headings = {"item_number": "Item", "code": "Codigo", "description": "Descricao", "ncm": "NCM", "cest": "CEST", "cfop": "CFOP", "cst_pis": "CST PIS", "cst_cofins": "CST COFINS", "aliquota_pis": "Aliq PIS", "aliquota_cofins": "Aliq COFINS", "quantity": "Quantidade", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "pis_value": "Valor PIS", "base_cofins": "Base COFINS", "cofins_value": "Valor COFINS"}
        widths = {"item_number": 60, "code": 90, "description": 300, "ncm": 100, "cest": 90, "cfop": 80, "cst_pis": 80, "cst_cofins": 95, "aliquota_pis": 90, "aliquota_cofins": 100, "quantity": 90, "sale_value": 110, "base_pis": 110, "pis_value": 110, "base_cofins": 110, "cofins_value": 110}
        for column_id in columns:
            items_tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(items_tree, columns, headings)
        items_tree.grid(row=0, column=0, sticky="nsew")
        items_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(items_tree))
        items_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, items_tree))
        scroll_y = ttk.Scrollbar(items_box, orient="vertical", command=items_tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(items_box, orient="horizontal", command=items_tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        items_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        total_quantity = Decimal("0")
        total_sale = Decimal("0")
        total_base_pis = Decimal("0")
        total_base_cofins = Decimal("0")
        for item in note_items:
            quantity = Decimal(item.get("quantity", Decimal("0")))
            sale_value = Decimal(item.get("sale_value", Decimal("0")))
            base_pis = Decimal(item.get("base_pis", Decimal("0")))
            base_cofins = Decimal(item.get("base_cofins", Decimal("0")))
            items_tree.insert("", END, values=(item.get("item_number", ""), item.get("code", ""), item.get("description", ""), item.get("ncm", ""), item.get("cest", ""), item.get("cfop", ""), item.get("cst_pis", ""), item.get("cst_cofins", ""), format_decimal_sped(Decimal(item.get("aliquota_pis", Decimal("0")))), format_decimal_sped(Decimal(item.get("aliquota_cofins", Decimal("0")))), format_decimal_sped(quantity), format_decimal_sped(sale_value), format_decimal_sped(base_pis), format_decimal_sped(Decimal(item.get("pis_value", Decimal("0")))), format_decimal_sped(base_cofins), format_decimal_sped(Decimal(item.get("cofins_value", Decimal("0"))))))
            total_quantity += quantity
            total_sale += sale_value
            total_base_pis += base_pis
            total_base_cofins += base_cofins
        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Quantidade: {format_decimal_sped(total_quantity)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(total_sale)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base PIS: {format_decimal_sped(total_base_pis)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base COFINS: {format_decimal_sped(total_base_cofins)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor PIS: {format_decimal_sped(total_pis)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor COFINS: {format_decimal_sped(total_cofins)}", style="Strong.TLabel").pack(side=LEFT)

    def export_complete_contrib_note(self, note_snapshot: dict[str, object], output_type: str) -> None:
        note_items = list(note_snapshot.get("items", []))
        if not note_items:
            messagebox.showwarning("Exportar nota", "Nao ha itens disponiveis para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(title="Salvar nota completa", defaultextension=suffix, initialfile=f"nota_pis_cofins_{note_snapshot.get('document_number', 'sem_numero')}{suffix}", filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")])
        if not selected:
            return
        output_path = Path(selected)
        header_headers = ["Campo", "Valor"]
        header_rows = [["Documento", note_snapshot.get("document_number", "")], ["Chave", note_snapshot.get("document_key", "")], ["Data", note_snapshot.get("document_date", "")], ["Serie", note_snapshot.get("document_series", "")], ["Modelo", note_snapshot.get("document_model", "")], ["Participante", note_snapshot.get("participant_name", "")], ["Cod. Participante", note_snapshot.get("participant_code", "")], ["CPF/CNPJ", note_snapshot.get("participant_tax_id", "")], ["Quantidade de Itens", len(note_items)]]
        item_headers = ["Item", "Codigo", "Descricao", "NCM", "CEST", "CFOP", "CST PIS", "CST COFINS", "Aliq PIS", "Aliq COFINS", "Quantidade", "Valor Operacao", "Base PIS", "Valor PIS", "Base COFINS", "Valor COFINS"]
        item_rows = [[item.get("item_number", ""), item.get("code", ""), item.get("description", ""), item.get("ncm", ""), item.get("cest", ""), item.get("cfop", ""), item.get("cst_pis", ""), item.get("cst_cofins", ""), Decimal(item.get("aliquota_pis", Decimal("0"))), Decimal(item.get("aliquota_cofins", Decimal("0"))), Decimal(item.get("quantity", Decimal("0"))), Decimal(item.get("sale_value", Decimal("0"))), Decimal(item.get("base_pis", Decimal("0"))), Decimal(item.get("pis_value", Decimal("0"))), Decimal(item.get("base_cofins", Decimal("0"))), Decimal(item.get("cofins_value", Decimal("0")))] for item in note_items]
        try:
            if output_path.suffix.lower() == ".csv":
                header_output = output_path.with_name(f"{output_path.stem}_cabecalho.csv")
                write_simple_csv_file(header_output, header_headers, header_rows)
                write_simple_csv_file(output_path, item_headers, item_rows)
            else:
                write_simple_excel_workbook(output_path, [("Cabecalho", header_headers, header_rows), ("Itens", item_headers, item_rows)])
            self.log_message(f"Nota PIS/COFINS exportada com sucesso em: {output_path}")
            messagebox.showinfo("Exportar nota", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar nota PIS/COFINS: {exc}")
            messagebox.showerror("Exportar nota", str(exc))

    def open_contrib_consultation_launch_details(self, row: dict[str, object]) -> None:
        launch_details = list(row.get("launch_details", []))
        if not launch_details:
            messagebox.showwarning("Detalhamento", "Nao ha lancamentos detalhados para esta linha.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Detalhamento PIS/COFINS - {row.get('code', '')} - {row.get('period', '')}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1380, 520, 1120, 420, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        ttk.Label(
            container,
            text=(
                f"Produto {row.get('code', '')} - {row.get('description', '')} | Periodo {row.get('period', '')} | "
                f"Docs {row.get('document_count', 0)} | Lancamentos {row.get('launch_count', 0)}"
            ),
            style="Title.TLabel",
            wraplength=1300,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))
        ttk.Label(
            container,
            text="Duplo clique em um item para abrir o espelho completo da nota.",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=2, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("document_date", "document_number", "participant_name", "item_number", "cest", "cfop", "cst_pis", "cst_cofins", "aliquota_pis", "aliquota_cofins", "sale_value", "base_pis", "base_cofins", "pis_value", "cofins_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=12, selectmode="extended")
        headings = {"document_date": "Data", "document_number": "Documento", "participant_name": "Participante", "item_number": "Item", "cest": "CEST", "cfop": "CFOP", "cst_pis": "CST PIS", "cst_cofins": "CST COFINS", "aliquota_pis": "Aliq PIS", "aliquota_cofins": "Aliq COFINS", "sale_value": "Valor Operacao", "base_pis": "Base PIS", "base_cofins": "Base COFINS", "pis_value": "Valor PIS", "cofins_value": "Valor COFINS"}
        widths = {"document_date": 90, "document_number": 100, "participant_name": 220, "item_number": 60, "cest": 90, "cfop": 80, "cst_pis": 80, "cst_cofins": 95, "aliquota_pis": 85, "aliquota_cofins": 95, "sale_value": 110, "base_pis": 100, "base_cofins": 110, "pis_value": 95, "cofins_value": 110}
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        enriched_launch_details: list[dict[str, object]] = []
        for detail in launch_details:
            enriched_detail = dict(detail)
            note_snapshot = enriched_detail.get("note_snapshot")
            if not isinstance(note_snapshot, dict):
                document_identity = normalize_document_key(str(enriched_detail.get("document_key", ""))) or str(enriched_detail.get("document_number", "")).strip()
                document_date = str(enriched_detail.get("document_date", "")).strip()
                period = str(enriched_detail.get("period", row.get("period", ""))).strip()
                if document_identity and document_date:
                    for candidate in launch_details:
                        candidate_snapshot = candidate.get("note_snapshot")
                        candidate_identity = normalize_document_key(str(candidate.get("document_key", ""))) or str(candidate.get("document_number", "")).strip()
                        candidate_date = str(candidate.get("document_date", "")).strip()
                        candidate_period = str(candidate.get("period", row.get("period", ""))).strip()
                        if candidate_identity == document_identity and candidate_date == document_date and candidate_period == period and isinstance(candidate_snapshot, dict):
                            enriched_detail["note_snapshot"] = candidate_snapshot
                            break
            enriched_launch_details.append(enriched_detail)
            tree.insert(
                "",
                END,
                values=(
                    enriched_detail.get("document_date", ""),
                    enriched_detail.get("document_number", ""),
                    enriched_detail.get("participant_name", ""),
                    enriched_detail.get("item_number", ""),
                    enriched_detail.get("cest", ""),
                    enriched_detail.get("cfop", ""),
                    enriched_detail.get("cst_pis", ""),
                    enriched_detail.get("cst_cofins", ""),
                    format_decimal_sped(Decimal(enriched_detail.get("aliquota_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("aliquota_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("sale_value", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("base_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("base_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(enriched_detail.get("cofins_value", Decimal("0")))),
                ),
            )

        def handle_tree_double_click(event: object) -> None:
            row_id = tree.identify_row(getattr(event, "y", 0))
            if not row_id:
                return
            try:
                row_index = tree.index(row_id)
            except Exception:
                return
            if row_index < 0 or row_index >= len(enriched_launch_details):
                return
            note_snapshot = enriched_launch_details[row_index].get("note_snapshot")
            if not isinstance(note_snapshot, dict):
                messagebox.showwarning("Espelho da nota", "Nao foi possivel localizar o espelho completo da nota para este item.")
                return
            self.open_complete_contrib_note_popup(note_snapshot)

        tree.bind("<Double-1>", handle_tree_double_click)

    def handle_consult_summary_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "consult_summary_tree"):
            return
        row_id = self.consult_summary_tree.identify_row(getattr(event, "y", 0))
        column_id = self.consult_summary_tree.identify_column(getattr(event, "x", 0))
        if not row_id:
            return
        row_index = self.consult_summary_tree.index(row_id)
        if row_index < 0 or row_index >= len(self.consult_summary_rows):
            return
        row = self.consult_summary_rows[row_index]
        if column_id == "#4" and int(row.get("supplier_count", 0)) > 1:
            self.open_consultation_supplier_popup(row)
            return
        self.open_summary_product_origin_documents(row, list(self.consult_filtered_rows), "Entradas")

    def handle_sales_consult_summary_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "sales_consult_summary_tree"):
            return
        row_id = self.sales_consult_summary_tree.identify_row(getattr(event, "y", 0))
        column_id = self.sales_consult_summary_tree.identify_column(getattr(event, "x", 0))
        if not row_id:
            return
        row_index = self.sales_consult_summary_tree.index(row_id)
        if row_index < 0 or row_index >= len(self.sales_consult_summary_rows):
            return
        row = self.sales_consult_summary_rows[row_index]
        if column_id == "#4" and int(row.get("supplier_count", 0)) > 1:
            self.open_consultation_supplier_popup(row)
            return
        self.open_summary_product_origin_documents(row, list(self.sales_consult_filtered_rows), "Saidas")

    def open_consultation_supplier_popup(self, summary_row: dict[str, object]) -> None:
        supplier_details = list(summary_row.get("supplier_details", []))
        if not supplier_details:
            messagebox.showwarning("Fornecedores", "Nao ha fornecedores disponiveis para este produto.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Fornecedores - {summary_row.get('code', '')}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 980, 420, 860, 340, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text=f"Produto {summary_row.get('code', '')} - {summary_row.get('description', '')}",
            style="Title.TLabel",
            wraplength=920,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=1, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)
        columns = ("name", "tax_id", "periods", "documents", "launches", "sale_value", "base_icms", "icms_value")
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=12, selectmode="extended")
        headings = {
            "name": "Fornecedor",
            "tax_id": "CPF/CNPJ",
            "periods": "Periodos",
            "documents": "Docs",
            "launches": "Lanc.",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "name": 260,
            "tax_id": 140,
            "periods": 70,
            "documents": 70,
            "launches": 70,
            "sale_value": 120,
            "base_icms": 120,
            "icms_value": 120,
        }
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))
        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        for supplier in supplier_details:
            tree.insert(
                "",
                END,
                values=(
                    supplier.get("name", ""),
                    supplier.get("tax_id", ""),
                    len(supplier.get("periods", set())),
                    len(supplier.get("document_keys", set())),
                    supplier.get("launch_count", 0),
                    format_decimal_sped(Decimal(supplier.get("sale_value", Decimal("0")))),
                    format_decimal_sped(Decimal(supplier.get("base_icms", Decimal("0")))),
                    format_decimal_sped(Decimal(supplier.get("icms_value", Decimal("0")))),
                ),
            )

        export_rows = [
            [
                supplier.get("name", ""),
                supplier.get("tax_id", ""),
                len(supplier.get("periods", set())),
                len(supplier.get("document_keys", set())),
                supplier.get("launch_count", 0),
                Decimal(supplier.get("sale_value", Decimal("0"))),
                Decimal(supplier.get("base_icms", Decimal("0"))),
                Decimal(supplier.get("icms_value", Decimal("0"))),
            ]
            for supplier in supplier_details
        ]
        export_actions = ttk.Frame(container)
        export_actions.grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"fornecedores_{summary_row.get('code', 'produto')}",
                "Fornecedores",
                ["Fornecedor", "CPF/CNPJ", "Periodos", "Docs", "Lanc.", "Valor Operacao", "Base ICMS", "Valor ICMS"],
                export_rows,
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_simple_popup_dataset(
                "Exportar popup",
                f"fornecedores_{summary_row.get('code', 'produto')}",
                "Fornecedores",
                ["Fornecedor", "CPF/CNPJ", "Periodos", "Docs", "Lanc.", "Valor Operacao", "Base ICMS", "Valor ICMS"],
                export_rows,
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        footer = ttk.Frame(container)
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            footer,
            text=f"Fornecedores: {len(supplier_details)}",
            style="Strong.TLabel",
        ).pack(side=LEFT)

    def open_product_monthly_summary_popup(self, operation_type: str, source_view: str) -> None:
        context = self.load_current_popup_context(source_view)
        filtered_rows = list(self.sales_consult_filtered_rows if source_view == "sales_consult" else self.consult_filtered_rows)
        periods, export_headers, display_rows, export_rows = build_product_monthly_linear_dataset(filtered_rows, operation_type)
        if not display_rows:
            messagebox.showwarning("Levantamento por produto", "Nao ha dados para os filtros atuais.")
            return

        if not bool(context.get("filter_descriptions")):
            summary_totals_by_period = build_summary_operation_totals_by_period(
                list(context.get("source_sped_paths", [])),
                operation_type,
            )
            fixed_column_count = 6
            monthly_block_size = 4
            adjustment_values: list[object] = ["", "(SEM PRODUTO)", "Ajuste fiscal do resumo do SPED", "", "", ""]
            has_adjustment = False
            total_adjustment_sale = Decimal("0")
            total_adjustment_base = Decimal("0")
            total_adjustment_icms = Decimal("0")
            for period_index, period in enumerate(periods):
                period_products_sale = sum(
                    (
                        Decimal(row[fixed_column_count + (period_index * monthly_block_size)])
                        for row in export_rows
                    ),
                    Decimal("0"),
                )
                period_products_base = sum(
                    (
                        Decimal(row[fixed_column_count + (period_index * monthly_block_size) + 1])
                        for row in export_rows
                    ),
                    Decimal("0"),
                )
                period_products_icms = sum(
                    (
                        Decimal(row[fixed_column_count + (period_index * monthly_block_size) + 2])
                        for row in export_rows
                    ),
                    Decimal("0"),
                )
                period_summary = summary_totals_by_period.get(
                    period,
                    {
                        "total_operation_value": Decimal("0"),
                        "base_icms": Decimal("0"),
                        "icms_value": Decimal("0"),
                    },
                )
                period_difference_sale = (Decimal(period_summary["total_operation_value"]) - period_products_sale).quantize(Decimal("0.01"))
                period_difference_base = (Decimal(period_summary["base_icms"]) - period_products_base).quantize(Decimal("0.01"))
                period_difference_icms = (Decimal(period_summary["icms_value"]) - period_products_icms).quantize(Decimal("0.01"))
                period_difference_rate = (
                    (period_difference_icms * Decimal("100") / period_difference_base).quantize(Decimal("0.01"))
                    if period_difference_base > 0
                    else Decimal("0")
                )
                if (
                    period_difference_sale != Decimal("0.00")
                    or period_difference_base != Decimal("0.00")
                    or period_difference_icms != Decimal("0.00")
                ):
                    has_adjustment = True
                adjustment_values.extend([period_difference_sale, period_difference_base, period_difference_icms, period_difference_rate])
                total_adjustment_sale += period_difference_sale
                total_adjustment_base += period_difference_base
                total_adjustment_icms += period_difference_icms
            total_adjustment_rate = (
                (total_adjustment_icms * Decimal("100") / total_adjustment_base).quantize(Decimal("0.01"))
                if total_adjustment_base > 0
                else Decimal("0")
            )
            adjustment_values.extend([total_adjustment_sale, total_adjustment_base, total_adjustment_icms, total_adjustment_rate])
            has_positive_adjustment = (
                total_adjustment_sale > Decimal("0.00")
                or total_adjustment_base > Decimal("0.00")
                or total_adjustment_icms > Decimal("0.00")
            )
            if has_adjustment and has_positive_adjustment:
                adjustment_row = {"curve_abc": "", "values": adjustment_values}
                display_rows.append(adjustment_row)
                export_rows.append(adjustment_values)

        dialog = Toplevel(self.root)
        dialog.title(f"Levantamento por Produto - {operation_type}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1360, 620, 1100, 500, margin_y=190)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        ttk.Label(container, text=f"Levantamento por Produto/Mes - {operation_type}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text="Resultado linear por produto: cada mes aparece em colunas separadas com operacao, valor, base, ICMS e aliquota calculada.",
            wraplength=1320,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=2)
        product_var = StringVar()
        curve_var = StringVar()
        search_var = StringVar()

        for column_index, (label_text, variable) in enumerate((
            ("Produto", product_var),
            ("Curva ABC", curve_var),
            ("Busca", search_var),
        )):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 2 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))

        month_band = ttk.Frame(container)
        month_band.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        month_band.columnconfigure(0, weight=0)
        if periods:
            ttk.Label(month_band, text="Meses:", style="Strong.TLabel").pack(side=LEFT, padx=(0, 8))
        for period_index, period in enumerate(periods):
            band = tk.Label(
                month_band,
                text=f"{period}  Valor | Base | ICMS | Aliq",
                bg="#d9eaf7" if period_index % 2 == 0 else "#fbe7c6",
                fg="#184c73",
                padx=10,
                pady=4,
                relief="solid",
                bd=1,
            )
            band.pack(side=LEFT, padx=(0, 6))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=4, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = ["curve_abc", "code", "description", "ncm", "cfops", "csts"]
        for period in periods:
            period_key = period.replace("/", "_")
            columns.extend(
                [
                    f"{period_key}_sale",
                    f"{period_key}_base",
                    f"{period_key}_icms",
                    f"{period_key}_rate",
                ]
            )
        columns.extend(["total_sale", "total_base", "total_icms", "total_rate"])
        columns = tuple(columns)
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=16, selectmode="extended")
        headings: dict[str, str] = {
            "curve_abc": "Curva ABC",
            "code": "Codigo",
            "description": "Descricao",
            "ncm": "NCM",
            "cfops": "CFOPs",
            "csts": "CSTs",
            "total_sale": "Total Valor Operacao",
            "total_base": "Total Base ICMS",
            "total_icms": "Total Valor ICMS",
            "total_rate": "Aliq Total",
        }
        widths: dict[str, int] = {
            "curve_abc": 80,
            "code": 95,
            "description": 260,
            "ncm": 90,
            "cfops": 110,
            "csts": 100,
            "total_sale": 130,
            "total_base": 130,
            "total_icms": 130,
            "total_rate": 90,
        }
        for period in periods:
            period_key = period.replace("/", "_")
            headings[f"{period_key}_sale"] = f"{period} Valor"
            headings[f"{period_key}_base"] = f"{period} Base"
            headings[f"{period_key}_icms"] = f"{period} ICMS"
            headings[f"{period_key}_rate"] = f"{period} Aliq"
            widths[f"{period_key}_sale"] = 110
            widths[f"{period_key}_base"] = 110
            widths[f"{period_key}_icms"] = 110
            widths[f"{period_key}_rate"] = 90
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.tag_configure("curve_a", background="#f4fff4")
        tree.tag_configure("curve_b", background="#fff8e1")
        tree.tag_configure("curve_c", background="#ffe8e8")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))

        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        footer = ttk.Frame(container)
        footer.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        total_rows_var = StringVar(value="Linhas: 0")
        total_products_var = StringVar(value="Produtos: 0")
        total_sale_var = StringVar(value="Valor Operacao: 0,00")
        total_base_var = StringVar(value="Base ICMS: 0,00")
        total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        total_icms_var = StringVar(value="Valor ICMS: 0,00")
        ttk.Label(footer, textvariable=total_rows_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_products_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        export_actions = ttk.Frame(container)
        export_actions.grid(row=6, column=0, sticky="w", pady=(10, 0))

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)

            product_filter = normalize_text(product_var.get())
            curve_filter = normalize_text(curve_var.get())
            search_filter = normalize_text(search_var.get())
            filtered_products: set[str] = set()
            filtered_export_rows: list[list[object]] = []

            for row in display_rows:
                values = list(row["values"])
                search_text = normalize_text(" ".join(serialize_value_for_clipboard(value) for value in values))
                code_text = normalize_text(serialize_value_for_clipboard(values[1]))
                if product_filter and product_filter not in code_text:
                    continue
                if curve_filter and curve_filter not in normalize_text(str(row["curve_abc"])):
                    continue
                if search_filter and search_filter not in search_text:
                    continue

                tree.insert(
                    "",
                    END,
                    values=tuple(serialize_value_for_clipboard(value) for value in values),
                    tags=(f"curve_{str(row['curve_abc']).strip().lower()}",),
                )
                filtered_products.add(str(values[1]))
                filtered_export_rows.append(values)

            total_sale = sum(
                (
                    Decimal(source_row.get("sale_value", Decimal("0")))
                    for source_row in filtered_rows
                    if str(source_row.get("code", "")).strip() in filtered_products
                ),
                Decimal("0"),
            )
            total_base = sum(
                (
                    Decimal(source_row.get("base_icms", Decimal("0")))
                    for source_row in filtered_rows
                    if str(source_row.get("code", "")).strip() in filtered_products
                ),
                Decimal("0"),
            )
            total_icms = sum(
                (
                    Decimal(source_row.get("icms_value", Decimal("0")))
                    for source_row in filtered_rows
                    if str(source_row.get("code", "")).strip() in filtered_products
                ),
                Decimal("0"),
            )
            adjustment_total_sale = sum(
                (
                    Decimal(row[-4])
                    for row in filtered_export_rows
                    if str(row[1]).strip() == "(SEM PRODUTO)"
                ),
                Decimal("0"),
            )
            adjustment_total_base = sum(
                (
                    Decimal(row[-3])
                    for row in filtered_export_rows
                    if str(row[1]).strip() == "(SEM PRODUTO)"
                ),
                Decimal("0"),
            )
            adjustment_total_icms = sum(
                (
                    Decimal(row[-2])
                    for row in filtered_export_rows
                    if str(row[1]).strip() == "(SEM PRODUTO)"
                ),
                Decimal("0"),
            )
            total_sale += adjustment_total_sale
            total_base += adjustment_total_base
            total_icms += adjustment_total_icms

            tree._filtered_popup_rows = filtered_export_rows
            total_rows_var.set(f"Linhas: {len(filtered_export_rows)}")
            total_products_var.set(f"Produtos: {len(filtered_products)}")
            total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
            total_base_ratio_var.set(
                f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}"
            )
            total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

        for variable in (product_var, curve_var, search_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()

        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_product_monthly_popup_dataset(
                f"levantamento_produto_{operation_type.lower()}",
                f"Levantamento {operation_type}",
                export_headers,
                list(getattr(tree, "_filtered_popup_rows", [])),
                periods,
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_product_monthly_popup_dataset(
                f"levantamento_produto_{operation_type.lower()}",
                f"Levantamento {operation_type}",
                export_headers,
                list(getattr(tree, "_filtered_popup_rows", [])),
                periods,
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        self.log_message(f"Levantamento linear por produto de {operation_type.lower()} aberto com {len(export_rows)} linha(s).")

    def handle_consult_tree_click(self, event: object) -> None:
        if not hasattr(self, "consult_tree"):
            return
        row_id = self.consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id or column_id not in {"#10", "#11"}:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.consult_filtered_rows):
            return
        self.open_consultation_launch_details(self.consult_filtered_rows[row_index])

    def handle_sales_consult_tree_click(self, event: object) -> None:
        if not hasattr(self, "sales_consult_tree"):
            return
        row_id = self.sales_consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.sales_consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id or column_id not in {"#10", "#11"}:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.sales_consult_filtered_rows):
            return
        self.open_consultation_launch_details(self.sales_consult_filtered_rows[row_index])

    def handle_consult_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "consult_tree"):
            return
        row_id = self.consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.consult_filtered_rows):
            return
        target_row = self.consult_filtered_rows[row_index]
        if column_id == "#5":
            self.open_consultation_supplier_popup(self.build_supplier_popup_row_from_detail_row(target_row))
            return
        self.open_consultation_launch_details(target_row)

    def handle_sales_consult_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "sales_consult_tree"):
            return
        row_id = self.sales_consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.sales_consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.sales_consult_filtered_rows):
            return
        target_row = self.sales_consult_filtered_rows[row_index]
        if column_id == "#5":
            self.open_consultation_supplier_popup(self.build_supplier_popup_row_from_detail_row(target_row))
            return
        self.open_consultation_launch_details(target_row)

    def handle_contrib_consult_tree_click(self, event: object) -> None:
        if not hasattr(self, "contrib_consult_tree"):
            return
        row_id = self.contrib_consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.contrib_consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id or column_id not in {"#12", "#13"}:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.contrib_consult_filtered_rows):
            return
        self.open_contrib_consultation_launch_details(self.contrib_consult_filtered_rows[row_index])

    def handle_contrib_sales_consult_tree_click(self, event: object) -> None:
        if not hasattr(self, "contrib_sales_consult_tree"):
            return
        row_id = self.contrib_sales_consult_tree.identify_row(getattr(event, "y", 0))
        column_id = self.contrib_sales_consult_tree.identify_column(getattr(event, "x", 0))
        if not row_id or column_id not in {"#12", "#13"}:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.contrib_sales_consult_filtered_rows):
            return
        self.open_contrib_consultation_launch_details(self.contrib_sales_consult_filtered_rows[row_index])

    def handle_contrib_consult_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "contrib_consult_tree"):
            return
        row_id = self.contrib_consult_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.contrib_consult_filtered_rows):
            return
        self.open_contrib_consultation_launch_details(self.contrib_consult_filtered_rows[row_index])

    def handle_contrib_sales_consult_tree_double_click(self, event: object) -> None:
        if not hasattr(self, "contrib_sales_consult_tree"):
            return
        row_id = self.contrib_sales_consult_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            return
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(self.contrib_sales_consult_filtered_rows):
            return
        self.open_contrib_consultation_launch_details(self.contrib_sales_consult_filtered_rows[row_index])

    def open_consultation_launch_details(self, row: dict[str, object]) -> None:
        launch_details = list(row.get("launch_details", []))
        if not launch_details:
            messagebox.showwarning("Detalhamento", "Nao ha lancamentos detalhados para esta linha.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Detalhamento - {row['code']} - {row['period']}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1200, 500, 980, 420, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        ttk.Label(
            container,
            text=(
                f"Produto {row['code']} - {row['description']} | Periodo {row['period']} | "
                f"Docs {row['document_count']} | Lancamentos {row['launch_count']}"
            ),
            style="Title.TLabel",
            wraplength=1120,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        documents_summary: dict[str, int] = defaultdict(int)
        for detail in launch_details:
            document_number = str(detail.get("document_number", "")).strip() or "(sem numero)"
            documents_summary[document_number] += 1
        documents_summary_box = ttk.Frame(container)
        documents_summary_box.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(documents_summary_box, text="Documentos encontrados:").pack(anchor="w")

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=2)
        detail_doc_filter_var = StringVar()
        detail_cfop_filter_var = StringVar()
        detail_cst_filter_var = StringVar()
        detail_search_var = StringVar()

        documents_badge_row = ttk.Frame(documents_summary_box)
        documents_badge_row.pack(fill="x", pady=(4, 0))
        ttk.Button(
            documents_badge_row,
            text="Todos",
            style="Secondary.TButton",
            command=lambda: detail_doc_filter_var.set(""),
        ).pack(side=LEFT, padx=(0, 6))
        for document, count in sorted(documents_summary.items(), key=lambda item: item[0]):
            badge_text = f"{document}: {count} lanc."
            badge_button = ttk.Button(
                documents_badge_row,
                text=badge_text,
                style="Secondary.TButton",
                command=lambda current=document: detail_doc_filter_var.set(current),
            )
            badge_button.pack(side=LEFT, padx=(0, 6))

        doc_filter = ttk.Frame(filter_box)
        doc_filter.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(doc_filter, text="Filtro Documento:").pack(anchor="w")
        ttk.Entry(doc_filter, textvariable=detail_doc_filter_var).pack(fill="x", pady=(4, 0))

        cfop_filter = ttk.Frame(filter_box)
        cfop_filter.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Label(cfop_filter, text="Filtro CFOP:").pack(anchor="w")
        ttk.Entry(cfop_filter, textvariable=detail_cfop_filter_var).pack(fill="x", pady=(4, 0))

        cst_filter = ttk.Frame(filter_box)
        cst_filter.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        ttk.Label(cst_filter, text="Filtro CST:").pack(anchor="w")
        ttk.Entry(cst_filter, textvariable=detail_cst_filter_var).pack(fill="x", pady=(4, 0))

        search_filter = ttk.Frame(filter_box)
        search_filter.grid(row=0, column=3, sticky="ew")
        ttk.Label(search_filter, text="Busca:").pack(anchor="w")
        ttk.Entry(search_filter, textvariable=detail_search_var).pack(fill="x", pady=(4, 0))

        detail_box = ttk.Frame(container)
        detail_box.grid(row=3, column=0, sticky="nsew")
        detail_box.columnconfigure(0, weight=1)
        detail_box.rowconfigure(0, weight=1)
        detail_box.rowconfigure(1, weight=0)

        columns = (
            "document_date",
            "document_number",
            "participant_name",
            "document_key",
            "item_number",
            "cest",
            "cfop",
            "cst_icms",
            "icms_rate",
            "effective_icms_rate",
            "quantity",
            "sale_value",
            "ipi_value",
            "base_icms",
            "operation_base_gap",
            "difference_origin",
            "icms_value",
        )
        detail_tree = ttk.Treeview(detail_box, columns=columns, show="headings", height=10)
        headings = {
            "document_date": "Data",
            "document_number": "Documento",
            "participant_name": "Fornecedor",
            "document_key": "Chave",
            "item_number": "Item",
            "cest": "CEST",
            "cfop": "CFOP",
            "cst_icms": "CST",
            "icms_rate": "Aliq ICMS",
            "effective_icms_rate": "Aliq Efetiva",
            "quantity": "Quantidade",
            "sale_value": "Total Operacao",
            "ipi_value": "Valor IPI",
            "base_icms": "Base ICMS",
            "operation_base_gap": "Dif. Oper/Base",
            "difference_origin": "Origem Dif.",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "document_date": 90,
            "document_number": 90,
            "participant_name": 220,
            "document_key": 260,
            "item_number": 60,
            "cest": 90,
            "cfop": 80,
            "cst_icms": 80,
            "icms_rate": 90,
            "effective_icms_rate": 95,
            "quantity": 90,
            "sale_value": 110,
            "ipi_value": 100,
            "base_icms": 110,
            "operation_base_gap": 110,
            "difference_origin": 110,
            "icms_value": 110,
        }
        for column_id in columns:
            detail_tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(detail_tree, columns, headings)
        detail_tree.grid(row=0, column=0, sticky="nsew")
        detail_tree.tag_configure("normal", background="#ffffff")
        detail_tree.tag_configure("divergent", background="#ffe8e8")
        detail_tree.tag_configure("reduced", background="#e8f2ff")
        detail_tree.tag_configure("dynamic_rule", background="#eadcff")
        detail_tree.bind("<Double-1>", lambda event: self.handle_detail_tree_double_click(event, detail_tree, launch_details))

        detail_scroll_y = ttk.Scrollbar(detail_box, orient="vertical", command=detail_tree.yview)
        detail_scroll_y.grid(row=0, column=1, sticky="ns")
        detail_scroll_x = ttk.Scrollbar(detail_box, orient="horizontal", command=detail_tree.xview)
        detail_scroll_x.grid(row=1, column=0, sticky="ew")
        detail_tree.configure(yscrollcommand=detail_scroll_y.set, xscrollcommand=detail_scroll_x.set)

        def scroll_detail_horizontally(event: object) -> str:
            delta = getattr(event, "delta", 0)
            if delta:
                detail_tree.xview_scroll(int(-5 * (delta / 120)), "units")
            return "break"

        detail_tree.bind("<Shift-MouseWheel>", scroll_detail_horizontally)

        cfops_seen = {str(detail.get("cfop", "")).strip() for detail in launch_details if str(detail.get("cfop", "")).strip()}
        csts_seen = {str(detail.get("cst_icms", "")).strip() for detail in launch_details if str(detail.get("cst_icms", "")).strip()}
        bases_seen = {
            Decimal(detail.get("base_icms", Decimal("0"))).quantize(Decimal("0.01"))
            for detail in launch_details
        }
        icms_seen = {
            Decimal(detail.get("icms_value", Decimal("0"))).quantize(Decimal("0.01"))
            for detail in launch_details
        }
        has_launch_divergence = len(cfops_seen) > 1 or len(csts_seen) > 1 or len(bases_seen) > 1 or len(icms_seen) > 1
        info_text = "Lancamentos consistentes no periodo."
        if has_launch_divergence:
            info_text = "Atencao: ha divergencia de CFOP/CST/Base/ICMS entre documentos/lancamentos deste periodo."
        info_label = ttk.Label(container, text=info_text, style="Title.TLabel", wraplength=1120)
        info_label.grid(row=4, column=0, sticky="w", pady=(10, 0))

        export_actions = ttk.Frame(container)
        export_actions.grid(row=5, column=0, sticky="w", pady=(10, 0))
        ttk.Button(
            export_actions,
            text="Exportar Popup Excel",
            style="Secondary.TButton",
            command=lambda: self.export_launch_details(
                list(getattr(detail_tree, "_filtered_launch_details", launch_details)),
                row,
                "xlsx",
            ),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Popup CSV",
            style="Secondary.TButton",
            command=lambda: self.export_launch_details(
                list(getattr(detail_tree, "_filtered_launch_details", launch_details)),
                row,
                "csv",
            ),
        ).pack(side=LEFT, padx=(8, 0))

        footer = ttk.Frame(container)
        footer.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        detail_total_items_var = StringVar(value=f"Lancamentos: {len(launch_details)}")
        detail_total_quantity_var = StringVar(value="Quantidade: 0,00")
        detail_total_sale_var = StringVar(value="Valor Operacao: 0,00")
        detail_total_ipi_var = StringVar(value="Valor IPI: 0,00")
        detail_total_base_var = StringVar(value="Base ICMS: 0,00")
        detail_total_base_ratio_var = StringVar(value="% Base/Oper: 0,00")
        detail_total_icms_var = StringVar(value="Valor ICMS: 0,00")
        ttk.Label(footer, textvariable=detail_total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_quantity_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_ipi_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_base_ratio_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=detail_total_icms_var, style="Strong.TLabel").pack(side=LEFT)
        detail_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(detail_tree))
        detail_tree.bind("<Button-3>", lambda event: self.show_detail_tree_menu(event, detail_tree, launch_details, row))

        def refresh_detail_tree(*_args: object) -> None:
            for item_id in detail_tree.get_children():
                detail_tree.delete(item_id)

            document_filter = normalize_text(detail_doc_filter_var.get())
            cfop_filter_value = normalize_text(detail_cfop_filter_var.get())
            cst_filter_value = normalize_text(detail_cst_filter_var.get())
            search_filter_value = normalize_text(detail_search_var.get())

            total_quantity = Decimal("0")
            total_sale = Decimal("0")
            total_ipi = Decimal("0")
            total_base = Decimal("0")
            total_icms = Decimal("0")
            launch_count = 0
            filtered_launch_details: list[dict[str, object]] = []
            runtime_rules = self.get_current_runtime_rules_for_highlight()

            for detail in launch_details:
                document_number = str(detail.get("document_number", "")).strip()
                cfop_value = str(detail.get("cfop", "")).strip()
                cst_value = str(detail.get("cst_icms", "")).strip()
                search_text = normalize_text(
                    f"{detail.get('document_number', '')} {detail.get('document_key', '')} {detail.get('item_number', '')} "
                    f"{detail.get('cest', '')} {detail.get('cfop', '')} {detail.get('cst_icms', '')} {detail.get('description', '')}"
                )
                if document_filter and document_filter not in normalize_text(document_number):
                    continue
                if cfop_filter_value and cfop_filter_value not in normalize_text(cfop_value):
                    continue
                if cst_filter_value and cst_filter_value not in normalize_text(cst_value):
                    continue
                if search_filter_value and search_filter_value not in search_text:
                    continue

                filtered_launch_details.append(detail)
                quantity = Decimal(detail.get("quantity", Decimal("0")))
                sale_value = get_launch_total_operation_value(detail)
                ipi_value = Decimal(detail.get("ipi_value", Decimal("0")))
                base_icms = Decimal(detail.get("base_icms", Decimal("0")))
                icms_value = Decimal(detail.get("icms_value", Decimal("0")))
                operation_base_gap = get_operation_base_difference(detail)
                difference_origin = describe_operation_base_difference(
                    sale_value,
                    base_icms,
                    ipi_value,
                    Decimal(detail.get("icms_st_value", Decimal("0"))),
                    Decimal(detail.get("discount_value", Decimal("0"))),
                )
                base_key = base_icms.quantize(Decimal("0.01"))
                icms_key = icms_value.quantize(Decimal("0.01"))
                is_divergent_launch = (
                    (len(cfops_seen) > 1 and cfop_value in cfops_seen)
                    or (len(csts_seen) > 1 and cst_value in csts_seen)
                    or (len(bases_seen) > 1 and base_key in bases_seen)
                    or (len(icms_seen) > 1 and icms_key in icms_seen)
                )
                has_reduction = has_icms_reduction(detail.get("icms_rate", Decimal("0")), sale_value, base_icms, icms_value)
                current_row_index = launch_count
                row_tag = "dynamic_rule" if get_first_matching_icms_rule(runtime_rules, detail) is not None else (
                    "divergent" if is_divergent_launch and has_launch_divergence else "reduced" if has_reduction else "normal"
                )
                detail_tree.insert(
                    "",
                    END,
                    iid=str(current_row_index),
                    values=(
                        detail.get("document_date", ""),
                        detail.get("document_number", ""),
                        detail.get("participant_name", ""),
                        detail.get("document_key", ""),
                        detail.get("item_number", ""),
                        detail.get("cest", ""),
                        cfop_value,
                        cst_value,
                        format_decimal_sped(Decimal(detail.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(compute_display_icms_rate(detail.get("icms_rate", Decimal("0")), sale_value, base_icms, icms_value)),
                        format_decimal_sped(quantity),
                        format_decimal_sped(sale_value),
                        format_decimal_sped(ipi_value),
                        format_decimal_sped(base_icms),
                        format_decimal_sped(operation_base_gap),
                        difference_origin,
                        format_decimal_sped(icms_value),
                    ),
                    tags=(row_tag,),
                )
                total_quantity += quantity
                total_sale += sale_value
                total_ipi += ipi_value
                total_base += base_icms
                total_icms += icms_value
                launch_count += 1

            detail_tree._filtered_launch_details = filtered_launch_details
            detail_total_items_var.set(f"Lancamentos: {launch_count}")
            detail_total_quantity_var.set(f"Quantidade: {format_decimal_sped(total_quantity)}")
            detail_total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            detail_total_ipi_var.set(f"Valor IPI: {format_decimal_sped(total_ipi)}")
            detail_total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
            detail_total_base_ratio_var.set(
                f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}"
            )
            detail_total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

        for variable in (detail_doc_filter_var, detail_cfop_filter_var, detail_cst_filter_var, detail_search_var):
            variable.trace_add("write", refresh_detail_tree)
        refresh_detail_tree()

    def show_detail_tree_menu(
        self,
        event: object,
        detail_tree: ttk.Treeview,
        launch_details: list[dict[str, object]],
        parent_row: dict[str, object],
    ) -> None:
        row_id = detail_tree.identify_row(getattr(event, "y", 0))
        if not row_id:
            self.show_treeview_copy_menu(event, detail_tree)
            return
        detail_tree.selection_set(row_id)
        filtered_launch_details = getattr(detail_tree, "_filtered_launch_details", launch_details)
        try:
            row_index = int(row_id)
        except ValueError:
            self.show_treeview_copy_menu(event, detail_tree)
            return
        if row_index < 0 or row_index >= len(filtered_launch_details):
            self.show_treeview_copy_menu(event, detail_tree)
            return

        detail = filtered_launch_details[row_index]
        menu = Menu(detail_tree, tearoff=0)
        has_matching_rule = self.add_matching_runtime_rule_menu_items(menu, detail)
        if not has_matching_rule:
            menu.add_command(
                label="Criar Regra Dinamica",
                command=lambda current_detail=detail, current_row_id=row_id: self.open_runtime_rule_builder_for_launch_detail(
                    current_detail,
                    parent_row,
                    detail_tree,
                    current_row_id,
                ),
            )
        menu.add_separator()
        menu.add_command(label="Copiar Selecao", command=lambda: self.copy_treeview_selection(detail_tree))
        menu.tk_popup(getattr(event, "x_root", 0), getattr(event, "y_root", 0))

    def open_runtime_rule_builder_for_launch_detail(
        self,
        detail: dict[str, object],
        parent_row: dict[str, object],
        source_tree: ttk.Treeview | None = None,
        source_row_id: str | None = None,
    ) -> None:
        operation_value = normalize_operation_type(str(detail.get("operation_type", "")).strip())
        if not operation_value:
            parent_period = normalize_text(str(parent_row.get("period", "")))
            if "entrada" in parent_period:
                operation_value = "Entrada"
            elif "saida" in parent_period:
                operation_value = "Saida"

        document_key = normalize_document_key(str(detail.get("document_key", "")))
        tax_id_value = extract_tax_id_from_document_key(document_key)
        if not tax_id_value:
            fallback_tax_id = normalize_document_key(
                str(detail.get("document_tax_id", "") or detail.get("participant_tax_id", ""))
            )
            tax_id_value = fallback_tax_id if len(fallback_tax_id) in {11, 14} else ""

        rate_value = detail.get("icms_rate", Decimal("0"))
        try:
            rate_text = format_decimal_sped(Decimal(rate_value))
        except Exception:
            rate_text = str(rate_value or "").strip()

        def mark_source_row(_rule_line: str) -> None:
            if source_tree is not None and source_row_id:
                try:
                    source_tree.item(source_row_id, tags=("dynamic_rule",))
                except Exception:
                    pass
            if operation_value == "Entrada":
                self.refresh_consultation_tree()
            elif operation_value == "Saida":
                self.refresh_sales_consultation_tree()

        self.open_runtime_rule_builder(
            {
                "operation": operation_value.lower() or "entrada",
                "document": str(detail.get("document_number", "")).strip(),
                "tax_id": tax_id_value,
                "cst": str(detail.get("cst_icms", "")).strip().split("|", 1)[0].strip(),
                "cfop": str(detail.get("cfop", "")).strip().split("|", 1)[0].strip(),
                "rate": rate_text,
                "codes": str(detail.get("code", "")).strip(),
            },
            on_rule_added=mark_source_row,
        )

    def handle_detail_tree_double_click(
        self,
        event: object,
        detail_tree: ttk.Treeview,
        launch_details: list[dict[str, object]],
    ) -> None:
        row_id = detail_tree.identify_row(getattr(event, "y", 0))
        column_id = detail_tree.identify_column(getattr(event, "x", 0))
        if not row_id or column_id != "#4":
            return
        filtered_launch_details = getattr(detail_tree, "_filtered_launch_details", launch_details)
        try:
            row_index = int(row_id)
        except ValueError:
            return
        if row_index < 0 or row_index >= len(filtered_launch_details):
            return
        detail = filtered_launch_details[row_index]
        note_snapshot = detail.get("note_snapshot")
        if not isinstance(note_snapshot, dict):
            messagebox.showwarning("Nota completa", "Nao foi possivel localizar os dados completos da nota.")
            return
        self.open_complete_note_popup(note_snapshot)

    def apply_current_icms_rules_to_note_snapshot(
        self,
        note_snapshot: dict[str, object],
        operation_type: str = "",
    ) -> dict[str, object]:
        runtime_rules = self.get_current_runtime_rules_for_highlight()
        selected_rule_profile = self.adjustment_mode_var.get().strip() if hasattr(self, "adjustment_mode_var") else ""
        if selected_rule_profile == "Filtros":
            selected_rule_profile = ""
        if not runtime_rules and not selected_rule_profile:
            return note_snapshot

        adjusted_snapshot = dict(note_snapshot)
        adjusted_items: list[dict[str, object]] = []
        normalized_operation = normalize_operation_type(operation_type) if str(operation_type).strip() else ""
        for item in list(note_snapshot.get("items", [])):
            item_data = dict(item)
            if normalized_operation and not str(item_data.get("operation_type", "")).strip():
                item_data["operation_type"] = normalized_operation
            adjusted_item = apply_sped_icms_consistency_rules(
                item_data,
                selected_rule_profile,
                runtime_rules,
            )
            adjusted_items.append(adjusted_item)
        adjusted_snapshot["items"] = adjusted_items
        adjusted_snapshot["_runtime_rules_applied"] = bool(runtime_rules or selected_rule_profile)
        return adjusted_snapshot

    def collect_note_snapshots_from_details(
        self,
        detailed_rows: list[dict[str, object]],
        operation_type: str = "",
    ) -> list[dict[str, object]]:
        normalized_operation = normalize_operation_type(operation_type) if str(operation_type).strip() else ""
        documents_by_identity: dict[tuple[str, str, str], dict[str, object]] = {}

        for item in detailed_rows:
            item_operation = normalize_operation_type(str(item.get("operation_type", "")))
            if normalized_operation and item_operation and item_operation != normalized_operation:
                continue
            note_snapshot = item.get("note_snapshot")
            if not isinstance(note_snapshot, dict):
                continue
            note_items = list(note_snapshot.get("items", []))
            if not note_items:
                continue
            document_key = normalize_document_key(str(note_snapshot.get("document_key", "")))
            document_number = str(note_snapshot.get("document_number", "")).strip()
            document_date = str(note_snapshot.get("document_date", "")).strip()
            period = str(note_snapshot.get("period", item.get("period", ""))).strip()
            identity = (document_key or document_number, document_date, period)
            if not identity[0]:
                continue
            bucket = documents_by_identity.setdefault(
                identity,
                {
                    "document_number": document_number,
                    "document_key": document_key,
                    "document_date": document_date,
                    "document_series": str(note_snapshot.get("document_series", "")).strip(),
                    "document_model": str(note_snapshot.get("document_model", "")).strip(),
                    "participant_code": str(note_snapshot.get("participant_code", "")).strip(),
                    "participant_name": str(note_snapshot.get("participant_name", "")).strip(),
                    "participant_tax_id": str(note_snapshot.get("participant_tax_id", "")).strip(),
                    "period": period,
                    "items": [],
                },
            )
            if not bucket["items"]:
                snapshot_with_operation = dict(note_snapshot)
                snapshot_with_operation["items"] = [
                    {
                        **dict(current_item),
                        "operation_type": str(current_item.get("operation_type", "")).strip() or item_operation,
                    }
                    for current_item in note_items
                ]
                adjusted_snapshot = self.apply_current_icms_rules_to_note_snapshot(
                    snapshot_with_operation,
                    item_operation,
                )
                bucket["items"] = [dict(current_item) for current_item in adjusted_snapshot.get("items", [])]
                if adjusted_snapshot.get("_runtime_rules_applied"):
                    bucket["_runtime_rules_applied"] = True

        snapshots = list(documents_by_identity.values())
        snapshots.sort(
            key=lambda snapshot: (
                period_label_sort_key(str(snapshot.get("period", ""))),
                parse_sped_document_date(str(snapshot.get("document_date", ""))) or dt.date.min,
                normalize_text(str(snapshot.get("document_number", ""))),
                normalize_text(str(snapshot.get("document_key", ""))),
            ),
            reverse=True,
        )
        return snapshots

    def open_fiscal_documents_popup(self, source_view: str, operation_type: str) -> None:
        try:
            context = self.load_current_popup_context(source_view)
        except Exception as exc:
            self.log_message(f"Falha ao abrir espelho de documentos: {exc}")
            messagebox.showerror("Espelho de documentos", str(exc))
            return

        note_snapshots = self.collect_note_snapshots_from_details(list(context.get("filtered_details", [])), operation_type)
        if not note_snapshots:
            messagebox.showwarning("Espelho de documentos", "Nao ha documentos com espelho disponivel para os filtros atuais.")
            return

        normalized_operation = normalize_operation_type(operation_type)
        operation_caption = "Entradas" if normalized_operation == "entrada" else "Saidas"

        dialog = Toplevel(self.root)
        dialog.title(f"Espelho de Documentos Fiscais - {operation_caption}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1380, 760, 1120, 620, margin_y=180)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        ttk.Label(
            container,
            text=f"Espelho de Documentos Fiscais - {operation_caption}",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text="Pesquise pelo numero do documento ou pela chave da NF-e/NFC-e. Duplo clique abre o espelho completo.",
            wraplength=1320,
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        filter_box = ttk.Frame(container)
        filter_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_box.columnconfigure(0, weight=2)
        filter_box.columnconfigure(1, weight=2)
        filter_box.columnconfigure(2, weight=1)
        filter_box.columnconfigure(3, weight=2)

        document_filter_var = StringVar()
        key_filter_var = StringVar()
        model_filter_var = StringVar()
        search_filter_var = StringVar()

        for column_index, (label_text, variable) in enumerate(
            (
                ("Numero Documento", document_filter_var),
                ("Chave", key_filter_var),
                ("Modelo", model_filter_var),
                ("Busca", search_filter_var),
            )
        ):
            box = ttk.Frame(filter_box)
            box.grid(row=0, column=column_index, sticky="ew", padx=(0, 8 if column_index < 3 else 0))
            ttk.Label(box, text=f"{label_text}:").pack(anchor="w")
            ttk.Entry(box, textvariable=variable).pack(fill="x", pady=(4, 0))

        tree_box = ttk.Frame(container)
        tree_box.grid(row=3, column=0, sticky="nsew")
        tree_box.columnconfigure(0, weight=1)
        tree_box.rowconfigure(0, weight=1)

        columns = (
            "period",
            "document_date",
            "document_number",
            "document_series",
            "document_model",
            "participant_name",
            "participant_tax_id",
            "document_key",
            "items",
            "sale_value",
            "base_icms",
            "icms_value",
        )
        tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=18, selectmode="extended")
        headings = {
            "period": "Periodo",
            "document_date": "Data",
            "document_number": "Documento",
            "document_series": "Serie",
            "document_model": "Modelo",
            "participant_name": "Fornecedor/Participante",
            "participant_tax_id": "CPF/CNPJ",
            "document_key": "Chave",
            "items": "Itens",
            "sale_value": "Valor Operacao",
            "base_icms": "Base ICMS",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "period": 90,
            "document_date": 90,
            "document_number": 100,
            "document_series": 70,
            "document_model": 70,
            "participant_name": 240,
            "participant_tax_id": 150,
            "document_key": 300,
            "items": 60,
            "sale_value": 120,
            "base_icms": 120,
            "icms_value": 120,
        }
        for column_id in columns:
            tree.column(column_id, width=widths[column_id], anchor="center")
        self.enable_treeview_sorting(tree, columns, headings)
        tree.grid(row=0, column=0, sticky="nsew")
        tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(tree))
        tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, tree))

        scroll_y = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        total_documents_var = StringVar(value=f"Documentos: {len(note_snapshots)}")
        total_items_var = StringVar(value="Itens: 0")
        total_sale_var = StringVar(value="Valor Operacao: 0,00")
        total_base_var = StringVar(value="Base ICMS: 0,00")
        total_icms_var = StringVar(value="Valor ICMS: 0,00")
        ttk.Label(footer, textvariable=total_documents_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_items_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_sale_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_base_var, style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, textvariable=total_icms_var, style="Strong.TLabel").pack(side=LEFT)

        def refresh_tree(*_args: object) -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)

            document_filter = normalize_text(document_filter_var.get())
            key_filter = normalize_text(key_filter_var.get())
            model_filter = normalize_text(model_filter_var.get())
            search_filter = normalize_text(search_filter_var.get())

            filtered_snapshots: list[dict[str, object]] = []
            total_items = 0
            total_sale = Decimal("0")
            total_base = Decimal("0")
            total_icms = Decimal("0")

            for snapshot in note_snapshots:
                document_number = str(snapshot.get("document_number", "")).strip()
                document_key = str(snapshot.get("document_key", "")).strip()
                document_model = str(snapshot.get("document_model", "")).strip()
                participant_name = str(snapshot.get("participant_name", "")).strip()
                participant_tax_id = str(snapshot.get("participant_tax_id", "")).strip()
                search_text = normalize_text(
                    f"{document_number} {document_key} {participant_name} {participant_tax_id} {document_model}"
                )
                if document_filter and document_filter not in normalize_text(document_number):
                    continue
                if key_filter and key_filter not in normalize_text(document_key):
                    continue
                if model_filter and model_filter not in normalize_text(document_model):
                    continue
                if search_filter and search_filter not in search_text:
                    continue

                snapshot_items = list(snapshot.get("items", []))
                sale_value = sum((get_launch_total_operation_value(item) for item in snapshot_items), Decimal("0"))
                base_icms = sum((Decimal(item.get("base_icms", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                icms_value = sum((Decimal(item.get("icms_value", Decimal("0"))) for item in snapshot_items), Decimal("0"))
                row_index = len(filtered_snapshots)
                tree.insert(
                    "",
                    END,
                    iid=str(row_index),
                    values=(
                        snapshot.get("period", ""),
                        snapshot.get("document_date", ""),
                        document_number,
                        snapshot.get("document_series", ""),
                        document_model,
                        participant_name,
                        participant_tax_id,
                        document_key,
                        len(snapshot_items),
                        format_decimal_sped(sale_value),
                        format_decimal_sped(base_icms),
                        format_decimal_sped(icms_value),
                    ),
                )
                filtered_snapshots.append(snapshot)
                total_items += len(snapshot_items)
                total_sale += sale_value
                total_base += base_icms
                total_icms += icms_value

            tree._filtered_note_snapshots = filtered_snapshots
            total_documents_var.set(f"Documentos: {len(filtered_snapshots)}")
            total_items_var.set(f"Itens: {total_items}")
            total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
            total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
            total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

        def handle_tree_double_click(event: object) -> None:
            row_id = tree.identify_row(getattr(event, "y", 0))
            if not row_id:
                return
            filtered_snapshots = getattr(tree, "_filtered_note_snapshots", note_snapshots)
            try:
                row_index = int(row_id)
            except ValueError:
                return
            if row_index < 0 or row_index >= len(filtered_snapshots):
                return
            self.open_complete_note_popup(
                self.apply_current_icms_rules_to_note_snapshot(filtered_snapshots[row_index], operation_type)
            )

        tree.bind("<Double-1>", handle_tree_double_click)
        for variable in (document_filter_var, key_filter_var, model_filter_var, search_filter_var):
            variable.trace_add("write", refresh_tree)
        refresh_tree()

    def open_complete_note_popup(self, note_snapshot: dict[str, object]) -> None:
        note_items = list(note_snapshot.get("items", []))
        if not note_items:
            messagebox.showwarning("Nota completa", "Nao ha itens disponiveis para esta nota.")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Espelho da Nota - {note_snapshot.get('document_number', '')}")
        dialog.transient(self.root)
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1320, 760, 1080, 620, margin_y=190)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        header_box = ttk.LabelFrame(container, text="Cabecalho da Nota", padding=10)
        header_box.grid(row=0, column=0, sticky="ew")
        header_box.columnconfigure(0, weight=1)
        header_box.columnconfigure(1, weight=1)
        header_left = ttk.Frame(header_box)
        header_left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        header_right = ttk.Frame(header_box)
        header_right.grid(row=0, column=1, sticky="nw")

        header_rows_left = [
            ("Documento", note_snapshot.get("document_number", "")),
            ("Chave", note_snapshot.get("document_key", "")),
            ("Data", note_snapshot.get("document_date", "")),
            ("Serie", note_snapshot.get("document_series", "")),
            ("Modelo", note_snapshot.get("document_model", "")),
        ]
        header_rows_right = [
            ("Fornecedor", note_snapshot.get("participant_name", "")),
            ("Cod. Participante", note_snapshot.get("participant_code", "")),
            ("CPF/CNPJ", note_snapshot.get("participant_tax_id", "")),
            ("Aliq ICMS Efetiva", format_decimal_sped(compute_display_icms_rate(Decimal("0"), sum((get_launch_total_operation_value(item) for item in note_items), Decimal("0")), sum((Decimal(item.get("base_icms", Decimal("0"))) for item in note_items), Decimal("0")), sum((Decimal(item.get("icms_value", Decimal("0"))) for item in note_items), Decimal("0"))))),
            ("Itens", len(note_items)),
        ]
        if note_snapshot.get("_runtime_rules_applied"):
            header_rows_right.append(("Regras dinamicas", "aplicadas"))
        for label_text, value_text in header_rows_left:
            ttk.Label(header_left, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)
        for label_text, value_text in header_rows_right:
            ttk.Label(header_right, text=f"{label_text}: {value_text}").pack(anchor="w", pady=2)

        ttk.Label(
            container,
            text="Duplo clique na chave no popup anterior para abrir este espelho completo da nota.",
        ).grid(row=1, column=0, sticky="w", pady=(10, 8))

        export_actions = ttk.Frame(container)
        export_actions.grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Button(
            export_actions,
            text="Exportar Nota Excel",
            style="Secondary.TButton",
            command=lambda: self.export_complete_note(note_snapshot, "xlsx"),
        ).pack(side=LEFT)
        ttk.Button(
            export_actions,
            text="Exportar Nota CSV",
            style="Secondary.TButton",
            command=lambda: self.export_complete_note(note_snapshot, "csv"),
        ).pack(side=LEFT, padx=(8, 0))

        items_box = ttk.Frame(container)
        items_box.grid(row=3, column=0, sticky="nsew")
        items_box.columnconfigure(0, weight=1)
        items_box.rowconfigure(0, weight=1)
        items_box.rowconfigure(1, weight=0)

        columns = (
            "item_number",
            "code",
            "description",
            "ncm",
            "cest",
            "cfop",
            "cst_icms",
            "icms_rate",
            "effective_icms_rate",
            "quantity",
            "sale_value",
            "ipi_value",
            "base_icms",
            "operation_base_gap",
            "difference_origin",
            "icms_value",
        )
        items_tree = ttk.Treeview(items_box, columns=columns, show="headings", height=10, selectmode="extended")
        headings = {
            "item_number": "Item",
            "code": "Codigo",
            "description": "Descricao",
            "ncm": "NCM",
            "cest": "CEST",
            "cfop": "CFOP",
            "cst_icms": "CST",
            "icms_rate": "Aliq ICMS",
            "effective_icms_rate": "Aliq Efetiva",
            "quantity": "Quantidade",
            "sale_value": "Total Operacao",
            "ipi_value": "Valor IPI",
            "base_icms": "Base ICMS",
            "operation_base_gap": "Dif. Oper/Base",
            "difference_origin": "Origem Dif.",
            "icms_value": "Valor ICMS",
        }
        widths = {
            "item_number": 60,
            "code": 90,
            "description": 340,
            "ncm": 100,
            "cest": 90,
            "cfop": 80,
            "cst_icms": 80,
            "icms_rate": 90,
            "effective_icms_rate": 95,
            "quantity": 90,
            "sale_value": 110,
            "ipi_value": 100,
            "base_icms": 110,
            "operation_base_gap": 110,
            "difference_origin": 110,
            "icms_value": 110,
        }
        for column_id in columns:
            items_tree.column(column_id, width=widths[column_id], anchor="center", stretch=False)
        self.enable_treeview_sorting(items_tree, columns, headings)
        items_tree.grid(row=0, column=0, sticky="nsew")
        items_tree.tag_configure("normal", background="#ffffff")
        items_tree.tag_configure("reduced", background="#e8f2ff")
        items_tree.bind("<Control-c>", lambda event: self.copy_treeview_selection(items_tree))
        items_tree.bind("<Button-3>", lambda event: self.show_treeview_copy_menu(event, items_tree))

        items_scroll_y = ttk.Scrollbar(items_box, orient="vertical", command=items_tree.yview)
        items_scroll_y.grid(row=0, column=1, sticky="ns")
        items_scroll_x = ttk.Scrollbar(items_box, orient="horizontal", command=items_tree.xview)
        items_scroll_x.grid(row=1, column=0, sticky="ew")
        items_tree.configure(yscrollcommand=items_scroll_y.set, xscrollcommand=items_scroll_x.set)

        def scroll_items_horizontally(event: object) -> str:
            delta = getattr(event, "delta", 0)
            if delta:
                items_tree.xview_scroll(int(-5 * (delta / 120)), "units")
            return "break"

        items_tree.bind("<Shift-MouseWheel>", scroll_items_horizontally)

        total_quantity = Decimal("0")
        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        for item in note_items:
            quantity = Decimal(item.get("quantity", Decimal("0")))
            sale_value = get_launch_total_operation_value(item)
            base_icms = Decimal(item.get("base_icms", Decimal("0")))
            icms_value = Decimal(item.get("icms_value", Decimal("0")))
            operation_base_gap = get_operation_base_difference(item)
            difference_origin = describe_operation_base_difference(
                sale_value,
                base_icms,
                Decimal(item.get("ipi_value", Decimal("0"))),
                Decimal(item.get("icms_st_value", Decimal("0"))),
                Decimal(item.get("discount_value", Decimal("0"))),
            )
            items_tree.insert(
                "",
                END,
                values=(
                    item.get("item_number", ""),
                    item.get("code", ""),
                    item.get("description", ""),
                    item.get("ncm", ""),
                    item.get("cest", ""),
                    item.get("cfop", ""),
                    item.get("cst_icms", ""),
                    format_decimal_sped(Decimal(item.get("icms_rate", Decimal("0")))),
                    format_decimal_sped(compute_display_icms_rate(item.get("icms_rate", Decimal("0")), sale_value, base_icms, icms_value)),
                    format_decimal_sped(quantity),
                    format_decimal_sped(sale_value),
                    format_decimal_sped(Decimal(item.get("ipi_value", Decimal("0")))),
                    format_decimal_sped(base_icms),
                    format_decimal_sped(operation_base_gap),
                    difference_origin,
                    format_decimal_sped(icms_value),
                ),
                tags=("reduced" if has_icms_reduction(item.get("icms_rate", Decimal("0")), sale_value, base_icms, icms_value) else "normal",),
            )
            total_quantity += quantity
            total_sale += sale_value
            total_base += base_icms
            total_icms += icms_value

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(footer, text=f"Quantidade: {format_decimal_sped(total_quantity)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor Operacao: {format_decimal_sped(total_sale)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Base ICMS: {format_decimal_sped(total_base)}", style="Strong.TLabel").pack(side=LEFT, padx=(0, 18))
        ttk.Label(
            footer,
            text=f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}",
            style="Strong.TLabel",
        ).pack(side=LEFT, padx=(0, 18))
        ttk.Label(footer, text=f"Valor ICMS: {format_decimal_sped(total_icms)}", style="Strong.TLabel").pack(side=LEFT)

    def export_complete_note(self, note_snapshot: dict[str, object], output_type: str) -> None:
        note_items = list(note_snapshot.get("items", []))
        if not note_items:
            messagebox.showwarning("Exportar nota", "Nao ha itens disponiveis para exportar.")
            return

        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Salvar nota completa",
            defaultextension=suffix,
            initialfile=f"nota_{note_snapshot.get('document_number', 'sem_numero')}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        header_headers = ["Campo", "Valor"]
        header_rows = [
            ["Documento", note_snapshot.get("document_number", "")],
            ["Chave", note_snapshot.get("document_key", "")],
            ["Data", note_snapshot.get("document_date", "")],
            ["Serie", note_snapshot.get("document_series", "")],
            ["Modelo", note_snapshot.get("document_model", "")],
            ["Fornecedor", note_snapshot.get("participant_name", "")],
            ["Cod. Participante", note_snapshot.get("participant_code", "")],
            ["CPF/CNPJ", note_snapshot.get("participant_tax_id", "")],
            [
                "Aliq ICMS Efetiva",
                compute_display_icms_rate(
                    Decimal("0"),
                    sum((get_launch_total_operation_value(item) for item in note_items), Decimal("0")),
                    sum((Decimal(item.get("base_icms", Decimal("0"))) for item in note_items), Decimal("0")),
                    sum((Decimal(item.get("icms_value", Decimal("0"))) for item in note_items), Decimal("0")),
                ),
            ],
            ["Quantidade de Itens", len(note_items)],
        ]
        item_headers = [
            "Item",
            "Codigo",
            "Descricao",
            "NCM",
            "CEST",
            "CFOP",
            "CST",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Quantidade",
            "Total Operacao",
            "Valor IPI",
            "Base ICMS",
            "Dif. Oper/Base",
            "Origem Dif.",
            "Valor ICMS",
        ]
        item_rows = [
            [
                item.get("item_number", ""),
                item.get("code", ""),
                item.get("description", ""),
                item.get("ncm", ""),
                item.get("cest", ""),
                item.get("cfop", ""),
                item.get("cst_icms", ""),
                Decimal(item.get("icms_rate", Decimal("0"))),
                compute_display_icms_rate(
                    item.get("icms_rate", Decimal("0")),
                    get_launch_total_operation_value(item),
                    Decimal(item.get("base_icms", Decimal("0"))),
                    Decimal(item.get("icms_value", Decimal("0"))),
                ),
                Decimal(item.get("quantity", Decimal("0"))),
                get_launch_total_operation_value(item),
                Decimal(item.get("ipi_value", Decimal("0"))),
                Decimal(item.get("base_icms", Decimal("0"))),
                get_operation_base_difference(item),
                describe_operation_base_difference(
                    get_launch_total_operation_value(item),
                    Decimal(item.get("base_icms", Decimal("0"))),
                    Decimal(item.get("ipi_value", Decimal("0"))),
                    Decimal(item.get("icms_st_value", Decimal("0"))),
                    Decimal(item.get("discount_value", Decimal("0"))),
                ),
                Decimal(item.get("icms_value", Decimal("0"))),
            ]
            for item in note_items
        ]
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, item_headers, item_rows)
                header_output = output_path.with_name(f"{output_path.stem}_cabecalho.csv")
                write_simple_csv_file(header_output, header_headers, header_rows)
                self.log_message(f"Nota completa exportada em CSV: {output_path} e {header_output}")
                messagebox.showinfo("Exportar nota", f"Arquivos gerados com sucesso:\n{output_path}\n{header_output}")
                return

            write_simple_excel_workbook(
                output_path,
                [
                    ("Cabecalho", header_headers, header_rows),
                    ("Itens", item_headers, item_rows),
                ],
            )
            self.log_message(f"Nota completa exportada com sucesso em: {output_path}")
            messagebox.showinfo("Exportar nota", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar nota completa: {exc}")
            messagebox.showerror("Exportar nota", str(exc))

    def get_current_runtime_rules_for_highlight(self) -> list[dict[str, object]]:
        if not hasattr(self, "runtime_rules_text"):
            return []
        try:
            return parse_runtime_rule_lines(self.runtime_rules_text.get("1.0", END))
        except Exception:
            return []

    def iter_runtime_rule_text_lines(self) -> list[str]:
        if not hasattr(self, "runtime_rules_text"):
            return []
        return [
            line.strip()
            for line in self.runtime_rules_text.get("1.0", END).splitlines()
            if line.strip()
        ]

    def find_matching_runtime_rule_line(self, row: dict[str, object]) -> str:
        candidate_items = list(row.get("launch_details", []))
        if not candidate_items:
            candidate_items = [row]
        for rule_line in self.iter_runtime_rule_text_lines():
            try:
                parsed_rules = parse_runtime_rule_lines(rule_line)
            except Exception:
                continue
            for item in candidate_items:
                if isinstance(item, dict) and get_first_matching_icms_rule(parsed_rules, item) is not None:
                    return rule_line
        return ""

    def add_matching_runtime_rule_menu_items(self, menu: Menu, row: dict[str, object]) -> bool:
        rule_line = self.find_matching_runtime_rule_line(row)
        if not rule_line:
            return False
        menu.add_command(
            label="Visualizar Regra Dinamica",
            command=lambda current_rule=rule_line: self.view_runtime_rule_line(current_rule),
        )
        menu.add_command(
            label="Editar Regra Dinamica",
            command=lambda current_rule=rule_line: self.edit_runtime_rule_line(current_rule),
        )
        menu.add_command(
            label="Remover Regra Dinamica",
            command=lambda current_rule=rule_line: self.remove_runtime_rule_line(current_rule),
        )
        return True

    def consultation_row_matches_runtime_rules(
        self,
        row: dict[str, object],
        runtime_rules: list[dict[str, object]] | None = None,
    ) -> bool:
        current_rules = runtime_rules if runtime_rules is not None else self.get_current_runtime_rules_for_highlight()
        if not current_rules:
            return False
        launch_details = list(row.get("launch_details", []))
        if not launch_details:
            launch_details = [row]
        return any(
            isinstance(detail, dict) and get_first_matching_icms_rule(current_rules, detail) is not None
            for detail in launch_details
        )

    def get_consult_tree_tags(
        self,
        row: dict[str, object],
        runtime_rules: list[dict[str, object]] | None = None,
    ) -> tuple[str, ...]:
        tags: list[str] = []
        normalized_status = str(row.get("status", "") or "").strip()
        if bool(row.get("has_icms_reduction")) and normalized_status == "Ok":
            tags.append("reduced")
        elif normalized_status == "Ok":
            tags.append("ok")
        elif "Sem credito" in normalized_status or "Sem debito" in normalized_status:
            tags.append("warning")
        elif normalized_status == "Sem entrada" or normalized_status == "Sem saida":
            tags.append("warning")
        else:
            tags.append("divergent")
        if int(row.get("supplier_count", 0)) > 1:
            tags.append("multi_supplier")
        if self.consultation_row_matches_runtime_rules(row, runtime_rules):
            tags.append("dynamic_rule")
        return tuple(tags)

    def format_supplier_highlight(self, suppliers: object, supplier_count: object) -> str:
        supplier_text = str(suppliers or "").strip()
        try:
            count = int(supplier_count or 0)
        except (TypeError, ValueError):
            count = 0
        if count > 1:
            if supplier_text:
                return f"[MULTIPLOS: {count}] {supplier_text}"
            return f"[MULTIPLOS: {count}]"
        return supplier_text

    def refresh_consultation_tree(self) -> None:
        if not hasattr(self, "consult_tree"):
            return
        if not getattr(self, "_consultation_filters_ready", False):
            return

        for item_id in self.consult_tree.get_children():
            self.consult_tree.delete(item_id)
        if hasattr(self, "consult_summary_tree"):
            for item_id in self.consult_summary_tree.get_children():
                self.consult_summary_tree.delete(item_id)

        filtered_rows = self.get_consultation_filtered_rows()
        cst_filters = parse_filter_values(self.consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.consult_cfop_filter_var.get())
        effective_rows: list[dict[str, object]] = []
        for row in filtered_rows:
            effective_row = self.build_effective_consultation_row(row, cst_filters, cfop_filters)
            if effective_row is not None:
                effective_rows.append(effective_row)
        filtered_rows = effective_rows
        summary_rows = self.build_consultation_summary_rows(filtered_rows)

        filtered_rows.sort(
            key=lambda row: (
                self.get_consultation_sort_value(row, self.consult_tree_sort_column),
                normalize_text(str(row.get("period", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.consult_tree_sort_reverse,
        )
        summary_rows.sort(
            key=lambda row: (
                self.get_consultation_sort_value(row, self.consult_summary_sort_column),
                normalize_text(str(row.get("description", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.consult_summary_sort_reverse,
        )
        self.consult_filtered_rows = list(filtered_rows)
        self.consult_summary_rows = list(summary_rows)

        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        row_count = 0
        runtime_rules = self.get_current_runtime_rules_for_highlight()

        for row_index, row in enumerate(filtered_rows):
            values = (
                row["period"],
                row.get("curve_abc", "C"),
                row["code"],
                row["description"],
                self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                row["ncm"],
                row.get("cest", ""),
                row["cfop"],
                row["cst_icms"],
                format_decimal_sped(extract_representative_icms_rate(row)),
                format_decimal_sped(extract_effective_icms_rate(row)),
                row["document_count"],
                row["launch_count"],
                format_decimal_sped(Decimal(row["sale_value"])),
                format_decimal_sped(Decimal(row["base_icms"])),
                format_decimal_sped(Decimal(row["icms_value"])),
                row["status"],
            )
            self.consult_tree.insert("", END, iid=str(row_index), values=values, tags=self.get_consult_tree_tags(row, runtime_rules))
            total_sale += Decimal(row["sale_value"])
            total_base += Decimal(row["base_icms"])
            total_icms += Decimal(row["icms_value"])
            row_count += 1

        if hasattr(self, "consult_summary_tree"):
            for row in summary_rows:
                self.consult_summary_tree.insert(
                    "",
                    END,
                    values=(
                        row["code"],
                        row["description"],
                        row.get("curve_abc", "C"),
                        self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                        row["periods"],
                        row["cfops"],
                        row["csts"],
                        format_decimal_sped(Decimal(row.get("nominal_icms_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row["sale_value"])),
                        format_decimal_sped(Decimal(row["base_icms"])),
                        format_decimal_sped(Decimal(row["icms_value"])),
                    ),
                    tags=(
                        ("reduced" if bool(row.get("has_icms_reduction")) else "normal"),
                        *(() if int(row.get("supplier_count", 0)) <= 1 else ("multi_supplier",)),
                    ),
                )

        self.consult_total_items_var.set(f"Linhas: {row_count}")
        self.consult_total_products_var.set(f"Produtos: {len(summary_rows)}")
        self.consult_total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
        self.consult_total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
        self.consult_total_base_ratio_var.set(
            f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}"
        )
        self.consult_total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

    def refresh_sales_consultation_tree(self) -> None:
        if not hasattr(self, "sales_consult_tree"):
            return
        if not getattr(self, "_sales_consultation_filters_ready", False):
            return

        for item_id in self.sales_consult_tree.get_children():
            self.sales_consult_tree.delete(item_id)
        if hasattr(self, "sales_consult_summary_tree"):
            for item_id in self.sales_consult_summary_tree.get_children():
                self.sales_consult_summary_tree.delete(item_id)

        filtered_rows = self.get_sales_consultation_filtered_rows()
        cst_filters = parse_filter_values(self.sales_consult_cst_filter_var.get())
        cfop_filters = parse_filter_values(self.sales_consult_cfop_filter_var.get())
        effective_rows: list[dict[str, object]] = []
        for row in filtered_rows:
            effective_row = self.build_effective_sales_consultation_row(row, cst_filters, cfop_filters)
            if effective_row is not None:
                effective_rows.append(effective_row)
        filtered_rows = effective_rows
        summary_rows = self.build_consultation_summary_rows(filtered_rows)

        filtered_rows.sort(
            key=lambda row: (
                self.get_consultation_sort_value(row, self.sales_consult_tree_sort_column),
                normalize_text(str(row.get("period", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.sales_consult_tree_sort_reverse,
        )
        summary_rows.sort(
            key=lambda row: (
                self.get_consultation_sort_value(row, self.sales_consult_summary_sort_column),
                normalize_text(str(row.get("description", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.sales_consult_summary_sort_reverse,
        )
        self.sales_consult_filtered_rows = list(filtered_rows)
        self.sales_consult_summary_rows = list(summary_rows)

        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        row_count = 0
        runtime_rules = self.get_current_runtime_rules_for_highlight()

        for row_index, row in enumerate(filtered_rows):
            values = (
                row["period"],
                row.get("curve_abc", "C"),
                row["code"],
                row["description"],
                self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                row["ncm"],
                row.get("cest", ""),
                row["cfop"],
                row["cst_icms"],
                format_decimal_sped(extract_representative_icms_rate(row)),
                format_decimal_sped(extract_effective_icms_rate(row)),
                row["document_count"],
                row["launch_count"],
                format_decimal_sped(Decimal(row["sale_value"])),
                format_decimal_sped(Decimal(row["base_icms"])),
                format_decimal_sped(Decimal(row["icms_value"])),
                row["status"],
            )
            self.sales_consult_tree.insert("", END, iid=str(row_index), values=values, tags=self.get_consult_tree_tags(row, runtime_rules))
            total_sale += Decimal(row["sale_value"])
            total_base += Decimal(row["base_icms"])
            total_icms += Decimal(row["icms_value"])
            row_count += 1

        if hasattr(self, "sales_consult_summary_tree"):
            for row in summary_rows:
                self.sales_consult_summary_tree.insert(
                    "",
                    END,
                    values=(
                        row["code"],
                        row["description"],
                        row.get("curve_abc", "C"),
                        self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                        row["periods"],
                        row["cfops"],
                        row["csts"],
                        format_decimal_sped(Decimal(row.get("nominal_icms_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row.get("icms_rate", Decimal("0")))),
                        format_decimal_sped(Decimal(row["sale_value"])),
                        format_decimal_sped(Decimal(row["base_icms"])),
                        format_decimal_sped(Decimal(row["icms_value"])),
                    ),
                    tags=(
                        ("reduced" if bool(row.get("has_icms_reduction")) else "normal"),
                        *(() if int(row.get("supplier_count", 0)) <= 1 else ("multi_supplier",)),
                    ),
                )

        self.sales_consult_total_items_var.set(f"Linhas: {row_count}")
        self.sales_consult_total_products_var.set(f"Produtos: {len(summary_rows)}")
        self.sales_consult_total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
        self.sales_consult_total_base_var.set(f"Base ICMS: {format_decimal_sped(total_base)}")
        self.sales_consult_total_base_ratio_var.set(
            f"% Base/Oper: {format_decimal_sped((total_base * Decimal('100') / total_sale).quantize(Decimal('0.01')) if total_sale > 0 else Decimal('0.00'))}"
        )
        self.sales_consult_total_icms_var.set(f"Valor ICMS: {format_decimal_sped(total_icms)}")

    def get_contrib_consultation_sort_value(self, row: dict[str, object], column_id: str) -> object:
        value = row.get(column_id, "")
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return value
        return normalize_text(str(value))

    def get_contrib_row_tags(self, row: dict[str, object]) -> tuple[str, ...]:
        status_value = str(row.get("status", "")).strip()
        if status_value == "Ok":
            return ("ok",)
        if "Sem " in status_value:
            return ("warning",)
        return ("divergent",)

    def refresh_contrib_consultation_tree(self) -> None:
        if not hasattr(self, "contrib_consult_tree") or not getattr(self, "_contrib_consultation_filters_ready", False):
            return
        for item_id in self.contrib_consult_tree.get_children():
            self.contrib_consult_tree.delete(item_id)
        filtered_rows = self.get_contrib_consultation_filtered_rows()
        filtered_rows.sort(
            key=lambda row: (
                self.get_contrib_consultation_sort_value(row, self.contrib_consult_tree_sort_column),
                normalize_text(str(row.get("period", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.contrib_consult_tree_sort_reverse,
        )
        self.contrib_consult_filtered_rows = list(filtered_rows)
        total_sale = Decimal("0")
        total_pis_base = Decimal("0")
        total_cofins_base = Decimal("0")
        total_pis = Decimal("0")
        total_cofins = Decimal("0")
        seen_products: set[str] = set()
        for row_index, row in enumerate(filtered_rows):
            self.contrib_consult_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row.get("period", ""),
                    row.get("curve_abc", "C"),
                    row.get("code", ""),
                    row.get("description", ""),
                    self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                    row.get("ncm", ""),
                    row.get("cest", ""),
                    row.get("cfop", ""),
                    row.get("cst_pis", ""),
                    row.get("cst_cofins", ""),
                    format_decimal_sped(Decimal(row.get("aliquota_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("aliquota_cofins", Decimal("0")))),
                    row.get("document_count", 0),
                    row.get("launch_count", 0),
                    format_decimal_sped(Decimal(row.get("sale_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("cofins_value", Decimal("0")))),
                    row.get("status", ""),
                ),
                tags=self.get_contrib_row_tags(row),
            )
            seen_products.add(str(row.get("code", "")).strip())
            total_sale += Decimal(row.get("sale_value", Decimal("0")))
            total_pis_base += Decimal(row.get("base_pis", Decimal("0")))
            total_cofins_base += Decimal(row.get("base_cofins", Decimal("0")))
            total_pis += Decimal(row.get("pis_value", Decimal("0")))
            total_cofins += Decimal(row.get("cofins_value", Decimal("0")))
        self.contrib_consult_total_items_var.set(f"Linhas: {len(filtered_rows)}")
        self.contrib_consult_total_products_var.set(f"Produtos: {len(seen_products)}")
        self.contrib_consult_total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
        self.contrib_consult_total_pis_base_var.set(f"Base PIS: {format_decimal_sped(total_pis_base)}")
        self.contrib_consult_total_cofins_base_var.set(f"Base COFINS: {format_decimal_sped(total_cofins_base)}")
        self.contrib_consult_total_pis_var.set(f"Valor PIS: {format_decimal_sped(total_pis)}")
        self.contrib_consult_total_cofins_var.set(f"Valor COFINS: {format_decimal_sped(total_cofins)}")

    def refresh_contrib_sales_consultation_tree(self) -> None:
        if not hasattr(self, "contrib_sales_consult_tree") or not getattr(self, "_contrib_sales_consultation_filters_ready", False):
            return
        for item_id in self.contrib_sales_consult_tree.get_children():
            self.contrib_sales_consult_tree.delete(item_id)
        filtered_rows = self.get_contrib_sales_consultation_filtered_rows()
        filtered_rows.sort(
            key=lambda row: (
                self.get_contrib_consultation_sort_value(row, self.contrib_sales_consult_tree_sort_column),
                normalize_text(str(row.get("period", ""))),
                normalize_text(str(row.get("code", ""))),
            ),
            reverse=self.contrib_sales_consult_tree_sort_reverse,
        )
        self.contrib_sales_consult_filtered_rows = list(filtered_rows)
        total_sale = Decimal("0")
        total_pis_base = Decimal("0")
        total_cofins_base = Decimal("0")
        total_pis = Decimal("0")
        total_cofins = Decimal("0")
        seen_products: set[str] = set()
        for row_index, row in enumerate(filtered_rows):
            self.contrib_sales_consult_tree.insert(
                "",
                END,
                iid=str(row_index),
                values=(
                    row.get("period", ""),
                    row.get("curve_abc", "C"),
                    row.get("code", ""),
                    row.get("description", ""),
                    self.format_supplier_highlight(row.get("suppliers", ""), row.get("supplier_count", 0)),
                    row.get("ncm", ""),
                    row.get("cest", ""),
                    row.get("cfop", ""),
                    row.get("cst_pis", ""),
                    row.get("cst_cofins", ""),
                    format_decimal_sped(Decimal(row.get("aliquota_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("aliquota_cofins", Decimal("0")))),
                    row.get("document_count", 0),
                    row.get("launch_count", 0),
                    format_decimal_sped(Decimal(row.get("sale_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_pis", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("base_cofins", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("pis_value", Decimal("0")))),
                    format_decimal_sped(Decimal(row.get("cofins_value", Decimal("0")))),
                    row.get("status", ""),
                ),
                tags=self.get_contrib_row_tags(row),
            )
            seen_products.add(str(row.get("code", "")).strip())
            total_sale += Decimal(row.get("sale_value", Decimal("0")))
            total_pis_base += Decimal(row.get("base_pis", Decimal("0")))
            total_cofins_base += Decimal(row.get("base_cofins", Decimal("0")))
            total_pis += Decimal(row.get("pis_value", Decimal("0")))
            total_cofins += Decimal(row.get("cofins_value", Decimal("0")))
        self.contrib_sales_consult_total_items_var.set(f"Linhas: {len(filtered_rows)}")
        self.contrib_sales_consult_total_products_var.set(f"Produtos: {len(seen_products)}")
        self.contrib_sales_consult_total_sale_var.set(f"Valor Operacao: {format_decimal_sped(total_sale)}")
        self.contrib_sales_consult_total_pis_base_var.set(f"Base PIS: {format_decimal_sped(total_pis_base)}")
        self.contrib_sales_consult_total_cofins_base_var.set(f"Base COFINS: {format_decimal_sped(total_cofins_base)}")
        self.contrib_sales_consult_total_pis_var.set(f"Valor PIS: {format_decimal_sped(total_pis)}")
        self.contrib_sales_consult_total_cofins_var.set(f"Valor COFINS: {format_decimal_sped(total_cofins)}")

    def export_current_consultation_filter(self) -> None:
        if not self.consult_filtered_rows:
            messagebox.showwarning("Exportar consulta", "Nao ha dados filtrados para exportar.")
            return

        output = filedialog.asksaveasfilename(
            title="Salvar consulta filtrada",
            defaultextension=".xlsx",
            initialfile="consulta_entradas_filtrada.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not output:
            return

        output_path = Path(output)
        detail_headers = [
            "Periodo",
            "Arquivo SPED",
            "Curva ABC",
            "Codigo",
            "Descricao",
            "Fornecedor",
            "NCM",
            "CEST",
            "CFOP",
            "CST",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Documentos",
            "Lancamentos",
            "Quantidade",
            "Valor Operacao",
            "Base ICMS",
            "Valor ICMS",
            "Status",
        ]
        detail_rows = [
            [
                row["period"],
                row["file_name"],
                row.get("curve_abc", "C"),
                row["code"],
                row["description"],
                row.get("suppliers", ""),
                row["ncm"],
                row.get("cest", ""),
                row["cfop"],
                row["cst_icms"],
                extract_representative_icms_rate(row),
                extract_effective_icms_rate(row),
                row["document_count"],
                row["launch_count"],
                row["quantity"],
                row["sale_value"],
                row["base_icms"],
                row["icms_value"],
                row["status"],
            ]
            for row in self.consult_filtered_rows
        ]
        summary_headers = [
            "Codigo",
            "Descricao",
            "Curva ABC",
            "Fornecedores",
            "Periodos",
            "CEST",
            "CFOPs",
            "CSTs",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Valor Operacao",
            "Base ICMS",
            "Valor ICMS",
        ]
        summary_rows = [
            [
                row["code"],
                row["description"],
                row.get("curve_abc", "C"),
                row.get("suppliers", ""),
                row["periods"],
                row.get("cests", ""),
                row["cfops"],
                row["csts"],
                Decimal(row.get("nominal_icms_rate", Decimal("0"))),
                Decimal(row.get("icms_rate", Decimal("0"))),
                row["sale_value"],
                row["base_icms"],
                row["icms_value"],
            ]
            for row in self.consult_summary_rows
        ]
        try:
            if output_path.suffix.lower() == ".csv":
                summary_output = output_path.with_name(f"{output_path.stem}_resumo_produtos.csv")
                write_simple_csv_file(output_path, detail_headers, detail_rows)
                write_simple_csv_file(summary_output, summary_headers, summary_rows)
                self.log_message(f"Consulta filtrada exportada em CSV: {output_path} e {summary_output}")
                messagebox.showinfo(
                    "Exportar consulta",
                    f"Arquivos gerados com sucesso:\n{output_path}\n{summary_output}",
                )
                return

            write_simple_excel_workbook(
                output_path,
                [
                    ("Consulta Detalhada", detail_headers, detail_rows),
                    ("Resumo Produtos", summary_headers, summary_rows),
                ],
            )
            self.log_message(f"Consulta filtrada exportada com sucesso em: {output_path}")
            messagebox.showinfo("Exportar consulta", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar consulta filtrada: {exc}")
            messagebox.showerror("Exportar consulta", str(exc))

    def export_current_sales_consultation_filter(self) -> None:
        if not self.sales_consult_filtered_rows:
            messagebox.showwarning("Exportar consulta", "Nao ha dados filtrados para exportar.")
            return

        output = filedialog.asksaveasfilename(
            title="Salvar consulta filtrada",
            defaultextension=".xlsx",
            initialfile="consulta_saidas_filtrada.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not output:
            return

        output_path = Path(output)
        detail_headers = [
            "Periodo",
            "Arquivo SPED",
            "Curva ABC",
            "Codigo",
            "Descricao",
            "Fornecedor",
            "NCM",
            "CEST",
            "CFOP",
            "CST",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Documentos",
            "Lancamentos",
            "Quantidade",
            "Valor Operacao",
            "Base ICMS",
            "Valor ICMS",
            "Status",
        ]
        detail_rows = [
            [
                row["period"],
                row["file_name"],
                row.get("curve_abc", "C"),
                row["code"],
                row["description"],
                row.get("suppliers", ""),
                row["ncm"],
                row.get("cest", ""),
                row["cfop"],
                row["cst_icms"],
                extract_representative_icms_rate(row),
                extract_effective_icms_rate(row),
                row["document_count"],
                row["launch_count"],
                row["quantity"],
                row["sale_value"],
                row["base_icms"],
                row["icms_value"],
                row["status"],
            ]
            for row in self.sales_consult_filtered_rows
        ]
        summary_headers = [
            "Codigo",
            "Descricao",
            "Curva ABC",
            "Fornecedores",
            "Periodos",
            "CEST",
            "CFOPs",
            "CSTs",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Valor Operacao",
            "Base ICMS",
            "Valor ICMS",
        ]
        summary_rows = [
            [
                row["code"],
                row["description"],
                row.get("curve_abc", "C"),
                row.get("suppliers", ""),
                row["periods"],
                row.get("cests", ""),
                row["cfops"],
                row["csts"],
                Decimal(row.get("nominal_icms_rate", Decimal("0"))),
                Decimal(row.get("icms_rate", Decimal("0"))),
                row["sale_value"],
                row["base_icms"],
                row["icms_value"],
            ]
            for row in self.sales_consult_summary_rows
        ]
        try:
            if output_path.suffix.lower() == ".csv":
                summary_output = output_path.with_name(f"{output_path.stem}_resumo_produtos.csv")
                write_simple_csv_file(output_path, detail_headers, detail_rows)
                write_simple_csv_file(summary_output, summary_headers, summary_rows)
                self.log_message(f"Consulta de saidas exportada em CSV: {output_path} e {summary_output}")
                messagebox.showinfo(
                    "Exportar consulta",
                    f"Arquivos gerados com sucesso:\n{output_path}\n{summary_output}",
                )
                return

            write_simple_excel_workbook(
                output_path,
                [
                    ("Consulta Detalhada", detail_headers, detail_rows),
                    ("Resumo Produtos", summary_headers, summary_rows),
                ],
            )
            self.log_message(f"Consulta de saidas exportada com sucesso em: {output_path}")
            messagebox.showinfo("Exportar consulta", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar consulta de saidas: {exc}")
            messagebox.showerror("Exportar consulta", str(exc))

    def export_current_contrib_consultation_filter(self) -> None:
        self.export_contrib_consultation_rows(self.contrib_consult_filtered_rows, "consulta_entradas_pis_cofins.xlsx")

    def export_current_contrib_sales_consultation_filter(self) -> None:
        self.export_contrib_consultation_rows(self.contrib_sales_consult_filtered_rows, "consulta_saidas_pis_cofins.xlsx")

    def export_contrib_consultation_rows(self, rows: list[dict[str, object]], initialfile: str) -> None:
        if not rows:
            messagebox.showwarning("Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output = filedialog.asksaveasfilename(
            title="Salvar consulta filtrada",
            defaultextension=".xlsx",
            initialfile=initialfile,
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")],
        )
        if not output:
            return
        output_path = Path(output)
        headers = [
            "Periodo", "Arquivo SPED", "Curva ABC", "Codigo", "Descricao", "Participante", "NCM", "CEST",
            "CFOP", "CST PIS", "CST COFINS", "Aliq PIS", "Aliq COFINS", "Documentos", "Lancamentos", "Quantidade",
            "Valor Operacao", "Base PIS", "Base COFINS", "Valor PIS", "Valor COFINS", "Status",
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
                Decimal(row.get("aliquota_pis", Decimal("0"))),
                Decimal(row.get("aliquota_cofins", Decimal("0"))),
                row.get("document_count", 0),
                row.get("launch_count", 0),
                Decimal(row.get("quantity", Decimal("0"))),
                Decimal(row.get("sale_value", Decimal("0"))),
                Decimal(row.get("base_pis", Decimal("0"))),
                Decimal(row.get("base_cofins", Decimal("0"))),
                Decimal(row.get("pis_value", Decimal("0"))),
                Decimal(row.get("cofins_value", Decimal("0"))),
                row.get("status", ""),
            ]
            for row in rows
        ]
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, export_rows)
            else:
                write_simple_excel_workbook(output_path, [("Consulta", headers, export_rows)])
            self.log_message(f"Consulta PIS/COFINS exportada: {output_path}")
            messagebox.showinfo("Exportar consulta", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar consulta PIS/COFINS: {exc}")
            messagebox.showerror("Exportar consulta", str(exc))

    def export_launch_details(self, launch_details: list[dict[str, object]], row: dict[str, object], output_type: str) -> None:
        if not launch_details:
            messagebox.showwarning("Exportar popup", "Nao ha lancamentos para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Salvar detalhamento do popup",
            defaultextension=suffix,
            initialfile=f"detalhamento_{row['code']}_{str(row['period']).replace('/', '_')}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        headers = [
            "Periodo",
            "Codigo",
            "Descricao",
            "Data",
            "Documento",
            "Fornecedor",
            "Chave",
            "Item",
            "CEST",
            "CFOP",
            "CST",
            "Aliq ICMS",
            "Aliq Efetiva",
            "Quantidade",
            "Valor Operacao",
            "Valor IPI",
            "Base ICMS",
            "Dif. Oper/Base",
            "Origem Dif.",
            "Valor ICMS",
        ]
        data_rows = [
            [
                row["period"],
                row["code"],
                row["description"],
                detail.get("document_date", ""),
                detail.get("document_number", ""),
                detail.get("participant_name", ""),
                detail.get("document_key", ""),
                detail.get("item_number", ""),
                detail.get("cest", ""),
                detail.get("cfop", ""),
                detail.get("cst_icms", ""),
                Decimal(detail.get("icms_rate", Decimal("0"))),
                compute_display_icms_rate(
                    detail.get("icms_rate", Decimal("0")),
                    get_launch_total_operation_value(detail),
                    Decimal(detail.get("base_icms", Decimal("0"))),
                    Decimal(detail.get("icms_value", Decimal("0"))),
                ),
                Decimal(detail.get("quantity", Decimal("0"))),
                get_launch_total_operation_value(detail),
                Decimal(detail.get("ipi_value", Decimal("0"))),
                Decimal(detail.get("base_icms", Decimal("0"))),
                get_operation_base_difference(detail),
                describe_operation_base_difference(
                    get_launch_total_operation_value(detail),
                    Decimal(detail.get("base_icms", Decimal("0"))),
                    Decimal(detail.get("ipi_value", Decimal("0"))),
                    Decimal(detail.get("icms_st_value", Decimal("0"))),
                    Decimal(detail.get("discount_value", Decimal("0"))),
                ),
                Decimal(detail.get("icms_value", Decimal("0"))),
            ]
            for detail in launch_details
        ]
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, data_rows)
            else:
                write_simple_excel_workbook(output_path, [("Detalhamento", headers, data_rows)])
            self.log_message(f"Detalhamento exportado com sucesso em: {output_path}")
            messagebox.showinfo("Exportar popup", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar detalhamento: {exc}")
            messagebox.showerror("Exportar popup", str(exc))

    def export_simple_popup_dataset(
        self,
        title: str,
        initialfile: str,
        sheet_name: str,
        headers: list[str],
        rows: list[list[object]],
        output_type: str,
    ) -> None:
        if not rows:
            messagebox.showwarning(title, "Nao ha dados para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title=title,
            defaultextension=suffix,
            initialfile=f"{initialfile}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                write_simple_excel_workbook(output_path, [(sheet_name, headers, rows)])
            self.log_message(f"Exportacao do popup concluida em: {output_path}")
            messagebox.showinfo(title, f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar popup: {exc}")
            messagebox.showerror(title, str(exc))

    def export_product_monthly_popup_dataset(
        self,
        initialfile: str,
        sheet_name: str,
        headers: list[str],
        rows: list[list[object]],
        periods: list[str],
        output_type: str,
    ) -> None:
        if not rows:
            messagebox.showwarning("Exportar popup", "Nao ha dados para exportar.")
            return
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title="Exportar popup",
            defaultextension=suffix,
            initialfile=f"{initialfile}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, headers, rows)
            else:
                divergence_rows = build_monthly_aliquota_divergence_rows(headers, rows, periods)
                sheets = [(sheet_name, headers, rows, periods)]
                if divergence_rows:
                    sheets.append(("Divergencia Aliquota", headers, divergence_rows, periods))
                write_monthly_colored_excel_workbook_with_sheets(output_path, sheets)
            self.log_message(f"Exportacao do popup concluida em: {output_path}")
            messagebox.showinfo("Exportar popup", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar popup: {exc}")
            messagebox.showerror("Exportar popup", str(exc))

    def export_credit_diagnostic_dataset(
        self,
        summary_rows: list[list[object]],
        detail_rows: list[list[object]],
        output_type: str,
        raw_detail_rows: list[dict[str, object]] | None = None,
        dialog_title: str = "Diagnostico de credito",
        initial_filename_base: str = "diagnostico_credito_icms",
    ) -> None:
        suffix = ".csv" if output_type == "csv" else ".xlsx"
        selected = filedialog.asksaveasfilename(
            title=f"Exportar {dialog_title.lower()}",
            defaultextension=suffix,
            initialfile=f"{initial_filename_base}{suffix}",
            filetypes=[("Arquivo CSV", "*.csv"), ("Arquivo Excel", "*.xlsx")],
        )
        if not selected:
            return

        output_path = Path(selected)
        summary_headers = ["Periodo", "Motivo", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        detail_headers = ["Periodo", "Motivo", "Registro", "CST", "CFOP", "Aliq", "Aliq Efetiva", "Linhas", "Valor Operacao", "Base ICMS", "Perda de Base", "% Base/Oper", "Valor ICMS"]
        try:
            if output_path.suffix.lower() == ".csv":
                write_simple_csv_file(output_path, detail_headers, detail_rows)
                summary_output = output_path.with_name(f"{output_path.stem}_resumo.csv")
                write_simple_csv_file(summary_output, summary_headers, summary_rows)
            else:
                workbook_sheets: list[tuple[str, list[str], list[list[object]]]] = [
                    ("Resumo Motivos", summary_headers, summary_rows),
                    ("Detalhamento Fiscal", detail_headers, detail_rows),
                ]
                if raw_detail_rows:
                    grouped_item_headers = [
                        "Periodo",
                        "Motivo",
                        "Registro",
                        "CST",
                        "CFOP",
                        "Aliq",
                        "Aliq Efetiva",
                        "Linha SPED",
                        "Valor Operacao",
                        "Base ICMS",
                        "Perda de Base",
                        "% Base/Oper",
                        "Valor ICMS",
                        "Linha Original",
                    ]
                    product_item_headers = [
                        "Periodo",
                        "Motivo",
                        "Registro",
                        "CST",
                        "CFOP",
                        "Codigo",
                        "Produto",
                        "NCM",
                        "Aliq",
                        "Aliq Efetiva",
                        "Linhas",
                        "Valor Operacao",
                        "Base ICMS",
                        "% Base/Oper",
                        "Valor ICMS",
                    ]
                    all_grouped_rows: list[list[object]] = []
                    all_product_rows, product_footer_rows = build_credit_diagnostic_product_sheet_payload(raw_detail_rows)
                    grouped_rows_by_reason: dict[str, list[list[object]]] = defaultdict(list)

                    for detail_row in raw_detail_rows:
                        period = str(detail_row.get("period", ""))
                        reason = str(detail_row.get("reason", ""))
                        source_register = str(detail_row.get("source_register", ""))
                        cst_icms = str(detail_row.get("cst_icms", ""))
                        cfop = str(detail_row.get("cfop", ""))
                        nominal_rate = Decimal(detail_row.get("icms_rate", Decimal("0")))

                        for grouped_row in detail_row.get("grouped_rows", []):
                            operation_value = Decimal(grouped_row.get("total_operation_value", Decimal("0")))
                            base_icms = Decimal(grouped_row.get("base_icms", Decimal("0")))
                            icms_value = Decimal(grouped_row.get("icms_value", Decimal("0")))
                            base_gap = max(Decimal("0"), operation_value - base_icms)
                            base_ratio = (
                                (base_icms * Decimal("100") / operation_value).quantize(Decimal("0.01"))
                                if operation_value > 0
                                else Decimal("0.00")
                            )
                            effective_rate = compute_display_icms_rate(
                                Decimal(grouped_row.get("icms_rate", Decimal("0"))),
                                operation_value,
                                base_icms,
                                icms_value,
                            )
                            export_row = [
                                period,
                                reason,
                                source_register,
                                cst_icms,
                                cfop,
                                Decimal(grouped_row.get("icms_rate", nominal_rate)),
                                effective_rate,
                                grouped_row.get("line_number", ""),
                                operation_value,
                                base_icms,
                                base_gap,
                                base_ratio,
                                icms_value,
                                grouped_row.get("raw_line", ""),
                            ]
                            all_grouped_rows.append(export_row)
                            grouped_rows_by_reason[reason].append(export_row)

                    if all_grouped_rows:
                        workbook_sheets.append(("Itens Detalhados", grouped_item_headers, all_grouped_rows))
                    if all_product_rows:
                        workbook_sheets.append(
                            (
                                "Produtos Detalhados",
                                product_item_headers,
                                all_product_rows,
                                {"footer_rows": product_footer_rows},
                            )
                        )
                    for reason, reason_rows in sorted(grouped_rows_by_reason.items(), key=lambda item: item[0]):
                        if reason_rows:
                            workbook_sheets.append((f"Motivo {reason}", grouped_item_headers, reason_rows))

                write_simple_excel_workbook(output_path, workbook_sheets)
            self.log_message(f"{dialog_title} exportado em: {output_path}")
            messagebox.showinfo("Exportar diagnostico", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar {dialog_title.lower()}: {exc}")
            messagebox.showerror("Exportar diagnostico", str(exc))

    def clear_consultation_filters(self) -> None:
        self.consult_cst_filter_var.set("")
        self.consult_cfop_filter_var.set("")
        self.consult_status_filter_var.set("Todos")
        self.consult_search_var.set("")
        self.consult_selected_product_code = ""
        if hasattr(self, "consult_period_check_vars"):
            for variable in self.consult_period_check_vars.values():
                variable.set(True)
        self.refresh_consultation_tree()

    def clear_sales_consultation_filters(self) -> None:
        self.sales_consult_cst_filter_var.set("")
        self.sales_consult_cfop_filter_var.set("")
        self.sales_consult_status_filter_var.set("Todos")
        self.sales_consult_search_var.set("")
        self.contrib_consult_cst_filter_var.set("")
        self.contrib_consult_cfop_filter_var.set("")
        self.contrib_consult_status_filter_var.set("Todos")
        self.contrib_consult_search_var.set("")
        self.contrib_sales_consult_cst_filter_var.set("")
        self.contrib_sales_consult_cfop_filter_var.set("")
        self.contrib_sales_consult_status_filter_var.set("Todos")
        self.contrib_sales_consult_search_var.set("")
        self.sales_consult_selected_product_code = ""
        if hasattr(self, "sales_consult_period_check_vars"):
            for variable in self.sales_consult_period_check_vars.values():
                variable.set(True)
        self.refresh_sales_consultation_tree()

    def clear_contrib_consultation_filters(self) -> None:
        self.contrib_consult_cst_filter_var.set("")
        self.contrib_consult_cfop_filter_var.set("")
        self.contrib_consult_status_filter_var.set("Todos")
        self.contrib_consult_search_var.set("")
        self.contrib_consult_selected_product_code = ""
        if hasattr(self, "contrib_consult_period_check_vars"):
            for variable in self.contrib_consult_period_check_vars.values():
                variable.set(True)
        self.refresh_contrib_consultation_tree()

    def clear_contrib_sales_consultation_filters(self) -> None:
        self.contrib_sales_consult_cst_filter_var.set("")
        self.contrib_sales_consult_cfop_filter_var.set("")
        self.contrib_sales_consult_status_filter_var.set("Todos")
        self.contrib_sales_consult_search_var.set("")
        self.contrib_sales_consult_selected_product_code = ""
        if hasattr(self, "contrib_sales_consult_period_check_vars"):
            for variable in self.contrib_sales_consult_period_check_vars.values():
                variable.set(True)
        self.refresh_contrib_sales_consultation_tree()

    def clear_tree_items(self, tree_name: str) -> None:
        tree = getattr(self, tree_name, None)
        if tree is not None:
            children = tree.get_children()
            if children:
                tree.delete(*children)

    def clear_widget_children(self, widget_name: str) -> None:
        widget = getattr(self, widget_name, None)
        if widget is not None:
            for child in widget.winfo_children():
                child.destroy()

    def clear_compare_screen(self, show_message: bool = True) -> None:
        self.compare_sped_icms_path_var.set("")
        self.compare_sped_contrib_path_var.set("")
        self.compare_xml_folder_var.set("")
        self.compare_sheet_path_var.set("")
        self.compare_mode_var.set("icms")
        self.compare_operation_scope_var.set("Ambos")
        self.compare_progress_var.set(0)
        self.compare_stats_var.set("Nenhuma comparacao executada.")
        self.compare_status_var.set("Selecione os SPEDs e a pasta de XML ou uma planilha.")
        self.compare_last_results = []
        self.compare_last_result_origin = ""
        self.compare_last_operation_scope = ""
        self.compare_last_mode = ""
        self.compare_last_stats = {}
        self.clear_tree_items("compare_tree")
        self.compare_diagnostic_status_var.set("Execute uma comparacao em SPED x XML/Planilha e clique em Atualizar diagnostico.")
        self.compare_diagnostic_conclusion_var.set("Nenhum diagnostico carregado.")
        self.compare_diagnostic_footer_var.set("")
        for tree_name in ("compare_diagnostic_motive_tree", "compare_diagnostic_check_tree", "compare_diagnostic_composition_tree"):
            self.clear_tree_items(tree_name)
        if hasattr(self, "compare_run_btn"):
            self.compare_run_btn.configure(state="normal")
        self.write_audit_log("LIMPAR_TELA", "Tela SPED x XML/Planilha limpa.")
        if show_message:
            self.compare_status_var.set("Tela limpa. Selecione os arquivos para iniciar novamente.")

    def clear_xml_screen(self, show_message: bool = True) -> None:
        self.xml_sped_path_var.set("")
        self.xml_source_path_var.set("")
        self.xml_company_tax_id_var.set("CNPJ/CPF empresa: nao carregado")
        self.xml_operation_scope_var.set("Todos")
        self.xml_status_var.set("Selecione uma pasta ou arquivo XML para processar.")
        self.xml_progress_var.set(0)
        self.xml_summary_all_rows = []
        self.xml_summary_rows = []
        self.xml_summary_stats = {}
        self.xml_ignored_log_rows = []
        self.xml_document_log_rows = []
        self.xml_cancelled_display_rows = []
        self.xml_total_files_var.set("XMLs: 0")
        self.xml_total_operation_files_var.set("XMLs da operacao: 0")
        self.xml_total_cfops_var.set("CFOPs: 0")
        self.xml_total_operation_var.set("Total Operacao: 0,00")
        self.xml_total_base_icms_var.set("Base ICMS: 0,00")
        self.xml_total_icms_var.set("Valor ICMS: 0,00")
        self.clear_tree_items("xml_summary_tree")

        self.xml_credit_status_var.set("Processe os XMLs para listar os creditos destacados nas entradas.")
        self.xml_credit_invoice_rows = []
        self.xml_credit_cfop_rows = []
        self.xml_credit_stats = {}
        self.xml_credit_total_notes_var.set("Notas entrada: 0")
        self.xml_credit_total_items_var.set("Itens: 0")
        self.xml_credit_total_operation_var.set("Total Operacao: 0,00")
        self.xml_credit_total_base_icms_var.set("Base ICMS: 0,00")
        self.xml_credit_total_icms_var.set("Credito ICMS: 0,00")
        self.xml_credit_total_ipi_var.set("Credito IPI: 0,00")
        self.clear_tree_items("xml_credit_notes_tree")
        self.clear_tree_items("xml_credit_cfop_tree")
        if hasattr(self, "xml_process_btn"):
            self.xml_process_btn.configure(state="normal")
        self.write_audit_log("LIMPAR_TELA", "Tela XML limpa.")
        if show_message:
            self.xml_status_var.set("Tela limpa. Selecione o SPED/XML para iniciar novamente.")

    def clear_form(self) -> None:
        previous_filter_flags = (
            getattr(self, "_consultation_filters_ready", False),
            getattr(self, "_sales_consultation_filters_ready", False),
            getattr(self, "_contrib_consultation_filters_ready", False),
            getattr(self, "_contrib_sales_consultation_filters_ready", False),
        )
        self._consultation_filters_ready = False
        self._sales_consultation_filters_ready = False
        self._contrib_consultation_filters_ready = False
        self._contrib_sales_consultation_filters_ready = False
        self.sped_path_var.set("")
        self.multi_sped_paths_var.set("")
        self.consult_sped_paths_var.set("")
        self.sales_consult_sped_paths_var.set("")
        self.contrib_consult_sped_paths_var.set("")
        self.contrib_sales_consult_sped_paths_var.set("")
        self.contrib_consult_xml_paths_var.set("")
        self.contrib_sales_consult_xml_paths_var.set("")
        self.consult_xml_paths_var.set("")
        self.sales_consult_xml_paths_var.set("")
        self.excel_path_var.set("")
        self.xml_path_var.set("")
        self.contrib_path_var.set("")
        self.nfce_ncm_filter_var.set("")
        self.icms_rate_var.set("")
        self.cst_filter_var.set("")
        self.cfop_filter_var.set("")
        self.new_cst_var.set("")
        self.new_cfop_var.set("")
        self.zero_icms_cfop_var.set("")
        self.adjustment_mode_var.set("Filtros")
        self.adjusted_sped_scope_var.set("Ambos")
        self.consult_cst_filter_var.set("")
        self.consult_cfop_filter_var.set("")
        self.consult_status_filter_var.set("Todos")
        self.consult_search_var.set("")
        self.sales_consult_cst_filter_var.set("")
        self.sales_consult_cfop_filter_var.set("")
        self.sales_consult_status_filter_var.set("Todos")
        self.sales_consult_search_var.set("")
        self.consult_period_labels = []
        self.consult_comparison_rows = []
        self.consult_filtered_rows = []
        self.consult_summary_rows = []
        self.sales_consult_period_labels = []
        self.sales_consult_comparison_rows = []
        self.sales_consult_filtered_rows = []
        self.sales_consult_summary_rows = []
        self.contrib_consult_period_labels = []
        self.contrib_consult_comparison_rows = []
        self.contrib_consult_filtered_rows = []
        self.contrib_sales_consult_period_labels = []
        self.contrib_sales_consult_comparison_rows = []
        self.contrib_sales_consult_filtered_rows = []
        self.consult_period_path_map = {}
        self.sales_consult_period_path_map = {}
        self.contrib_consult_period_path_map = {}
        self.contrib_sales_consult_period_path_map = {}
        self.consult_tree_sort_column = "code"
        self.consult_tree_sort_reverse = False
        self.consult_summary_sort_column = "code"
        self.consult_summary_sort_reverse = False
        self.sales_consult_tree_sort_column = "code"
        self.sales_consult_tree_sort_reverse = False
        self.sales_consult_summary_sort_column = "code"
        self.sales_consult_summary_sort_reverse = False
        self.contrib_consult_tree_sort_column = "code"
        self.contrib_consult_tree_sort_reverse = False
        self.contrib_sales_consult_tree_sort_column = "code"
        self.contrib_sales_consult_tree_sort_reverse = False
        self.consult_selected_product_code = ""
        self.sales_consult_selected_product_code = ""
        self.contrib_consult_selected_product_code = ""
        self.contrib_sales_consult_selected_product_code = ""
        self.filter_text.delete("1.0", END)
        self.runtime_rules_text.delete("1.0", END)
        self.clear_widget_children("consult_periods_checks_box")
        self.clear_widget_children("sales_consult_periods_checks_box")
        self.clear_widget_children("contrib_consult_periods_checks_box")
        self.clear_widget_children("contrib_sales_consult_periods_checks_box")
        self.consult_period_check_vars = {}
        self.sales_consult_period_check_vars = {}
        self.contrib_consult_period_check_vars = {}
        self.contrib_sales_consult_period_check_vars = {}
        for tree_name in (
            "consult_tree",
            "sales_consult_tree",
            "contrib_consult_tree",
            "contrib_sales_consult_tree",
            "consult_summary_tree",
            "sales_consult_summary_tree",
            "contrib_consult_summary_tree",
            "contrib_sales_consult_summary_tree",
        ):
            self.clear_tree_items(tree_name)
        self.consult_total_items_var.set("Linhas: 0")
        self.consult_total_products_var.set("Produtos: 0")
        self.consult_total_sale_var.set("Valor Operacao: 0,00")
        self.consult_total_base_var.set("Base ICMS: 0,00")
        self.consult_total_base_ratio_var.set("% Base/Oper: 0,00")
        self.consult_total_icms_var.set("Valor ICMS: 0,00")
        self.sales_consult_total_items_var.set("Linhas: 0")
        self.sales_consult_total_products_var.set("Produtos: 0")
        self.sales_consult_total_sale_var.set("Valor Operacao: 0,00")
        self.sales_consult_total_base_var.set("Base ICMS: 0,00")
        self.sales_consult_total_base_ratio_var.set("% Base/Oper: 0,00")
        self.sales_consult_total_icms_var.set("Valor ICMS: 0,00")
        self.contrib_consult_total_items_var.set("Linhas: 0")
        self.contrib_consult_total_products_var.set("Produtos: 0")
        self.contrib_consult_total_sale_var.set("Valor Operacao: 0,00")
        self.contrib_consult_total_pis_base_var.set("Base PIS: 0,00")
        self.contrib_consult_total_cofins_base_var.set("Base COFINS: 0,00")
        self.contrib_consult_total_pis_var.set("Valor PIS: 0,00")
        self.contrib_consult_total_cofins_var.set("Valor COFINS: 0,00")
        self.contrib_sales_consult_total_items_var.set("Linhas: 0")
        self.contrib_sales_consult_total_products_var.set("Produtos: 0")
        self.contrib_sales_consult_total_sale_var.set("Valor Operacao: 0,00")
        self.contrib_sales_consult_total_pis_base_var.set("Base PIS: 0,00")
        self.contrib_sales_consult_total_cofins_base_var.set("Base COFINS: 0,00")
        self.contrib_sales_consult_total_pis_var.set("Valor PIS: 0,00")
        self.contrib_sales_consult_total_cofins_var.set("Valor COFINS: 0,00")
        self.clear_compare_screen(show_message=False)
        self.clear_xml_screen(show_message=False)
        self.clear_tree_items("log")
        self.status_var.set("Selecione o SPED, a planilha, ou ambos para gerar o Excel.")
        (
            self._consultation_filters_ready,
            self._sales_consultation_filters_ready,
            self._contrib_consultation_filters_ready,
            self._contrib_sales_consultation_filters_ready,
        ) = previous_filter_flags
        self.write_audit_log("LIMPAR_TELA", "Limpeza geral executada; telas retornaram ao estado inicial.")
        messagebox.showinfo("Limpar Tela", "Tela limpa. O sistema voltou ao estado inicial.")

    def generate_multi_sped_entry_analysis_excel(self) -> None:
        sped_paths = parse_selected_paths(self.multi_sped_paths_var.get())
        if not sped_paths:
            sped_paths = parse_selected_paths(self.consult_sped_paths_var.get())
        if not sped_paths:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione de 1 a 12 arquivos SPED para gerar a analise.")
            return
        if len(sped_paths) > 12:
            messagebox.showwarning("Limite excedido", "A analise aceita no maximo 12 arquivos SPED por vez.")
            return

        output = filedialog.asksaveasfilename(
            title="Salvar analise de entradas",
            defaultextension=".xlsx",
            initialfile="analise_entradas_produtos.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not output:
            return

        try:
            self.log_message("Lendo SPEDs selecionados para a analise de entradas...")
            (
                summary_headers,
                summary_rows,
                headers,
                rows,
                divergence_headers,
                divergence_rows,
                monthly_summary_headers,
                monthly_summary_rows,
            ) = build_multi_sped_entry_analysis(sped_paths)
            self.log_message("Gerando Excel da analise de entradas...")
            write_simple_excel_workbook(
                Path(output),
                [
                    ("Resumo Produtos", summary_headers, summary_rows),
                    ("Resumo Mensal", monthly_summary_headers, monthly_summary_rows),
                    ("Analise Entradas", headers, rows),
                    ("Divergencias", divergence_headers, divergence_rows),
                ],
            )
            self.log_message(f"Analise de entradas gerada com sucesso em: {output}")
            self.show_front_message("showinfo", "Processo concluido", f"Arquivo gerado com sucesso:\n{output}")
        except Exception as exc:
            self.log_message(f"Falha ao gerar analise de entradas: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))

    def generate_nfce_ncm_excel(self) -> None:
        xml_source = self.xml_path_var.get().strip()
        ncm_filters = parse_filter_values(self.nfce_ncm_filter_var.get())
        if not xml_source:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione ao menos um XML ou uma pasta com XMLs da NFC-e.")
            return
        if not ncm_filters:
            messagebox.showwarning("Filtro obrigatorio", "Informe ao menos um NCM para exportar os itens.")
            return

        xml_paths = parse_selected_paths(xml_source)
        missing_xml_paths = [str(path) for path in xml_paths if not path.exists()]
        if missing_xml_paths:
            messagebox.showerror("Arquivo nao encontrado", "Os seguintes XMLs/pastas nao existem mais:\n" + "\n".join(missing_xml_paths))
            return

        base_name = xml_paths[0].stem if len(xml_paths) == 1 and xml_paths[0].is_file() else "nfce_ncm"
        output = filedialog.asksaveasfilename(
            title="Salvar exportacao NFC-e por NCM",
            defaultextension=".xlsx",
            initialfile=f"{base_name}_itens_nfce_ncm.xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not output:
            return

        try:
            self.log_message("Lendo XMLs de NFC-e e aplicando filtro por NCM...")
            total_xmls, total_itens, total_ignorados = export_nfce_items_by_ncm(xml_paths, ncm_filters, Path(output))
            self.log_message(
                f"Exportacao NFC-e concluida: {total_itens} item(ns) de {total_xmls} XML(s). XMLs ignorados: {total_ignorados}."
            )
            self.show_front_message("showinfo", "Processo concluido", f"Arquivo gerado com sucesso:\n{output}")
        except Exception as exc:
            self.log_message(f"Falha ao exportar NFC-e por NCM: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))

    def load_filter_descriptions(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o arquivo de filtros",
            filetypes=[("Arquivos suportados", "*.txt *.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if not selected:
            return

        try:
            descriptions = read_filter_descriptions_file(Path(selected))
            self.filter_text.delete("1.0", END)
            if descriptions:
                self.filter_text.insert("1.0", "\n".join(descriptions))
            self.log_message(f"Filtros carregados: {selected}")
        except Exception as exc:
            self.log_message(f"Falha ao carregar filtros: {exc}")
            messagebox.showerror("Erro ao carregar filtros", str(exc))

    def load_runtime_rule_history(self) -> None:
        if not self.runtime_rule_history_path.exists():
            self.runtime_rule_history = []
            return
        try:
            loaded = json.loads(self.runtime_rule_history_path.read_text(encoding="utf-8"))
            self.runtime_rule_history = loaded if isinstance(loaded, list) else []
        except Exception:
            self.runtime_rule_history = []

    def save_runtime_rule_history(self) -> None:
        self.runtime_rule_history_path.parent.mkdir(parents=True, exist_ok=True)
        self.runtime_rule_history_path.write_text(
            json.dumps(self.runtime_rule_history, ensure_ascii=True, indent=2),
            encoding="utf-8",
        )

    def clear_runtime_rule_filter(self) -> None:
        self.runtime_rule_filter_var.set("")
        self.refresh_runtime_rule_history_list()

    def update_runtime_rule_history_preview(self) -> None:
        if not hasattr(self, "runtime_rule_history_preview"):
            return
        self.runtime_rule_history_preview.configure(state="normal")
        self.runtime_rule_history_preview.delete("1.0", END)
        if not hasattr(self, "runtime_rule_history_tree"):
            self.runtime_rule_history_preview.configure(state="disabled")
            return
        selected = self.runtime_rule_history_tree.selection()
        if not selected:
            self.runtime_rule_history_preview.configure(state="disabled")
            return
        self.runtime_rule_history_preview.insert("1.0", selected[0])
        self.runtime_rule_history_preview.configure(state="disabled")

    def refresh_runtime_rule_history_list(self) -> None:
        if not hasattr(self, "runtime_rule_history_tree"):
            return
        filter_text = normalize_text(self.runtime_rule_filter_var.get())
        for item_id in self.runtime_rule_history_tree.get_children():
            self.runtime_rule_history_tree.delete(item_id)

        def sort_key(item: dict[str, object]) -> tuple[str, str]:
            return (str(item.get("last_used_at", "")), str(item.get("rule_line", "")))

        for item in sorted(self.runtime_rule_history, key=sort_key, reverse=True):
            rule_line = str(item.get("rule_line", "")).strip()
            searchable_text = normalize_text(f"{rule_line} {item.get('summary', '')}")
            if filter_text and filter_text not in searchable_text:
                continue
            self.runtime_rule_history_tree.insert(
                "",
                END,
                iid=rule_line,
                values=(rule_line,),
            )
        self.update_runtime_rule_history_preview()

    def remember_runtime_rules(self, rule_lines: list[str], source: str) -> None:
        if not rule_lines:
            return
        now = dt.datetime.now().replace(microsecond=0).isoformat()
        history_index = {
            str(item.get("rule_line", "")).strip(): item
            for item in self.runtime_rule_history
            if str(item.get("rule_line", "")).strip()
        }
        for rule_line in rule_lines:
            normalized_line = str(rule_line or "").strip()
            if not normalized_line:
                continue
            item = history_index.get(normalized_line)
            if item is None:
                item = {
                    "rule_line": normalized_line,
                    "summary": runtime_rule_summary(normalized_line),
                    "use_count": 0,
                    "created_at": now,
                    "last_used_at": now,
                    "last_source": source,
                }
                self.runtime_rule_history.append(item)
                history_index[normalized_line] = item
            item["summary"] = runtime_rule_summary(normalized_line)
            item["use_count"] = int(item.get("use_count", 0)) + 1
            item["last_used_at"] = now
            item["last_source"] = source
        self.save_runtime_rule_history()
        self.refresh_runtime_rule_history_list()

    def insert_selected_runtime_rule(self) -> None:
        if not hasattr(self, "runtime_rule_history_tree"):
            return
        selected = self.runtime_rule_history_tree.selection()
        if not selected:
            messagebox.showwarning("Historico", "Selecione uma regra do historico.")
            return
        self.append_runtime_rule_line(selected[0])

    def delete_selected_runtime_rule(self) -> None:
        if not hasattr(self, "runtime_rule_history_tree"):
            return
        selected = self.runtime_rule_history_tree.selection()
        if not selected:
            messagebox.showwarning("Historico", "Selecione uma regra do historico.")
            return
        rule_line = str(selected[0]).strip()
        if not rule_line:
            return
        if not messagebox.askyesno("Historico", "Deseja excluir a regra selecionada do historico?"):
            return
        self.runtime_rule_history = [
            item
            for item in self.runtime_rule_history
            if str(item.get("rule_line", "")).strip() != rule_line
        ]
        self.save_runtime_rule_history()
        self.refresh_runtime_rule_history_list()
        self.log_message("Regra removida do historico.")

    def append_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        current_text = self.runtime_rules_text.get("1.0", END).strip()
        if current_text:
            self.runtime_rules_text.insert(END, "\n" + normalized_line)
        else:
            self.runtime_rules_text.insert("1.0", normalized_line)
        self.remember_runtime_rules([normalized_line], "tela_regras_dinamicas")
        self.log_message("Regra dinamica adicionada.")
        self.write_audit_log("REGRA_ADICIONADA", normalized_line)
        self.schedule_runtime_rule_views_refresh(normalized_line, "Regra dinamica adicionada.", success_message="Regra criada com sucesso.")

    def get_runtime_rule_refresh_targets(self, *rule_lines: str) -> set[str]:
        operations: set[str] = set()
        for rule_line in rule_lines:
            try:
                parsed_rules = parse_runtime_rule_lines(rule_line)
            except Exception:
                parsed_rules = []
            for rule in parsed_rules:
                operation = normalize_operation_type(str(rule.get("operation_type", "")))
                if operation in {"Entrada", "Saida"}:
                    operations.add(operation)
        return operations or {"Entrada", "Saida"}

    def schedule_runtime_rule_views_refresh(
        self,
        rule_line: str,
        message: str,
        extra_rule_line: str = "",
        success_message: str = "Regra aplicada com sucesso.",
    ) -> None:
        if self.runtime_rule_refresh_after_id is not None:
            try:
                self.root.after_cancel(self.runtime_rule_refresh_after_id)
            except Exception:
                pass
            self.runtime_rule_refresh_after_id = None
        targets = self.get_runtime_rule_refresh_targets(rule_line, extra_rule_line)

        def run_refresh() -> None:
            self.runtime_rule_refresh_after_id = None
            total_steps = len(targets) + 1
            progress_dialog, progress_message_var, progress_percent_var = self.open_progress_dialog(
                "Atualizando Regra Dinamica",
                message,
            )
            try:
                self.update_progress_dialog(progress_dialog, progress_message_var, progress_percent_var, 0, total_steps, "Preparando atualizacao das telas...")
                current_step = 0
                if "Entrada" in targets:
                    current_step += 1
                    rows = len(self.consult_comparison_rows)
                    self.update_progress_dialog(progress_dialog, progress_message_var, progress_percent_var, current_step, total_steps, f"Atualizando consulta de entradas ({rows} linha(s))...")
                    self.refresh_consultation_tree()
                if "Saida" in targets:
                    current_step += 1
                    rows = len(self.sales_consult_comparison_rows)
                    self.update_progress_dialog(progress_dialog, progress_message_var, progress_percent_var, current_step, total_steps, f"Atualizando consulta de saidas ({rows} linha(s))...")
                    self.refresh_sales_consultation_tree()
                self.update_progress_dialog(progress_dialog, progress_message_var, progress_percent_var, total_steps, total_steps, "Regra aplicada nas telas.")
                self.log_message("Telas atualizadas apos regra dinamica.")
                self.close_progress_dialog(progress_dialog)
                messagebox.showinfo("Regra Dinamica", success_message, parent=self.root)
            except Exception as exc:
                self.close_progress_dialog(progress_dialog)
                error_message = f"Falha ao atualizar as telas apos a regra: {exc}"
                self.log_message(error_message)
                messagebox.showerror("Regra Dinamica", error_message, parent=self.root)

        self.runtime_rule_refresh_after_id = self.root.after(80, run_refresh)

    def replace_runtime_rule_line(self, old_rule_line: str, new_rule_line: str) -> None:
        old_normalized = str(old_rule_line or "").strip()
        new_normalized = str(new_rule_line or "").strip()
        if not old_normalized or not new_normalized:
            return
        lines = self.runtime_rules_text.get("1.0", END).splitlines()
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
        while updated_lines and not updated_lines[-1].strip():
            updated_lines.pop()
        self.runtime_rules_text.delete("1.0", END)
        self.runtime_rules_text.insert("1.0", "\n".join(updated_lines))
        self.remember_runtime_rules([new_normalized], "edicao_regra_dinamica")
        self.log_message("Regra dinamica editada.")
        self.write_audit_log("REGRA_EDITADA", f"antes={old_normalized}; depois={new_normalized}")
        self.schedule_runtime_rule_views_refresh(new_normalized, "Regra dinamica editada.", old_normalized, "Regra editada com sucesso.")

    def remove_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        if not messagebox.askyesno(
            "Remover Regra Dinamica",
            "Deseja remover esta regra dinamica?\n\n"
            f"{normalized_line}",
        ):
            return

        lines = self.runtime_rules_text.get("1.0", END).splitlines()
        updated_lines: list[str] = []
        removed = False
        for line in lines:
            if not removed and line.strip() == normalized_line:
                removed = True
                continue
            updated_lines.append(line)
        while updated_lines and not updated_lines[-1].strip():
            updated_lines.pop()
        self.runtime_rules_text.delete("1.0", END)
        self.runtime_rules_text.insert("1.0", "\n".join(updated_lines))
        self.log_message("Regra dinamica removida.")
        self.write_audit_log("REGRA_REMOVIDA", normalized_line)
        self.schedule_runtime_rule_views_refresh(normalized_line, "Regra dinamica removida.", success_message="Regra removida com sucesso.")

    def view_runtime_rule_line(self, rule_line: str) -> None:
        messagebox.showinfo("Regra Dinamica", str(rule_line or "").strip() or "Regra nao encontrada.")

    def build_runtime_rule_prefill_from_line(self, rule_line: str) -> dict[str, object]:
        parsed_rules = parse_runtime_rule_lines(rule_line)
        if not parsed_rules:
            return {}
        rule = parsed_rules[0]

        def decimal_text(key: str) -> str:
            value = rule.get(key, "")
            if isinstance(value, Decimal):
                return format_decimal_sped(value)
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

    def edit_runtime_rule_line(self, rule_line: str) -> None:
        normalized_line = str(rule_line or "").strip()
        if not normalized_line:
            return
        try:
            prefill = self.build_runtime_rule_prefill_from_line(normalized_line)
        except Exception as exc:
            messagebox.showerror("Editar regra", f"Nao foi possivel carregar a regra:\n{exc}")
            return
        self.open_runtime_rule_builder(prefill, replace_rule_line=normalized_line)

    def open_rule_export_dialog(self) -> None:
        dialog = Toplevel(self.root)
        dialog.title("Exportar Regras para Word")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 760, 520, 700, 460, margin_y=180)

        container = ttk.Frame(dialog, padding=16)
        container.pack(fill=BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text="Selecione os perfis que devem entrar no relatorio. Voce tambem pode incluir as regras dinamicas que estao na tela neste momento.",
            wraplength=700,
        ).grid(row=0, column=0, sticky="w", pady=(0, 12))

        profile_box = ttk.LabelFrame(container, text="Perfis disponiveis", padding=10)
        profile_box.grid(row=1, column=0, sticky="nsew")
        profile_box.columnconfigure(0, weight=1)
        profile_vars: dict[str, BooleanVar] = {}
        for row_index, profile_name in enumerate(sorted(DEFAULT_ICMS_RULE_PROFILES.keys()), start=0):
            profile_var = BooleanVar(value=(profile_name == self.adjustment_mode_var.get().strip()))
            profile_vars[profile_name] = profile_var
            ttk.Checkbutton(profile_box, text=profile_name, variable=profile_var).grid(
                row=row_index,
                column=0,
                sticky="w",
                pady=2,
            )

        include_runtime_var = BooleanVar(value=bool(self.runtime_rules_text.get("1.0", END).strip()))
        ttk.Checkbutton(
            container,
            text="Incluir as regras dinamicas digitadas atualmente",
            variable=include_runtime_var,
        ).grid(row=2, column=0, sticky="w", pady=(12, 0))

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        ttk.Button(
            actions,
            text="Marcar Todos",
            style="Secondary.TButton",
            command=lambda: [var.set(True) for var in profile_vars.values()],
        ).pack(side=LEFT)
        ttk.Button(
            actions,
            text="Limpar Perfis",
            style="Secondary.TButton",
            command=lambda: [var.set(False) for var in profile_vars.values()],
        ).pack(side=LEFT, padx=(8, 0))

        def export_selected() -> None:
            selected_profiles = [name for name, var in profile_vars.items() if var.get()]
            include_runtime = bool(include_runtime_var.get())
            if not selected_profiles and not include_runtime:
                messagebox.showwarning("Exportar regras", "Selecione ao menos um perfil ou marque as regras dinamicas.")
                return
            try:
                output_path = self.export_rules_report(selected_profiles, include_runtime)
            except Exception as exc:
                self.log_message(f"Falha ao exportar regras: {exc}")
                messagebox.showerror("Exportar regras", str(exc))
                return
            if output_path is None:
                return
            dialog.destroy()
            self.log_message(f"Relatorio de regras gerado: {output_path}")
            messagebox.showinfo("Exportar regras", f"Relatorio gerado com sucesso:\n{output_path}")

        ttk.Button(actions, text="Exportar", style="Primary.TButton", command=export_selected).pack(side=RIGHT)

    def export_rules_report(self, selected_profiles: list[str], include_runtime_rules: bool) -> Path | None:
        sections: list[dict[str, object]] = []

        for profile_name in selected_profiles:
            rules = DEFAULT_ICMS_RULE_PROFILES.get(profile_name, [])
            section_lines: list[str] = []
            for rule in rules:
                section_lines.extend(build_rule_report_entries(rule))
            sections.append({"title": f"Perfil: {profile_name}", "lines": section_lines})

        if include_runtime_rules:
            raw_runtime_text = self.runtime_rules_text.get("1.0", END).strip()
            if not raw_runtime_text:
                raise ValueError("Nao ha regras dinamicas preenchidas para exportar.")
            parsed_runtime_rules = parse_runtime_rule_lines(raw_runtime_text)
            section_lines: list[str] = []
            for rule in parsed_runtime_rules:
                section_lines.extend(build_rule_report_entries(rule))
            sections.append({"title": "Regras dinamicas atuais", "lines": section_lines})

        if not sections:
            raise ValueError("Nao foi encontrada nenhuma regra para exportar.")

        default_name = f"relatorio_regras_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        selected = filedialog.asksaveasfilename(
            title="Salvar relatorio de regras",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Documento Word", "*.docx")],
        )
        if not selected:
            return None

        output_path = Path(selected)
        write_rules_report_docx(output_path, "Relatorio de Regras de Recalculo", sections)
        return output_path

    def open_runtime_rule_builder(
        self,
        prefill: dict[str, object] | None = None,
        on_rule_added: Callable[[str], None] | None = None,
        replace_rule_line: str | None = None,
    ) -> None:
        prefill = prefill or {}
        dialog = Toplevel(self.root)
        dialog.title("Editar Regra Dinamica" if replace_rule_line else "Adicionar Regra Dinamica")
        dialog.transient(self.root)
        dialog.grab_set()
        self.bind_escape_to_close(dialog)
        self.set_dialog_screen_geometry(dialog, 1120, 700, 980, 620, margin_y=160)

        container = ttk.Frame(dialog, padding=16, style="Card.TFrame")
        container.pack(fill=BOTH, expand=True)

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(
            container,
            text="Monte a regra pelos campos abaixo. O IPI nao sera alterado por essa regra.",
            wraplength=1040,
            style="Title.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(
            container,
            text="Preencha apenas o necessario. Campos vazios sao ignorados.",
            wraplength=1040,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        operation_var = StringVar(value=str(prefill.get("operation", "entrada") or "entrada"))
        document_number_var = StringVar(value=str(prefill.get("document", "") or ""))
        document_tax_id_var = StringVar(value=str(prefill.get("tax_id", "") or ""))
        cst_var = StringVar(value=str(prefill.get("cst", "") or ""))
        cfop_var = StringVar(value=str(prefill.get("cfop", "") or ""))
        rate_var = StringVar(value=str(prefill.get("rate", "") or ""))
        codes_var = StringVar(value=str(prefill.get("codes", "") or ""))
        new_cst_var = StringVar(value=str(prefill.get("new_cst", "") or ""))
        new_cfop_var = StringVar(value=str(prefill.get("new_cfop", "") or ""))
        force_rate_var = StringVar(value=str(prefill.get("force_rate", "") or ""))
        set_base_var = StringVar(value=str(prefill.get("set_base", "") or ""))
        set_icms_value_var = StringVar(value=str(prefill.get("set_icms_value", "") or ""))
        base_reduction_percent_var = StringVar(value=str(prefill.get("base_reduction_percent", "") or ""))
        zero_icms_var = BooleanVar(value=bool(prefill.get("zero_icms", False)))
        use_sale_value_base_var = BooleanVar(value=bool(prefill.get("use_sale_value_base", False)))
        recalculate_icms_value_var = BooleanVar(value=bool(prefill.get("recalculate_icms_value", False)))
        recalculate_reduced_base_var = BooleanVar(value=bool(prefill.get("recalculate_reduced_base", False)))

        left_box = ttk.LabelFrame(container, text="Filtros da Regra", padding=12)
        left_box.grid(row=2, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))
        right_box = ttk.LabelFrame(container, text="Acoes de Recalculo", padding=12)
        right_box.grid(row=2, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))
        options_box = ttk.LabelFrame(container, text="Opcoes da Regra", padding=12)
        options_box.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        left_box.columnconfigure(0, weight=1)
        left_box.columnconfigure(1, weight=1)
        left_box.columnconfigure(2, weight=1)
        right_box.columnconfigure(0, weight=1)
        right_box.columnconfigure(1, weight=1)
        right_box.columnconfigure(2, weight=1)
        options_box.columnconfigure(0, weight=1)
        options_box.columnconfigure(1, weight=1)
        options_box.columnconfigure(2, weight=1)

        def add_entry(parent: ttk.LabelFrame, row: int, column: int, label: str, variable: StringVar) -> None:
            field = ttk.Frame(parent)
            field.grid(row=row, column=column, sticky="ew", padx=6, pady=4)
            field.columnconfigure(0, weight=1)
            ttk.Label(field, text=label).grid(row=0, column=0, sticky="w")
            ttk.Entry(field, textvariable=variable).grid(row=1, column=0, sticky="ew", pady=(3, 0))

        type_field = ttk.Frame(left_box)
        type_field.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        type_field.columnconfigure(0, weight=1)
        ttk.Label(type_field, text="Tipo de operacao").grid(row=0, column=0, sticky="w")
        ttk.Combobox(type_field, textvariable=operation_var, values=("entrada", "saida"), state="readonly").grid(row=1, column=0, sticky="ew", pady=(3, 0))

        add_entry(left_box, 0, 1, "CST", cst_var)
        add_entry(left_box, 0, 2, "CFOP", cfop_var)
        add_entry(left_box, 1, 0, "Aliquota", rate_var)
        add_entry(left_box, 1, 1, "Codigo(s)", codes_var)
        add_entry(left_box, 1, 2, "Documento", document_number_var)
        add_entry(left_box, 2, 0, "CNPJ", document_tax_id_var)

        add_entry(right_box, 0, 0, "Novo CST", new_cst_var)
        add_entry(right_box, 0, 1, "Novo CFOP", new_cfop_var)
        add_entry(right_box, 0, 2, "Nova aliquota", force_rate_var)
        add_entry(right_box, 1, 0, "Definir base ICMS", set_base_var)
        add_entry(right_box, 1, 1, "Definir valor ICMS", set_icms_value_var)
        add_entry(right_box, 1, 2, "% reducao de base", base_reduction_percent_var)

        ttk.Checkbutton(options_box, text="Zerar ICMS/Base/Valor", variable=zero_icms_var).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(options_box, text="Usar valor da operacao como base ICMS", variable=use_sale_value_base_var).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(options_box, text="Recalcular valor do ICMS", variable=recalculate_icms_value_var).grid(row=0, column=2, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(options_box, text="Recalcular base de calculo da reducao", variable=recalculate_reduced_base_var).grid(row=1, column=0, sticky="w", padx=6, pady=4)

        def sync_reduced_base_option(*_args: object) -> None:
            if recalculate_reduced_base_var.get():
                recalculate_icms_value_var.set(True)

        recalculate_reduced_base_var.trace_add("write", sync_reduced_base_option)

        def sync_reduction_percent(*_args: object) -> None:
            if base_reduction_percent_var.get().strip():
                recalculate_reduced_base_var.set(True)
                recalculate_icms_value_var.set(True)

        base_reduction_percent_var.trace_add("write", sync_reduction_percent)

        ttk.Label(
            container,
            text="Exemplo: tipo=entrada; cnpj=12345678000199; cst=000; cfop=1102; novo_cfop=1101",
            wraplength=1040,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 10))

        footer = ttk.Frame(container)
        footer.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        def save_runtime_rule() -> None:
            fragments = [f"tipo={operation_var.get().strip()}"]
            optional_fields = [
                ("documento", document_number_var.get()),
                ("cnpj", document_tax_id_var.get()),
                ("cst", cst_var.get()),
                ("cfop", cfop_var.get()),
                ("aliquota", rate_var.get()),
                ("codigos", codes_var.get()),
                ("novo_cst", new_cst_var.get()),
                ("novo_cfop", new_cfop_var.get()),
                ("nova_aliquota", force_rate_var.get()),
                ("definir_base_icms", set_base_var.get()),
                ("definir_valor_icms", set_icms_value_var.get()),
                ("percentual_reducao_base", base_reduction_percent_var.get()),
            ]
            for key, value in optional_fields:
                text = str(value or "").strip()
                if text:
                    fragments.append(f"{key}={text}")

            if zero_icms_var.get():
                fragments.append("zerar_icms=sim")
            if use_sale_value_base_var.get():
                fragments.append("usar_valor_operacao_como_base=sim")
            if recalculate_icms_value_var.get():
                fragments.append("recalcular_valor_icms=sim")
            if recalculate_reduced_base_var.get():
                fragments.append("recalcular_base_reducao=sim")

            if len(fragments) <= 1:
                messagebox.showwarning("Regra incompleta", "Informe pelo menos um filtro ou uma acao para montar a regra.", parent=dialog)
                return
            if recalculate_reduced_base_var.get() and not base_reduction_percent_var.get().strip():
                messagebox.showwarning("Regra incompleta", "Informe o percentual de reducao de base.", parent=dialog)
                return
            if base_reduction_percent_var.get().strip() and not recalculate_reduced_base_var.get():
                messagebox.showwarning("Regra incompleta", "Marque a opcao de recalcular base de calculo da reducao.", parent=dialog)
                return

            action_keys = {
                "novo_cst", "novo_cfop", "nova_aliquota", "definir_base_icms", "definir_valor_icms",
                "percentual_reducao_base", "zerar_icms=sim", "usar_valor_operacao_como_base=sim",
                "recalcular_valor_icms=sim", "recalcular_base_reducao=sim",
            }
            if not any(fragment.split("=", 1)[0] in action_keys or fragment in action_keys for fragment in fragments[1:]):
                messagebox.showwarning("Regra sem acao", "Informe ao menos uma acao de recalculo para a regra.", parent=dialog)
                return

            rule_line = "; ".join(fragments)
            try:
                parse_runtime_rule_lines(rule_line)
            except Exception as exc:
                messagebox.showerror("Regra invalida", str(exc), parent=dialog)
                return

            if replace_rule_line:
                self.replace_runtime_rule_line(replace_rule_line, rule_line)
            else:
                self.append_runtime_rule_line(rule_line)
            if on_rule_added is not None:
                on_rule_added(rule_line)
            dialog.destroy()

        ttk.Button(footer, text="Salvar" if replace_rule_line else "Adicionar", style="Primary.TButton", command=save_runtime_rule).pack(side=LEFT)
        ttk.Button(footer, text="Cancelar", style="Secondary.TButton", command=dialog.destroy).pack(side=LEFT, padx=(8, 0))
        dialog.wait_window()

    def generate_excel(self) -> None:
        sped_source = self.sped_path_var.get().strip()
        excel_source = self.excel_path_var.get().strip()
        if not sped_source and not excel_source:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione o arquivo SPED Fiscal, a planilha Excel, ou ambos.")
            return

        sped_path = Path(sped_source) if sped_source else None
        excel_path = Path(excel_source) if excel_source else None
        if sped_path and not sped_path.exists():
            messagebox.showerror("Arquivo nao encontrado", "O arquivo SPED selecionado nao existe mais.")
            return
        if excel_path and not excel_path.exists():
            messagebox.showerror("Arquivo nao encontrado", "A planilha selecionada nao existe mais.")
            return

        base_name = excel_path.stem if excel_path else sped_path.stem if sped_path else "saida"
        default_name = f"{base_name}_resumo_icms.xlsx"
        output = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not output:
            return

        output_path = Path(output)
        filter_descriptions = [
            line.strip()
            for line in self.filter_text.get("1.0", END).splitlines()
            if line.strip()
        ]
        cst_filter_values = parse_filter_values(self.cst_filter_var.get())
        cfop_filter_values = parse_filter_values(self.cfop_filter_var.get())
        icms_rate_override_text = self.icms_rate_var.get().strip()
        new_cst = parse_replacement_value(self.new_cst_var.get())
        new_cfop = parse_replacement_value(self.new_cfop_var.get())
        zero_icms_cfops = parse_filter_values(self.zero_icms_cfop_var.get())
        xml_ncm_filters = parse_filter_values(self.nfce_ncm_filter_var.get())
        selected_rule_profile = self.adjustment_mode_var.get().strip()
        if selected_rule_profile == "Filtros":
            selected_rule_profile = ""

        try:
            runtime_rules = parse_runtime_rule_lines(self.runtime_rules_text.get("1.0", END))
            runtime_rule_lines = [line.strip() for line in self.runtime_rules_text.get("1.0", END).splitlines() if line.strip()]
            sped_data = None
            excel_data = None

            if sped_path:
                self.log_message("Lendo registros do SPED Fiscal...")
                sped_data = read_sped_file(sped_path)

            if excel_path:
                self.log_message("Importando a planilha de produtos...")
                excel_data = read_detailed_product_excel(excel_path)

            product_rows, sales_rows, detailed_sales, c190_rows, c190_product_rows = combine_imported_data(sped_data, excel_data)

            if not product_rows and not sales_rows and not c190_rows:
                raise ValueError("Nenhum dado valido foi encontrado no arquivo informado.")

            self.log_message("Aplicando filtro de descricoes...")
            if selected_rule_profile or runtime_rules:
                filtered_detailed_sales = [
                    item for item in detailed_sales
                    if has_configured_icms_rule(selected_rule_profile, runtime_rules, item)
                ]
                filtered_rows = filter_sales(filtered_detailed_sales, [], set(), set())
            else:
                filtered_detailed_sales = filter_detailed_sales(
                    detailed_sales,
                    filter_descriptions,
                    cst_filter_values,
                    cfop_filter_values,
                )
                filtered_rows = filter_sales(
                    detailed_sales,
                    filter_descriptions,
                    cst_filter_values,
                    cfop_filter_values,
                )
            filtered_c190_rows, filtered_c190_product_rows = build_c190_rows_from_details(
                filtered_detailed_sales,
                selected_rule_profile,
                runtime_rules,
            )
            icms_rate_override = parse_rate(icms_rate_override_text) if icms_rate_override_text else None
            recalculated_detailed_sales = rebuild_detailed_sales_with_override(
                filtered_detailed_sales,
                [],
                set(),
                set(),
                icms_rate_override,
                new_cst,
                new_cfop,
                zero_icms_cfops,
                selected_rule_profile,
                runtime_rules,
            )
            _, recalculated_c190_product_rows = build_c190_rows_from_details(
                recalculated_detailed_sales,
                selected_rule_profile,
                runtime_rules,
            )

            self.log_message("Gerando planilha Excel...")
            write_excel(
                output_path,
                product_rows,
                sales_rows,
                detailed_sales,
                filtered_rows,
                filtered_c190_rows,
                filtered_c190_product_rows,
                recalculated_c190_product_rows,
            )
            self.remember_runtime_rules(runtime_rule_lines, "excel")
            self.log_message(f"Excel gerado com sucesso em: {output_path}")
            self.show_front_message("showinfo", "Processo concluido", f"Arquivo gerado com sucesso:\n{output_path}")
        except Exception as exc:
            self.log_message(f"Falha ao processar o arquivo: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))

    def get_selected_reprocess_sped_paths(self) -> list[Path]:
        candidate_paths: list[Path] = []
        primary_source = self.sped_path_var.get().strip()
        if primary_source:
            candidate_paths.append(Path(primary_source))
        candidate_paths.extend(parse_selected_paths(self.consult_sped_paths_var.get()))
        candidate_paths.extend(parse_selected_paths(self.sales_consult_sped_paths_var.get()))

        unique_paths: list[Path] = []
        seen_paths: set[Path] = set()
        for path in candidate_paths:
            resolved = path.resolve() if path.exists() else path
            if resolved in seen_paths:
                continue
            seen_paths.add(resolved)
            unique_paths.append(path)
        return unique_paths

    def run_adjusted_sped_generation(
        self,
        sped_path: Path,
        output_path: Path,
        excel_path: Path | None,
        xml_paths: list[Path],
        filter_descriptions: list[str],
        cst_filter_values: set[str],
        cfop_filter_values: set[str],
        icms_rate_override_text: str,
        new_cst: str,
        new_cfop: str,
        zero_icms_cfops: set[str],
        xml_ncm_filters: set[str],
        selected_rule_profile: str,
        adjusted_sped_scope: str,
    ) -> list[str]:
        runtime_rules = parse_runtime_rule_lines(self.runtime_rules_text.get("1.0", END))
        runtime_rule_lines = [line.strip() for line in self.runtime_rules_text.get("1.0", END).splitlines() if line.strip()]
        self.log_message(f"Lendo registros do SPED Fiscal: {sped_path.name}")
        self.write_audit_log(
            "REPROCESSAMENTO_INICIO",
            (
                f"sped_origem={sped_path}; sped_saida={output_path}; escopo={adjusted_sped_scope}; "
                f"planilha={excel_path or 'Nao informado'}; xmls={self.format_audit_paths(xml_paths)}; "
                f"perfil_regra={selected_rule_profile or 'Filtros'}; regras_dinamicas={len(runtime_rule_lines)}"
            ),
        )
        sped_data = read_sped_file(sped_path)
        excel_data = None
        if excel_path:
            self.log_message("Importando a planilha de produtos...")
            excel_data = read_detailed_product_excel(excel_path)
        if xml_paths:
            if xml_ncm_filters:
                self.log_message("Aplicando XMLs por chave fiscal e filtro de NCM no SPED ajustado...")
            else:
                self.log_message("Aplicando ST a partir dos XMLs de NF-e...")

        _, _, detailed_sales, _, _ = combine_imported_data(sped_data, excel_data)
        original_summary = self.summarize_audit_detail_rows(detailed_sales)
        icms_rate_override = parse_rate(icms_rate_override_text) if icms_rate_override_text else None

        self.log_message("Recalculando C190 por documento...")
        rebuilt_detailed_sales = write_adjusted_sped(
            output_path,
            sped_path,
            detailed_sales,
            filter_descriptions,
            cst_filter_values,
            cfop_filter_values,
            icms_rate_override,
            new_cst,
            new_cfop,
            zero_icms_cfops,
            adjusted_sped_scope,
            selected_rule_profile,
            xml_paths,
            xml_ncm_filters,
            runtime_rules,
        )
        adjusted_summary = self.summarize_audit_detail_rows(rebuilt_detailed_sales)
        self.write_audit_log(
            "REPROCESSAMENTO_VALORES",
            (
                f"sped_origem={sped_path}; antes=({self.format_audit_summary(original_summary)}); "
                f"depois=({self.format_audit_summary(adjusted_summary)})"
            ),
        )
        self.remember_runtime_rules(runtime_rule_lines, "sped_ajustado")
        cst_061_output = output_path.with_name(f"{output_path.stem}_entradas_cst_061.xlsx")
        write_cst_061_excel(cst_061_output, rebuilt_detailed_sales)
        self.log_message(f"Planilha CST 061 gerada em: {cst_061_output}")
        extra_outputs = [str(cst_061_output)]
        missing_xml_output = output_path.with_name(f"{output_path.stem}_st_1401_1403_sem_xml.xlsx")
        if missing_xml_output.exists():
            self.log_message(f"Planilha ST sem XML gerada em: {missing_xml_output}")
            extra_outputs.append(str(missing_xml_output))
        if selected_rule_profile == "CASA DE PAES - ENTRADAS":
            cfop_1252_1253_output = output_path.with_name(f"{output_path.stem}_entradas_cfop_1252_1253.xlsx")
            write_cfop_1252_1253_excel(cfop_1252_1253_output, rebuilt_detailed_sales)
            self.log_message(f"Planilha CFOP 1252/1253 gerada em: {cfop_1252_1253_output}")
            extra_outputs.append(str(cfop_1252_1253_output))
        self.log_message(f"SPED ajustado gerado com sucesso em: {output_path}")
        self.write_audit_log("REPROCESSAMENTO_FIM", f"sped_origem={sped_path}; arquivo_reprocessado={output_path}; extras={self.format_audit_paths(extra_outputs)}")
        return [str(output_path), *extra_outputs]

    def generate_adjusted_sped(self) -> None:
        selected_sped_paths = self.get_selected_reprocess_sped_paths()
        sped_source = str(selected_sped_paths[0]) if selected_sped_paths else self.sped_path_var.get().strip()
        excel_source = self.excel_path_var.get().strip()
        xml_source = self.xml_path_var.get().strip()
        if not selected_sped_paths and not sped_source:
            messagebox.showwarning("Arquivo obrigatorio", "Selecione ao menos um arquivo SPED para gerar o SPED ajustado.")
            return

        sped_paths = selected_sped_paths if selected_sped_paths else [Path(sped_source)]
        excel_path = Path(excel_source) if excel_source else None
        xml_paths = parse_selected_paths(xml_source)
        missing_sped_paths = [str(path) for path in sped_paths if not path.exists()]
        if missing_sped_paths:
            messagebox.showerror("Arquivo nao encontrado", "Os seguintes SPEDs nao existem mais:\n" + "\n".join(missing_sped_paths))
            return
        if excel_path and not excel_path.exists():
            messagebox.showerror("Arquivo nao encontrado", "A planilha selecionada nao existe mais.")
            return
        missing_xml_paths = [str(path) for path in xml_paths if not path.exists()]
        if missing_xml_paths:
            messagebox.showerror("Arquivo nao encontrado", "Os seguintes XMLs/pastas nao existem mais:\n" + "\n".join(missing_xml_paths))
            return

        output_paths: list[Path] = []
        if len(sped_paths) == 1:
            selected_output = filedialog.asksaveasfilename(
                title="Salvar SPED ajustado",
                defaultextension=".txt",
                initialfile=f"{sped_paths[0].stem}_ajustado.txt",
                filetypes=[("Arquivo SPED", "*.txt *.sped"), ("Todos os arquivos", "*.*")],
            )
            if not selected_output:
                return
            output_paths = [Path(selected_output)]
        else:
            selected_folder = filedialog.askdirectory(title="Selecione a pasta para salvar os SPEDs ajustados")
            if not selected_folder:
                return
            output_dir = Path(selected_folder)
            output_paths = [output_dir / f"{path.stem}_ajustado{path.suffix or '.txt'}" for path in sped_paths]

        filter_descriptions = [
            line.strip()
            for line in self.filter_text.get("1.0", END).splitlines()
            if line.strip()
        ]
        cst_filter_values = parse_filter_values(self.cst_filter_var.get())
        cfop_filter_values = parse_filter_values(self.cfop_filter_var.get())
        icms_rate_override_text = self.icms_rate_var.get().strip()
        new_cst = parse_replacement_value(self.new_cst_var.get())
        new_cfop = parse_replacement_value(self.new_cfop_var.get())
        zero_icms_cfops = parse_filter_values(self.zero_icms_cfop_var.get())
        xml_ncm_filters = parse_filter_values(self.nfce_ncm_filter_var.get())
        selected_rule_profile = self.adjustment_mode_var.get().strip()
        if selected_rule_profile == "Filtros":
            selected_rule_profile = ""
        adjusted_sped_scope = self.adjusted_sped_scope_var.get().strip() or "Ambos"

        try:
            generated_files: list[str] = []
            for sped_path, output_path in zip(sped_paths, output_paths):
                generated_files.extend(
                    self.run_adjusted_sped_generation(
                        sped_path,
                        output_path,
                        excel_path,
                        xml_paths,
                        filter_descriptions,
                        cst_filter_values,
                        cfop_filter_values,
                        icms_rate_override_text,
                        new_cst,
                        new_cfop,
                        zero_icms_cfops,
                        xml_ncm_filters,
                        selected_rule_profile,
                        adjusted_sped_scope,
                    )
                )
            generated_files_text = "\n".join(generated_files)
            self.show_front_message("showinfo", "Processo concluido", f"Arquivos gerados com sucesso:\n{generated_files_text}")
        except Exception as exc:
            self.log_message(f"Falha ao gerar SPED ajustado: {exc}")
            messagebox.showerror("Erro no processamento", str(exc))
