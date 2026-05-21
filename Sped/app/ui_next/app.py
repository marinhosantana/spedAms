from __future__ import annotations

import datetime as dt
import csv
import os
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import BOTH, BooleanVar, END, LEFT, RIGHT, StringVar, Tk, Toplevel, filedialog, messagebox, ttk

from app.config import MYSQL_DEFAULT_CONFIG
from app.parsers.sped_fiscal_parser import read_sped_file
from app.repositories.mysql_cadastro import MysqlCadastroRepository
from app.services.app_paths import (
    get_application_base_dir,
    get_application_environment,
    get_environment_config_path,
    get_project_root_dir,
    get_sped_archive_storage_dir,
)
from app.services.path_selection import (
    append_unique_paths,
    collapse_xml_selection_paths,
    format_selected_paths,
    limit_selected_paths,
    parse_selected_paths,
)
from app.services.period_comparisons import build_entry_period_comparison_rows, build_sale_period_comparison_rows
from app.services.sped_archive import archive_original_sped_file
from app.ui_next.theme import COLORS, configure_theme


class NextSpedApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.environment = get_application_environment()
        self.base_dir = get_application_base_dir(__file__)
        self.project_root_dir = get_project_root_dir(self.base_dir)
        self.storage_dir = get_sped_archive_storage_dir(self.project_root_dir)
        self.mysql_config_path = get_environment_config_path(self.project_root_dir / "app" / "ui", "mysql_config", self.environment)
        self.mysql_repo = MysqlCadastroRepository(
            self.mysql_config_path,
            self.project_root_dir / "mysql_schema.sql",
            self.get_mysql_default_config(),
        )
        self.page_frames: dict[str, ttk.Frame] = {}
        self.archive_profile_rows: dict[int, dict[str, object]] = {}
        self.archive_file_rows: dict[int, dict[str, object]] = {}
        self.status_var = StringVar(value="Pronto.")
        self.entry_sped_paths_var = StringVar()
        self.entry_xml_paths_var = StringVar()
        self.sale_sped_paths_var = StringVar()
        self.sale_xml_paths_var = StringVar()
        self.entry_summary_var = StringVar(value="Entradas: nenhum processamento executado.")
        self.sale_summary_var = StringVar(value="Saidas: nenhum processamento executado.")
        self.entry_consult_cst_var = StringVar()
        self.entry_consult_cfop_var = StringVar()
        self.entry_consult_status_var = StringVar(value="Todos")
        self.entry_consult_search_var = StringVar()
        self.entry_consult_footer_var = StringVar(value="Linhas: 0    Produtos: 0    Valor Operacao: 0,00    Base ICMS: 0,00    % Base/Oper: 0,00    Valor ICMS: 0,00")
        self.entry_period_labels: list[str] = []
        self.entry_period_check_vars: dict[str, BooleanVar] = {}
        self.entry_consult_rows: list[dict[str, object]] = []
        self.entry_consult_filtered_rows: list[dict[str, object]] = []

        configure_theme(root)
        self.root.title(f"Revisor de SPED Next [{self.environment.upper()}]")
        self.root.minsize(1100, 720)
        self.root.state("zoomed")
        self.build_shell()
        self.show_page("dashboard")

    def get_mysql_default_config(self) -> dict[str, str]:
        config = dict(MYSQL_DEFAULT_CONFIG)
        if self.environment == "dev":
            config["database"] = "sped_icms_dev"
        return config

    def build_shell(self) -> None:
        root_frame = ttk.Frame(self.root, style="Next.TFrame")
        root_frame.pack(fill=BOTH, expand=True)
        root_frame.columnconfigure(1, weight=1)
        root_frame.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(root_frame, style="Next.Sidebar.TFrame", width=230)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        ttk.Label(sidebar, text="SPED Next", style="Next.Sidebar.TLabel").pack(anchor="w", padx=18, pady=(18, 22))
        for text, page_id in (
            ("Dashboard", "dashboard"),
            ("SPEDs Arquivados", "archives"),
            ("Consultas Fiscais", "fiscal_queries"),
            ("Comparacoes", "placeholder_compare"),
            ("Configuracoes", "settings"),
        ):
            ttk.Button(sidebar, text=text, style="Next.Sidebar.TButton", command=lambda current=page_id: self.show_page(current)).pack(fill="x", padx=8, pady=2)

        content = ttk.Frame(root_frame, style="Next.TFrame", padding=16)
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        header = ttk.Frame(content, style="Next.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        self.page_title_var = StringVar(value="")
        ttk.Label(header, textvariable=self.page_title_var, style="Next.Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=f"Ambiente: {self.environment.upper()}", style="Next.Muted.TLabel").grid(row=0, column=1, sticky="e")

        self.content_host = ttk.Frame(content, style="Next.TFrame")
        self.content_host.grid(row=1, column=0, sticky="nsew")
        self.content_host.columnconfigure(0, weight=1)
        self.content_host.rowconfigure(0, weight=1)

        ttk.Label(content, textvariable=self.status_var, style="Next.Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))

        self.page_frames["dashboard"] = self.build_dashboard_page()
        self.page_frames["archives"] = self.build_archives_page()
        self.page_frames["fiscal_queries"] = self.build_fiscal_queries_page()
        self.page_frames["entry_consultation"] = self.build_entry_consultation_page()
        self.page_frames["settings"] = self.build_settings_page()
        self.page_frames["placeholder_compare"] = self.build_placeholder_page("Comparacoes", "A nova experiencia de comparacao entrara aqui sem remover a tela antiga.")

    def show_page(self, page_id: str) -> None:
        for frame in self.page_frames.values():
            frame.grid_remove()
        frame = self.page_frames[page_id]
        frame.grid(row=0, column=0, sticky="nsew")
        titles = {
            "dashboard": "Dashboard",
            "archives": "SPEDs Arquivados",
            "fiscal_queries": "Consultas Fiscais",
            "entry_consultation": "Consulta Entradas ICMS",
            "settings": "Configuracoes",
            "placeholder_compare": "Comparacoes",
        }
        self.page_title_var.set(titles.get(page_id, "SPED Next"))
        if page_id == "dashboard":
            self.refresh_dashboard()

    def build_dashboard_page(self) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        page.columnconfigure((0, 1, 2), weight=1)
        self.dashboard_cards: dict[str, StringVar] = {}
        for index, (key, title) in enumerate(
            (
                ("profiles", "Perfis Arquivados"),
                ("files", "Arquivos SPED"),
                ("status", "Banco de Dados"),
            )
        ):
            card = ttk.Frame(page, style="Next.Panel.TFrame", padding=16)
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 10, 0))
            value_var = StringVar(value="-")
            self.dashboard_cards[key] = value_var
            ttk.Label(card, text=title, style="Next.PanelMuted.TLabel").pack(anchor="w")
            ttk.Label(card, textvariable=value_var, style="Next.Panel.TLabel", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Button(page, text="Atualizar", style="Next.Primary.TButton", command=self.refresh_dashboard).grid(row=1, column=0, sticky="w", pady=(16, 0))
        return page

    def refresh_dashboard(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            profiles = self.mysql_repo.list_sped_profiles(self.environment)
            files = self.mysql_repo.list_sped_archives(self.environment)
            self.dashboard_cards["profiles"].set(str(len(profiles)))
            self.dashboard_cards["files"].set(str(len(files)))
            self.dashboard_cards["status"].set("Conectado")
            self.status_var.set(f"Dashboard atualizado em {dt.datetime.now().strftime('%H:%M:%S')}.")
        except Exception as exc:
            self.dashboard_cards["status"].set("Indisponivel")
            self.status_var.set(f"MySQL indisponivel: {exc}")

    def build_archives_page(self) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)
        actions = ttk.Frame(page, style="Next.TFrame")
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(actions, text="Importar SPED", style="Next.Primary.TButton", command=self.import_speds).pack(side=LEFT)
        ttk.Button(actions, text="Atualizar", style="Next.TButton", command=self.refresh_archives).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions, text="Abrir Pasta", style="Next.TButton", command=self.open_selected_archive_folder).pack(side=LEFT, padx=(8, 0))

        split = ttk.Frame(page, style="Next.TFrame")
        split.grid(row=1, column=0, sticky="nsew")
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=2)
        split.rowconfigure(0, weight=1)

        self.profile_box, self.profile_tree = self.create_tree(split, ("id", "empresa", "cnpj", "periodo", "arquivos"), {"id": 55, "empresa": 300, "cnpj": 140, "periodo": 165, "arquivos": 80})
        self.profile_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.profile_tree.bind("<<TreeviewSelect>>", lambda _event: self.refresh_archive_files())
        self.file_box, self.file_tree = self.create_tree(split, ("id", "tipo", "arquivo", "periodo", "0200", "docs", "c170", "c190", "hash"), {"id": 55, "tipo": 110, "arquivo": 250, "periodo": 165, "0200": 65, "docs": 65, "c170": 65, "c190": 65, "hash": 120})
        self.file_box.grid(row=0, column=1, sticky="nsew")
        return page

    def create_tree(self, parent: ttk.Frame, columns: tuple[str, ...], widths: dict[str, int]) -> tuple[ttk.Frame, ttk.Treeview]:
        box = ttk.Frame(parent, style="Next.Panel.TFrame", padding=8)
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        tree = ttk.Treeview(box, columns=columns, show="headings", selectmode="browse")
        for column in columns:
            tree.heading(column, text=column.replace("_", " ").upper())
            tree.column(column, width=widths.get(column, 100), anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll.set, xscrollcommand=scroll_x.set)
        return box, tree

    def refresh_archives(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            profiles = self.mysql_repo.list_sped_profiles(self.environment)
            self.archive_profile_rows = {int(row["id"]): row for row in profiles}
            tree = self.profile_tree
            tree.delete(*tree.get_children())
            for row in profiles:
                profile_id = int(row["id"])
                tree.insert(
                    "",
                    END,
                    iid=str(profile_id),
                    values=(
                        profile_id,
                        row.get("empresa_nome_sped", "") or row.get("empresa_nome", "") or row.get("nome", ""),
                        row.get("empresa_cnpj_sped", "") or row.get("empresa_cnpj", ""),
                        self.format_period(row.get("periodo_inicio"), row.get("periodo_fim")),
                        row.get("total_arquivos", 0),
                    ),
                )
            if profiles:
                tree.selection_set(str(int(profiles[0]["id"])))
                self.refresh_archive_files()
            self.status_var.set(f"{len(profiles)} perfil(is) carregado(s).")
        except Exception as exc:
            self.status_var.set(f"Falha ao carregar perfis: {exc}")
            messagebox.showerror("SPED Next", str(exc))

    def refresh_archive_files(self) -> None:
        profile_tree = self.profile_tree
        file_tree = self.file_tree
        selected = profile_tree.selection()
        if not selected:
            return
        archives = self.mysql_repo.list_sped_archives(self.environment, int(selected[0]))
        self.archive_file_rows = {int(row["id"]): row for row in archives}
        file_tree.delete(*file_tree.get_children())
        for row in archives:
            archive_id = int(row["id"])
            file_tree.insert(
                "",
                END,
                iid=str(archive_id),
                values=(
                    archive_id,
                    row.get("tipo_sped", ""),
                    row.get("arquivo_nome_original", ""),
                    self.format_period(row.get("periodo_inicio"), row.get("periodo_fim")),
                    row.get("total_produtos", 0),
                    row.get("total_documentos", 0),
                    row.get("total_itens", 0),
                    row.get("total_c190", 0),
                    str(row.get("arquivo_hash_sha256", ""))[:12],
                ),
            )

    def import_speds(self) -> None:
        selected_files = filedialog.askopenfilenames(
            title="Importar SPEDs",
            filetypes=[("Arquivos SPED", "*.txt *.sped *.efd"), ("Todos os arquivos", "*.*")],
        )
        imported = 0
        for selected in selected_files:
            if self.register_sped(Path(selected)):
                imported += 1
        self.refresh_archives()
        self.status_var.set(f"{imported} de {len(selected_files)} arquivo(s) importado(s).")

    def register_sped(self, sped_path: Path) -> int | None:
        metadata = archive_original_sped_file(sped_path, self.storage_dir, self.environment)
        self.mysql_repo.ensure_schema()
        existing = self.mysql_repo.get_sped_archive_by_hash(self.environment, metadata.file_hash_sha256)
        if existing:
            return int(existing["id"])
        company_id = self.mysql_repo.find_company_id_by_tax_id(metadata.company_tax_id)
        profile_id = self.mysql_repo.ensure_sped_profile(
            self.environment,
            metadata.default_profile_name,
            company_id,
            "Perfil criado pela nova interface.",
            metadata.company_name,
            metadata.company_tax_id,
        )
        archive_id = self.mysql_repo.save_sped_archive(
            {
                "perfil_id": profile_id,
                "empresa_id": company_id,
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
                "observacao": f"Importado pela nova interface. Empresa no SPED: {metadata.company_name}",
            }
        )
        if metadata.sped_type == "fiscal":
            products, _sales_rows, detailed_items, c190_rows, _c190_product_rows = read_sped_file(metadata.source_path)
            self.mysql_repo.replace_sped_extracted_data(archive_id, products, detailed_items, c190_rows)
        return archive_id

    def open_selected_archive_folder(self) -> None:
        file_tree = self.file_tree
        selected = file_tree.selection()
        if not selected:
            messagebox.showwarning("SPED Next", "Selecione um arquivo arquivado.")
            return
        row = self.archive_file_rows.get(int(selected[0]))
        if not row:
            return
        path = Path(str(row.get("caminho_arquivo_arquivado", "")))
        if not path.exists():
            path = Path(str(row.get("caminho_arquivo_original", "")))
        if path.exists():
            os.startfile(path.parent)

    def build_fiscal_queries_page(self) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        page.columnconfigure((0, 1, 2), weight=1)

        ttk.Label(
            page,
            text="Escolha uma consulta fiscal para abrir em tela propria.",
            style="Next.Muted.TLabel",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self.build_fiscal_menu_card(
            page,
            column=0,
            title="Consulta Entradas ICMS",
            description="Importar SPEDs/XMLs, comparar periodos e analisar entradas por produto, CST, CFOP, fornecedor e credito.",
            command=lambda: self.show_page("entry_consultation"),
        )
        self.build_fiscal_menu_card(
            page,
            column=1,
            title="Consulta Saidas ICMS",
            description="Proxima tela a migrar: saidas, recomposicao por XML e analise de debito.",
            command=lambda: messagebox.showinfo("SPED Next", "A tela de saidas sera migrada na proxima etapa."),
        )
        self.build_fiscal_menu_card(
            page,
            column=2,
            title="SPED x XML/Planilha",
            description="Comparacoes auxiliares ficarao separadas para nao misturar fluxos de trabalho.",
            command=lambda: messagebox.showinfo("SPED Next", "Esta consulta sera migrada em uma tela propria."),
        )
        return page

    def build_fiscal_menu_card(self, parent: ttk.Frame, column: int, title: str, description: str, command: object) -> None:
        card = ttk.Frame(parent, style="Next.Panel.TFrame", padding=18)
        card.grid(row=1, column=column, sticky="nsew", padx=(0 if column == 0 else 10, 0))
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text=title, style="Next.Panel.TLabel", font=("Segoe UI", 13, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(card, text=description, style="Next.PanelMuted.TLabel", wraplength=330).grid(row=1, column=0, sticky="ew", pady=(8, 16))
        ttk.Button(card, text="Abrir", style="Next.Primary.TButton", command=command).grid(row=2, column=0, sticky="w")

    def build_entry_consultation_page(self) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)

        toolbar = ttk.Frame(page, style="Next.Panel.TFrame", padding=12)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Button(toolbar, text="Voltar", style="Next.TButton", command=lambda: self.show_page("fiscal_queries")).pack(side=LEFT)
        ttk.Label(toolbar, text="Consulta Entradas ICMS", style="Next.Panel.TLabel", font=("Segoe UI", 11, "bold")).pack(side=LEFT, padx=(12, 0))
        ttk.Button(toolbar, text="Processar Entradas", style="Next.Primary.TButton", command=self.process_entry_query).pack(side=LEFT, padx=(16, 0))
        ttk.Button(toolbar, text="Limpar", style="Next.TButton", command=self.clear_entry_consultation_page).pack(side=RIGHT)

        setup = ttk.Frame(page, style="Next.TFrame")
        setup.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        setup.columnconfigure(0, weight=1)
        self.build_query_setup_card(
            setup,
            column=0,
            title="Consulta Entradas",
            description="Carregue SPEDs Fiscais e, se necessario, XMLs 55/65 para recompor itens sem C170 pela chave do C100.",
            sped_var=self.entry_sped_paths_var,
            xml_var=self.entry_xml_paths_var,
            process_command=self.process_entry_query,
            clear_command=lambda: self.clear_query_inputs(self.entry_sped_paths_var, self.entry_xml_paths_var),
        )

        results = ttk.Frame(page, style="Next.Panel.TFrame", padding=12)
        results.grid(row=2, column=0, sticky="nsew")
        results.columnconfigure(0, weight=1)
        results.rowconfigure(2, weight=1)
        summary_bar = ttk.Frame(results, style="Next.Panel.TFrame")
        summary_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(summary_bar, textvariable=self.entry_summary_var, style="Next.Panel.TLabel").pack(side=LEFT)
        self.build_entry_consultation_filters(results)

        entry_columns = ("periodo", "curva", "codigo", "descricao", "fornecedor", "ncm", "cest", "cfop", "cst", "aliq", "aliq_efetiva", "docs", "lancamentos")
        entry_widths = {"periodo": 90, "curva": 80, "codigo": 110, "descricao": 310, "fornecedor": 260, "ncm": 95, "cest": 95, "cfop": 85, "cst": 85, "aliq": 90, "aliq_efetiva": 100, "docs": 70, "lancamentos": 100}
        self.entry_query_box, self.entry_query_tree = self.create_tree(results, entry_columns, entry_widths)
        self.entry_query_box.grid(row=2, column=0, sticky="nsew")
        self.entry_query_tree.bind("<Double-1>", lambda _event: self.open_entry_docs_popup())
        return page

    def build_entry_consultation_filters(self, parent: ttk.Frame) -> None:
        filters = ttk.Frame(parent, style="Next.Panel.TFrame")
        filters.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        filters.columnconfigure(0, weight=2)
        filters.columnconfigure(1, weight=1)
        filters.columnconfigure(2, weight=1)
        filters.columnconfigure(3, weight=1)

        period_box = ttk.Frame(filters, style="Next.Panel.TFrame")
        period_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(period_box, text="Periodos para comparar", style="Next.PanelMuted.TLabel").pack(anchor="w")
        self.entry_period_checks_frame = ttk.Frame(period_box, style="Next.Panel.TFrame")
        self.entry_period_checks_frame.pack(fill="x", pady=(4, 0))

        cst_box = ttk.Frame(filters, style="Next.Panel.TFrame")
        cst_box.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        ttk.Label(cst_box, text="Filtro CST", style="Next.PanelMuted.TLabel").pack(anchor="w")
        ttk.Entry(cst_box, textvariable=self.entry_consult_cst_var).pack(fill="x", pady=(4, 0))
        ttk.Label(cst_box, text="Ex.: 000, 020, 060", style="Next.PanelMuted.TLabel").pack(anchor="w", pady=(4, 0))

        cfop_box = ttk.Frame(filters, style="Next.Panel.TFrame")
        cfop_box.grid(row=0, column=2, sticky="nsew", padx=(0, 8))
        ttk.Label(cfop_box, text="Filtro CFOP", style="Next.PanelMuted.TLabel").pack(anchor="w")
        ttk.Entry(cfop_box, textvariable=self.entry_consult_cfop_var).pack(fill="x", pady=(4, 0))
        ttk.Label(cfop_box, text="Ex.: 1101, 1401, 1556", style="Next.PanelMuted.TLabel").pack(anchor="w", pady=(4, 0))

        search_box = ttk.Frame(filters, style="Next.Panel.TFrame")
        search_box.grid(row=0, column=3, sticky="nsew")
        ttk.Label(search_box, text="Status / Busca", style="Next.PanelMuted.TLabel").pack(anchor="w")
        ttk.Combobox(
            search_box,
            textvariable=self.entry_consult_status_var,
            values=("Todos", "Ok", "Sem credito", "Sem entrada", "Com divergencia"),
            state="readonly",
        ).pack(fill="x", pady=(4, 4))
        ttk.Entry(search_box, textvariable=self.entry_consult_search_var).pack(fill="x")
        ttk.Label(search_box, text="Codigo ou descricao", style="Next.PanelMuted.TLabel").pack(anchor="w", pady=(4, 0))

        actions = ttk.Frame(parent, style="Next.Panel.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        for text, command, style in (
            ("Aplicar Filtros", self.refresh_entry_consultation_tree, "Next.Primary.TButton"),
            ("Limpar Filtros", self.clear_entry_consultation_filters, "Next.TButton"),
            ("Exportar Filtro Atual", self.export_entry_consultation_filter, "Next.TButton"),
            ("Entradas", self.open_entry_products_popup, "Next.TButton"),
            ("Comp. Diag. Credito", self.open_entry_credit_comparison_popup, "Next.TButton"),
            ("Diag. Credito", self.open_entry_credit_diagnostic_popup, "Next.TButton"),
            ("Curva ABC", self.open_entry_abc_popup, "Next.TButton"),
            ("Reducao BC", self.open_entry_reduction_popup, "Next.TButton"),
            ("Apuracao", self.open_entry_apuracao_popup, "Next.TButton"),
            ("Espelho Docs", self.open_entry_docs_popup, "Next.TButton"),
        ):
            ttk.Button(actions, text=text, style=style, command=command).pack(side=LEFT, padx=(0, 8))
        ttk.Label(parent, textvariable=self.entry_consult_footer_var, style="Next.Panel.TLabel").grid(row=4, column=0, sticky="w", pady=(8, 0))

    def build_query_setup_card(
        self,
        parent: ttk.Frame,
        column: int,
        title: str,
        description: str,
        sped_var: StringVar,
        xml_var: StringVar,
        process_command: object,
        clear_command: object,
    ) -> None:
        card = ttk.Frame(parent, style="Next.Panel.TFrame", padding=14)
        card.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 0))
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text=title, style="Next.Panel.TLabel", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(card, text=description, style="Next.PanelMuted.TLabel", wraplength=520).grid(row=1, column=0, sticky="ew", pady=(4, 10))
        self.build_path_row(card, 2, "SPEDs", sped_var, lambda: self.select_sped_files(sped_var), clear_command)
        self.build_path_row(card, 3, "XMLs 55/65", xml_var, lambda: self.select_xml_sources(xml_var), lambda: xml_var.set(""))
        actions = ttk.Frame(card, style="Next.Panel.TFrame")
        actions.grid(row=4, column=0, sticky="e", pady=(10, 0))
        ttk.Button(actions, text="Processar", style="Next.Primary.TButton", command=process_command).pack(side=LEFT)
        ttk.Button(actions, text="Limpar", style="Next.TButton", command=clear_command).pack(side=LEFT, padx=(8, 0))

    def build_path_row(self, parent: ttk.Frame, row: int, label: str, variable: StringVar, select_command: object, clear_command: object) -> None:
        line = ttk.Frame(parent, style="Next.Panel.TFrame")
        line.grid(row=row, column=0, sticky="ew", pady=(4, 0))
        line.columnconfigure(1, weight=1)
        ttk.Label(line, text=label, style="Next.PanelMuted.TLabel", width=10).grid(row=0, column=0, sticky="w")
        ttk.Entry(line, textvariable=variable).grid(row=0, column=1, sticky="ew", padx=(6, 8))
        ttk.Button(line, text="Selecionar", style="Next.TButton", command=select_command).grid(row=0, column=2)
        ttk.Button(line, text="Limpar", style="Next.TButton", command=clear_command).grid(row=0, column=3, padx=(6, 0))

    def select_sped_files(self, target_var: StringVar) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecionar SPEDs Fiscais",
                filetypes=[("Arquivos SPED", "*.txt *.sped *.efd"), ("Todos os arquivos", "*.*")],
            )
        )
        if not selected_files:
            return
        paths = append_unique_paths(parse_selected_paths(target_var.get()), selected_files)
        paths, limit_exceeded = limit_selected_paths(paths, 12)
        target_var.set(format_selected_paths(paths))
        if limit_exceeded:
            messagebox.showwarning("SPED Next", "A consulta aceita no maximo 12 SPEDs.")

    def select_xml_sources(self, target_var: StringVar) -> None:
        selected_files = list(
            filedialog.askopenfilenames(
                title="Selecionar XMLs 55/65",
                filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")],
            )
        )
        if selected_files:
            current_paths = parse_selected_paths(target_var.get())
            collapsed_paths, collapsed_to_folder = collapse_xml_selection_paths(selected_files)
            if collapsed_to_folder:
                current_paths = [path for path in current_paths if not path.is_file()]
            target_var.set(format_selected_paths(append_unique_paths(current_paths, collapsed_paths)))
            return
        selected_directory = filedialog.askdirectory(title="Selecionar pasta de XMLs 55/65")
        if selected_directory:
            target_var.set(format_selected_paths(append_unique_paths(parse_selected_paths(target_var.get()), [Path(selected_directory)])))

    def clear_query_inputs(self, sped_var: StringVar, xml_var: StringVar) -> None:
        sped_var.set("")
        xml_var.set("")

    def process_both_fiscal_queries(self) -> None:
        self.process_entry_query()
        self.process_sale_query()

    def process_entry_query(self) -> None:
        self.process_fiscal_query(
            "Entradas",
            self.entry_sped_paths_var,
            self.entry_xml_paths_var,
            self.entry_query_tree,
            self.entry_summary_var,
            build_entry_period_comparison_rows,
        )

    def process_sale_query(self) -> None:
        if not hasattr(self, "sale_query_tree"):
            messagebox.showinfo("SPED Next", "A tela de Consulta Saidas ICMS sera migrada separadamente.")
            return
        self.process_fiscal_query(
            "Saidas",
            self.sale_sped_paths_var,
            self.sale_xml_paths_var,
            self.sale_query_tree,
            self.sale_summary_var,
            build_sale_period_comparison_rows,
        )

    def process_fiscal_query(
        self,
        label: str,
        sped_var: StringVar,
        xml_var: StringVar,
        tree: ttk.Treeview,
        summary_var: StringVar,
        builder: object,
    ) -> None:
        try:
            sped_paths = parse_selected_paths(sped_var.get())
            xml_sources = parse_selected_paths(xml_var.get())
            period_labels, rows = builder(sped_paths, xml_sources)
            if label == "Entradas":
                self.entry_period_labels = list(period_labels)
                self.entry_consult_rows = list(rows)
                self.rebuild_entry_period_checks()
                self.refresh_entry_consultation_tree()
            else:
                self.fill_query_tree(tree, rows)
            summary_var.set(f"{label}: {len(rows)} linha(s), {len(period_labels)} periodo(s).")
            self.status_var.set(f"{label} processadas com sucesso.")
        except Exception as exc:
            summary_var.set(f"{label}: falha no processamento.")
            self.status_var.set(str(exc))
            messagebox.showerror("SPED Next", str(exc))

    def rebuild_entry_period_checks(self) -> None:
        if not hasattr(self, "entry_period_checks_frame"):
            return
        for child in self.entry_period_checks_frame.winfo_children():
            child.destroy()
        self.entry_period_check_vars = {}
        for label in self.entry_period_labels:
            variable = BooleanVar(value=True)
            self.entry_period_check_vars[label] = variable
            ttk.Checkbutton(
                self.entry_period_checks_frame,
                text=label,
                variable=variable,
                command=self.refresh_entry_consultation_tree,
            ).pack(side=LEFT, padx=(0, 8))

    def parse_filter_tokens(self, value: str) -> set[str]:
        return {part.strip() for part in value.replace(";", ",").replace("|", ",").split(",") if part.strip()}

    def normalize_search_text(self, value: object) -> str:
        text = str(value or "").lower()
        replacements = {
            "á": "a",
            "à": "a",
            "ã": "a",
            "â": "a",
            "é": "e",
            "ê": "e",
            "í": "i",
            "ó": "o",
            "ô": "o",
            "õ": "o",
            "ú": "u",
            "ç": "c",
        }
        for source, target in replacements.items():
            text = text.replace(source, target)
        return " ".join(text.split())

    def decimal_value(self, value: object) -> Decimal:
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(str(value or "0").replace(",", "."))
        except (InvalidOperation, ValueError):
            return Decimal("0")

    def first_row_value(self, row: dict[str, object], *keys: str, default: object = "") -> object:
        for key in keys:
            value = row.get(key)
            if value not in (None, ""):
                return value
        return default

    def get_selected_entry_periods(self) -> set[str]:
        if not self.entry_period_check_vars:
            return set(self.entry_period_labels)
        selected = {label for label, variable in self.entry_period_check_vars.items() if variable.get()}
        return selected or set(self.entry_period_labels)

    def get_entry_consultation_filtered_rows(self) -> list[dict[str, object]]:
        selected_periods = self.get_selected_entry_periods()
        cst_filters = self.parse_filter_tokens(self.entry_consult_cst_var.get())
        cfop_filters = self.parse_filter_tokens(self.entry_consult_cfop_var.get())
        status_filter = self.entry_consult_status_var.get().strip()
        search_text = self.normalize_search_text(self.entry_consult_search_var.get())

        filtered: list[dict[str, object]] = []
        for row in self.entry_consult_rows:
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
                    f"{row.get('code', '')} {row.get('description', '')} {row.get('cest', '')} "
                    f"{row.get('suppliers', '')} {row.get('supplier', '')} {row.get('participant_name', '')}"
                )
                if search_text not in searchable:
                    continue
            filtered.append(row)
        return filtered

    def refresh_entry_consultation_tree(self) -> None:
        if not hasattr(self, "entry_query_tree"):
            return
        self.entry_query_tree.delete(*self.entry_query_tree.get_children())
        filtered = self.get_entry_consultation_filtered_rows()
        self.entry_consult_filtered_rows = filtered

        total_sale = Decimal("0")
        total_base = Decimal("0")
        total_icms = Decimal("0")
        products: set[str] = set()
        for index, row in enumerate(filtered):
            sale_value = self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value", "operation_value"))
            base_icms = self.decimal_value(row.get("base_icms"))
            icms_value = self.decimal_value(row.get("icms_value"))
            total_sale += sale_value
            total_base += base_icms
            total_icms += icms_value
            code = str(row.get("code", "")).strip()
            if code:
                products.add(code)
            self.entry_query_tree.insert(
                "",
                END,
                iid=str(index),
                values=(
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
                ),
            )

        ratio = (total_base * Decimal("100") / total_sale).quantize(Decimal("0.01")) if total_sale else Decimal("0.00")
        self.entry_consult_footer_var.set(
            f"Linhas: {len(filtered)}    Produtos: {len(products)}    Valor Operacao: {self.format_number(total_sale)}    "
            f"Base ICMS: {self.format_number(total_base)}    % Base/Oper: {self.format_number(ratio)}    Valor ICMS: {self.format_number(total_icms)}"
        )

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

    def fill_query_tree(self, tree: ttk.Treeview, rows: list[dict[str, object]]) -> None:
        tree.delete(*tree.get_children())
        for row in rows[:2000]:
            tree.insert(
                "",
                END,
                values=(
                    row.get("period", row.get("period_label", "")),
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("cfop", ""),
                    row.get("cst_icms", ""),
                    self.format_number(row.get("sale_value", row.get("total_operation_value", ""))),
                    self.format_number(row.get("base_icms", "")),
                    self.format_number(row.get("icms_value", "")),
                    row.get("status", row.get("diagnostic_status", "")),
                ),
            )

    def open_table_popup(self, title: str, columns: tuple[str, ...], rows: list[tuple[object, ...]], size: str = "1100x620") -> None:
        if not rows:
            messagebox.showwarning(title, "Nao ha dados para os filtros atuais.")
            return
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.geometry(size)
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)
        ttk.Label(dialog, text=title, style="Next.Panel.TLabel", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        box = ttk.Frame(dialog, style="Next.Panel.TFrame", padding=8)
        box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        tree = ttk.Treeview(box, columns=columns, show="headings", selectmode="browse")
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=140, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        for row in rows:
            tree.insert("", END, values=row)

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
            for row in self.entry_consult_filtered_rows
        ]
        self.open_table_popup(
            "Entradas filtradas",
            ("Periodo", "Codigo", "Descricao", "Fornecedor", "CFOP", "CST", "Valor Operacao", "Base ICMS", "Valor ICMS", "Status"),
            rows,
            "1280x680",
        )

    def open_entry_abc_popup(self) -> None:
        grouped: dict[str, dict[str, Decimal | int]] = defaultdict(lambda: {"produtos": 0, "valor": Decimal("0"), "base": Decimal("0"), "icms": Decimal("0")})
        for row in self.entry_consult_filtered_rows:
            curve = str(self.first_row_value(row, "curve_abc", "abc_curve", default="C"))
            grouped[curve]["produtos"] = int(grouped[curve]["produtos"]) + 1
            grouped[curve]["valor"] = Decimal(grouped[curve]["valor"]) + self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value"))
            grouped[curve]["base"] = Decimal(grouped[curve]["base"]) + self.decimal_value(row.get("base_icms"))
            grouped[curve]["icms"] = Decimal(grouped[curve]["icms"]) + self.decimal_value(row.get("icms_value"))
        rows = [
            (curve, values["produtos"], self.format_number(values["valor"]), self.format_number(values["base"]), self.format_number(values["icms"]))
            for curve, values in sorted(grouped.items())
        ]
        self.open_table_popup("Curva ABC - Entradas", ("Curva", "Produtos", "Valor Operacao", "Base ICMS", "Valor ICMS"), rows, "820x480")

    def open_entry_reduction_popup(self) -> None:
        rows: list[tuple[object, ...]] = []
        for row in self.entry_consult_filtered_rows:
            sale_value = self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value"))
            base_icms = self.decimal_value(row.get("base_icms"))
            if sale_value <= 0 or base_icms >= sale_value:
                continue
            reduction = sale_value - base_icms
            rows.append(
                (
                    self.first_row_value(row, "period", "period_label"),
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("cfop", ""),
                    self.first_row_value(row, "cst_icms", "cst"),
                    self.format_number(sale_value),
                    self.format_number(base_icms),
                    self.format_number(reduction),
                )
            )
        self.open_table_popup("Reducao BC - Entradas", ("Periodo", "Codigo", "Descricao", "CFOP", "CST", "Valor Operacao", "Base ICMS", "Reducao BC"), rows, "1180x620")

    def open_entry_apuracao_popup(self) -> None:
        grouped: dict[tuple[str, str, str], dict[str, Decimal | int]] = defaultdict(lambda: {"lancamentos": 0, "valor": Decimal("0"), "base": Decimal("0"), "icms": Decimal("0")})
        for row in self.entry_consult_filtered_rows:
            key = (
                str(self.first_row_value(row, "period", "period_label")),
                str(row.get("cfop", "")),
                str(self.first_row_value(row, "cst_icms", "cst")),
            )
            grouped[key]["lancamentos"] = int(grouped[key]["lancamentos"]) + int(self.first_row_value(row, "launch_count", "launches", default=1) or 1)
            grouped[key]["valor"] = Decimal(grouped[key]["valor"]) + self.decimal_value(self.first_row_value(row, "sale_value", "total_operation_value"))
            grouped[key]["base"] = Decimal(grouped[key]["base"]) + self.decimal_value(row.get("base_icms"))
            grouped[key]["icms"] = Decimal(grouped[key]["icms"]) + self.decimal_value(row.get("icms_value"))
        rows = [
            (period, cfop, cst, values["lancamentos"], self.format_number(values["valor"]), self.format_number(values["base"]), self.format_number(values["icms"]))
            for (period, cfop, cst), values in sorted(grouped.items())
        ]
        self.open_table_popup("Apuracao ICMS - Entradas", ("Periodo", "CFOP", "CST", "Lancamentos", "Valor Operacao", "Base ICMS", "Valor ICMS"), rows, "1050x560")

    def open_entry_docs_popup(self) -> None:
        selected = self.entry_query_tree.selection() if hasattr(self, "entry_query_tree") else ()
        if not selected:
            messagebox.showwarning("Espelho Docs", "Selecione uma linha da consulta de entradas.")
            return
        row_index = int(selected[0])
        if row_index >= len(self.entry_consult_filtered_rows):
            return
        row = self.entry_consult_filtered_rows[row_index]
        detail_rows = row.get("launch_details") or row.get("details") or []
        rows: list[tuple[object, ...]] = []
        if isinstance(detail_rows, list) and detail_rows:
            for detail in detail_rows:
                if not isinstance(detail, dict):
                    continue
                rows.append(
                    (
                        detail.get("period", self.first_row_value(row, "period", "period_label")),
                        detail.get("document_date", detail.get("date", "")),
                        detail.get("document_number", ""),
                        detail.get("participant_name", ""),
                        detail.get("document_key", ""),
                        detail.get("item_number", ""),
                        detail.get("code", row.get("code", "")),
                        detail.get("description", row.get("description", "")),
                        detail.get("cfop", row.get("cfop", "")),
                        detail.get("cst_icms", self.first_row_value(row, "cst_icms", "cst")),
                        self.format_number(detail.get("sale_value", detail.get("total_operation_value", ""))),
                        self.format_number(detail.get("base_icms", "")),
                        self.format_number(detail.get("icms_value", "")),
                    )
                )
        else:
            rows.append(
                (
                    self.first_row_value(row, "period", "period_label"),
                    "",
                    "",
                    self.format_supplier_text(row),
                    "",
                    "",
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("cfop", ""),
                    self.first_row_value(row, "cst_icms", "cst"),
                    self.format_number(self.first_row_value(row, "sale_value", "total_operation_value")),
                    self.format_number(row.get("base_icms")),
                    self.format_number(row.get("icms_value")),
                )
            )
        self.open_table_popup(
            "Espelho Docs - Entradas",
            ("Periodo", "Data", "Documento", "Fornecedor", "Chave", "Item", "Codigo", "Descricao", "CFOP", "CST", "Valor Operacao", "Base ICMS", "Valor ICMS"),
            rows,
            "1380x720",
        )

    def open_entry_credit_diagnostic_popup(self) -> None:
        rows = []
        for row in self.entry_consult_filtered_rows:
            status = str(row.get("status", "")).strip()
            if status == "Ok":
                continue
            rows.append(
                (
                    self.first_row_value(row, "period", "period_label"),
                    row.get("code", ""),
                    row.get("description", ""),
                    row.get("cfop", ""),
                    self.first_row_value(row, "cst_icms", "cst"),
                    self.format_number(row.get("icms_value")),
                    status,
                )
            )
        self.open_table_popup("Diagnostico de Credito - Entradas", ("Periodo", "Codigo", "Descricao", "CFOP", "CST", "Valor ICMS", "Status"), rows, "1120x600")

    def open_entry_credit_comparison_popup(self) -> None:
        grouped: dict[str, dict[str, object]] = {}
        for row in self.entry_consult_filtered_rows:
            code = str(row.get("code", "")).strip()
            if not code:
                continue
            bucket = grouped.setdefault(code, {"descricao": row.get("description", ""), "periodos": set(), "credito": Decimal("0"), "status": set()})
            bucket["periodos"].add(str(self.first_row_value(row, "period", "period_label")))
            bucket["credito"] = Decimal(bucket["credito"]) + self.decimal_value(row.get("icms_value"))
            status = str(row.get("status", "")).strip()
            if status:
                bucket["status"].add(status)
        rows = [
            (code, data["descricao"], len(data["periodos"]), self.format_number(data["credito"]), " | ".join(sorted(data["status"])))
            for code, data in sorted(grouped.items())
        ]
        self.open_table_popup("Comparacao Diagnostico de Credito", ("Codigo", "Descricao", "Periodos", "Credito ICMS", "Status"), rows, "1060x580")

    def export_entry_consultation_filter(self) -> None:
        if not self.entry_consult_filtered_rows:
            messagebox.showwarning("Exportar consulta", "Nao ha dados filtrados para exportar.")
            return
        output = filedialog.asksaveasfilename(
            title="Salvar consulta filtrada",
            defaultextension=".csv",
            initialfile="consulta_entradas_filtrada.csv",
            filetypes=[("Arquivo CSV", "*.csv")],
        )
        if not output:
            return
        headers = ["Periodo", "Curva ABC", "Codigo", "Descricao", "Fornecedor", "NCM", "CEST", "CFOP", "CST", "Aliq ICMS", "Aliq Efetiva", "Docs", "Lancamentos", "Valor Operacao", "Base ICMS", "Valor ICMS", "Status"]
        with Path(output).open("w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.writer(csv_file, delimiter=";")
            writer.writerow(headers)
            for row in self.entry_consult_filtered_rows:
                writer.writerow(
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
                        self.format_number(self.first_row_value(row, "icms_rate", "nominal_icms_rate")),
                        self.format_number(self.first_row_value(row, "effective_icms_rate", "icms_effective_rate")),
                        self.first_row_value(row, "document_count", "documents", default=0),
                        self.first_row_value(row, "launch_count", "launches", default=0),
                        self.format_number(self.first_row_value(row, "sale_value", "total_operation_value")),
                        self.format_number(row.get("base_icms")),
                        self.format_number(row.get("icms_value")),
                        row.get("status", ""),
                    ]
                )
        self.status_var.set(f"Consulta de entradas exportada: {output}")
        messagebox.showinfo("Exportar consulta", f"Arquivo gerado com sucesso:\n{output}")

    def clear_entry_consultation_page(self) -> None:
        self.clear_query_inputs(self.entry_sped_paths_var, self.entry_xml_paths_var)
        if hasattr(self, "entry_query_tree"):
            self.entry_query_tree.delete(*self.entry_query_tree.get_children())
        self.entry_period_labels = []
        self.entry_consult_rows = []
        self.entry_consult_filtered_rows = []
        self.rebuild_entry_period_checks()
        self.entry_consult_footer_var.set("Linhas: 0    Produtos: 0    Valor Operacao: 0,00    Base ICMS: 0,00    % Base/Oper: 0,00    Valor ICMS: 0,00")
        self.entry_summary_var.set("Entradas: nenhum processamento executado.")

    def clear_fiscal_queries(self) -> None:
        self.clear_entry_consultation_page()
        self.clear_query_inputs(self.sale_sped_paths_var, self.sale_xml_paths_var)
        if hasattr(self, "sale_query_tree"):
            self.sale_query_tree.delete(*self.sale_query_tree.get_children())
        self.sale_summary_var.set("Saidas: nenhum processamento executado.")

    def clear_entry_consultation_filters(self) -> None:
        self.entry_consult_cst_var.set("")
        self.entry_consult_cfop_var.set("")
        self.entry_consult_status_var.set("Todos")
        self.entry_consult_search_var.set("")
        for variable in self.entry_period_check_vars.values():
            variable.set(True)
        self.refresh_entry_consultation_tree()

    def format_number(self, value: object) -> str:
        text = str(value if value is not None else "").strip()
        return text.replace(".", ",") if text else ""

    def build_settings_page(self) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        page.columnconfigure(0, weight=1)
        panel = ttk.Frame(page, style="Next.Panel.TFrame", padding=16)
        panel.grid(row=0, column=0, sticky="ew")
        ttk.Label(panel, text=f"Config MySQL: {self.mysql_config_path}", style="Next.Panel.TLabel").pack(anchor="w")
        ttk.Button(panel, text="Criar/Atualizar Banco", style="Next.Primary.TButton", command=self.create_schema).pack(anchor="w", pady=(12, 0))
        return page

    def create_schema(self) -> None:
        try:
            self.mysql_repo.ensure_schema()
            self.status_var.set("Banco e tabelas atualizados.")
            messagebox.showinfo("SPED Next", "Banco e tabelas atualizados.")
        except Exception as exc:
            self.status_var.set(f"Falha ao atualizar banco: {exc}")
            messagebox.showerror("SPED Next", str(exc))

    def build_placeholder_page(self, title: str, description: str) -> ttk.Frame:
        page = ttk.Frame(self.content_host, style="Next.TFrame")
        panel = ttk.Frame(page, style="Next.Panel.TFrame", padding=20)
        panel.grid(row=0, column=0, sticky="ew")
        ttk.Label(panel, text=title, style="Next.Panel.TLabel", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(panel, text=description, style="Next.PanelMuted.TLabel").pack(anchor="w", pady=(8, 0))
        return page

    def format_period(self, start: object, end: object) -> str:
        start_text = str(start or "")
        end_text = str(end or "")
        if start_text and end_text:
            return f"{start_text} a {end_text}"
        return start_text or end_text
