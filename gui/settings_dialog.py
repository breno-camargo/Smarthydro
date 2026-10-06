import os
import subprocess
import threading
import tempfile
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core.config_manager import (
    load_config, save_config, get_operators, get_active_operator, set_active_operator,
    get_default_condominio_emails
)
from core.database import get_available_odbc_drivers, test_db_connection
from core.report_generator import open_template_in_excel
from core.email_sender import (
    test_smtp_connection, get_email_template_path, load_email_template,
    render_email, prepare_html_for_preview, DEFAULT_SUBJECT_TEMPLATE
)
from gui.email_template_dialog import EmailTemplateDialog
from gui.scheduler_dialog import SchedulerDialog
from gui.operators_dialog import OperatorsDialog
from gui.ui_helpers import (
    apply_window_icon, center_modal, setup_common_styles,
    create_btn_primary, create_btn_secondary, create_btn_danger,
    bind_button_hover, create_card_frame, create_modern_badge, create_tooltip,
    COLOR_PRIMARY, COLOR_PRIMARY_HOVER, COLOR_ACCENT, COLOR_BG_LIGHT,
    COLOR_TEXT_MAIN, COLOR_TEXT_MUTED, COLOR_CARD_BG, COLOR_CARD_BORDER
)

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, on_save_callback=None, initial_tab=0):
        super().__init__(parent)
        self.title("Configurações do Sistema — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)

        self.on_save_callback = on_save_callback
        self.initial_tab = initial_tab
        self.config = load_config()

        # Configurar estilo visual moderno das abas
        self._configure_notebook_style()

        self._build_ui()
        self._load_values()

        if self.initial_tab > 0:
            try:
                self.notebook.select(self.initial_tab)
            except Exception:
                pass

        # Centralizar perfeitamente sobre a janela principal com dimensões ideais
        center_modal(self, parent, 630, 610)

    def _configure_notebook_style(self):
        self.style = ttk.Style(self)
        setup_common_styles(self.style)
        self.style.configure("Settings.TNotebook", background=COLOR_BG_LIGHT, borderwidth=0)
        self.style.configure(
            "Settings.TNotebook.Tab",
            font=("Segoe UI", 9, "bold"),
            padding=[10, 5]
        )

    def _build_ui(self):
        container = ttk.Frame(self, padding="14 10 14 10")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── HEADER BANNER CORPORATIVO ───
        frame_header = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_header.pack(fill=tk.X, pady=(0, 6))

        lbl_hdr_title = tk.Label(
            frame_header, text="⚙️  Configurações do Sistema",
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_hdr_title.pack(anchor=tk.W)
        lbl_hdr_sub = tk.Label(
            frame_header, text="Parâmetros de Conexão • Servidor E-mail • Operadores • Notificações & Backup",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_hdr_sub.pack(anchor=tk.W)

        # ─── BOTÕES DE AÇÃO INFERIORES (Fixos no rodapé com prioridade total de espaço) ───
        frame_btns = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

        btn_save = create_btn_primary(
            frame_btns, "✔ Salvar Alterações", self._save_and_close,
            padx=18, pady=6
        )
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = create_btn_secondary(
            frame_btns, "Cancelar", self.destroy,
            padx=14, pady=6
        )
        btn_cancel.pack(side=tk.RIGHT)

        # ─── NOTEBOOK DAS ABAS (Preenche todo o espaço acima do rodapé) ───
        self.notebook = ttk.Notebook(container, style="Settings.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # ─── ABA 1: BANCO DE DADOS & EXCEL ───
        tab_db = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_db, text=" 🏢 Banco ")
        self._build_tab_db(tab_db)

        # ─── ABA 2: E-MAIL & UOL PRO ───
        tab_email = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_email, text=" ✉ E-mail ")
        self._build_tab_email(tab_email)

        # ─── ABA 3: OPERADORES & PERFIS ───
        tab_ops = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_ops, text=" 👤 Operadores ")
        self._build_tab_operators(tab_ops)

        # ─── ABA 4: NOTIFICAÇÕES & WEBHOOKS ───
        tab_webhooks = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_webhooks, text=" 📲 Webhooks ")
        self._build_tab_webhooks(tab_webhooks)

        # ─── ABA 5: BACKUP & DADOS ───
        tab_backup = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_backup, text=" 💾 Backup ")
        self._build_tab_backup(tab_backup)

        # ─── ABA 6: DESENVOLVEDOR ───
        tab_dev = ttk.Frame(self.notebook, padding="14 8 14 8")
        self.notebook.add(tab_dev, text=" 💻 Sobre ")
        self._build_tab_dev(tab_dev)

    def _build_tab_db(self, parent):
        """Constrói a aba de banco de dados com cards modernos e organizados."""
        # Card 1: Conexão ao SQL Server
        card_sql = create_card_frame(parent, padx=12, pady=8)
        card_sql.pack(fill=tk.X, pady=(0, 6))
        card_sql.columnconfigure(1, weight=1)

        tk.Label(
            card_sql, text="🗄️ Conexão ao SQL Server (StruxureWare EBO)",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tk.Label(card_sql, text="Servidor/Instância:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ent_server = ttk.Entry(card_sql, width=36)
        self.ent_server.grid(row=1, column=1, sticky=tk.EW, pady=2)

        tk.Label(card_sql, text="Banco de Dados:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.ent_database = ttk.Entry(card_sql, width=36)
        self.ent_database.grid(row=2, column=1, sticky=tk.EW, pady=2)

        tk.Label(card_sql, text="Driver ODBC:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=3, column=0, sticky=tk.W, pady=2)
        drivers = get_available_odbc_drivers()
        if not drivers:
            drivers = ["ODBC Driver 17 for SQL Server", "SQL Server"]
        self.cmb_driver = ttk.Combobox(card_sql, values=drivers, state="readonly", width=34)
        self.cmb_driver.grid(row=3, column=1, sticky=tk.EW, pady=2)

        # Autenticação
        self.var_trusted = tk.BooleanVar(value=True)
        self.chk_trusted = ttk.Checkbutton(
            card_sql, text="Usar Autenticação Integrada do Windows (Trusted)",
            variable=self.var_trusted, command=self._toggle_auth
        )
        self.chk_trusted.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(3, 1))

        self.lbl_trusted_hint = tk.Label(
            card_sql,
            text="🔒 Autenticação integrada do Windows ativa (usuário e senha dispensados).",
            font=("Segoe UI", 8, "italic"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        )

        self.frame_auth = tk.Frame(card_sql, bg="#FFFFFF")
        self.frame_auth.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=1)

        tk.Label(self.frame_auth, text="Usuário:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=0, column=0, sticky=tk.W, pady=1)
        self.ent_user = ttk.Entry(self.frame_auth, width=15)
        self.ent_user.grid(row=0, column=1, sticky=tk.W, padx=4, pady=1)

        tk.Label(self.frame_auth, text="Senha:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=0, column=2, sticky=tk.W, padx=(12, 0), pady=1)
        self.ent_pass = ttk.Entry(self.frame_auth, width=15, show="*")
        self.ent_pass.grid(row=0, column=3, sticky=tk.W, padx=4, pady=1)

        f_sql_btns = tk.Frame(card_sql, bg="#FFFFFF")
        f_sql_btns.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))
        btn_test_db = create_btn_secondary(f_sql_btns, "🔌 Testar Conexão SQL", self._test_connection, pady=2, padx=10)
        btn_test_db.pack(side=tk.LEFT)

        # Card 2: Pasta Padrão & Modelo Excel
        card_rep = create_card_frame(parent, padx=12, pady=8)
        card_rep.pack(fill=tk.X, pady=(0, 6))
        card_rep.columnconfigure(1, weight=1)

        tk.Label(
            card_rep, text="📂 Pasta Padrão & Modelo Excel",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tk.Label(card_rep, text="Pasta de Saída:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=1, column=0, sticky=tk.W, pady=2)
        frame_dir = tk.Frame(card_rep, bg="#FFFFFF")
        frame_dir.grid(row=1, column=1, sticky=tk.EW, pady=2)
        self.ent_dir = ttk.Entry(frame_dir, width=24)
        self.ent_dir.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_browse_dir = create_btn_secondary(frame_dir, "Alterar...", self._browse_dir, pady=2, padx=8)
        btn_browse_dir.pack(side=tk.LEFT, padx=(4, 0))

        btn_edit_model = create_btn_secondary(
            card_rep, "✏ Abrir Modelo no Excel para Edição (modelo_relatorio.xlsx)",
            self._edit_template_action, pady=2, padx=10
        )
        btn_edit_model.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

        # Card 3: Preferências Gerais
        card_pref = create_card_frame(parent, padx=12, pady=8)
        card_pref.pack(fill=tk.X)

        tk.Label(
            card_pref, text="⚙️ Preferências Gerais do Relatório",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 3))

        self.var_open_excel = tk.BooleanVar(value=True)
        self.chk_open_excel = ttk.Checkbutton(
            card_pref, text="Abrir planilha no Microsoft Excel automaticamente após geração",
            variable=self.var_open_excel
        )
        self.chk_open_excel.pack(anchor=tk.W, pady=1)

        self.var_sort_consumption = tk.BooleanVar(value=True)
        self.chk_sort_consumption = ttk.Checkbutton(
            card_pref, text="Ordenar unidades por maior consumo no relatório (ranking decrescente)",
            variable=self.var_sort_consumption
        )
        self.chk_sort_consumption.pack(anchor=tk.W, pady=1)

        btn_scheduler = create_btn_secondary(
            card_pref,
            "⏰ Configurar Agendamento Automático no Windows (Executar dia 29 silencioso)",
            self._open_scheduler, pady=2, padx=10
        )
        btn_scheduler.pack(anchor=tk.W, pady=(4, 0))

    def _build_tab_email(self, parent):
        """Constrói a aba de e-mail com cards modernos para destinatários, SMTP e modelo."""
        # Card 1: Destinatários Oficiais
        card_dest = create_card_frame(parent, padx=12, pady=8)
        card_dest.pack(fill=tk.X, pady=(0, 6))
        card_dest.columnconfigure(1, weight=1)

        tk.Label(
            card_dest, text="🏢 Destinatários Padrão (Condomínio Praça Pamplona)",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tk.Label(card_dest, text="E-mail da Gerente (Para):", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ent_recipients = ttk.Entry(card_dest, width=36)
        self.ent_recipients.grid(row=1, column=1, sticky=tk.EW, pady=2)

        tk.Label(card_dest, text="Em Cópia Padrão (Cc):", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=2, column=0, sticky=tk.W, pady=2)
        frame_cc = tk.Frame(card_dest, bg="#FFFFFF")
        frame_cc.grid(row=2, column=1, sticky=tk.EW, pady=2)

        self.ent_cc = ttk.Entry(frame_cc, width=28)
        self.ent_cc.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_view_cc = create_btn_secondary(
            frame_cc, "🔍 Ver Todos", self._show_all_cc, pady=2, padx=8
        )
        btn_view_cc.pack(side=tk.LEFT, padx=(4, 0))

        lbl_hint_to = tk.Label(
            card_dest, text="Separe múltiplos e-mails por ponto-e-vírgula (;)",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_hint_to.grid(row=3, column=1, sticky=tk.W, pady=(0, 2))

        btn_restore_def = create_btn_secondary(
            card_dest, "🔄 Restaurar E-mails Padrão (Praça Pamplona)",
            self._restore_default_emails, pady=2, padx=10
        )
        btn_restore_def.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(3, 0))
        create_tooltip(btn_restore_def, "Restaura os e-mails oficiais da Gerente e lista Cc do Condomínio Praça Pamplona")

        # Card 2: Servidor SMTP UOL Pro
        card_smtp = create_card_frame(parent, padx=12, pady=8)
        card_smtp.pack(fill=tk.X, pady=(0, 6))
        card_smtp.columnconfigure(1, weight=1)

        tk.Label(
            card_smtp, text="✉️ Sua Conta de E-mail (UOL Pro / Remetente)",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tk.Label(card_smtp, text="Servidor SMTP:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ent_smtp_host = ttk.Entry(card_smtp, width=36)
        self.ent_smtp_host.grid(row=1, column=1, sticky=tk.EW, pady=2)

        tk.Label(card_smtp, text="Porta / Criptografia:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=2, column=0, sticky=tk.W, pady=2)
        frame_port_sec = tk.Frame(card_smtp, bg="#FFFFFF")
        frame_port_sec.grid(row=2, column=1, sticky=tk.W, pady=2)

        self.cmb_smtp_preset = ttk.Combobox(
            frame_port_sec,
            values=[
                "Porta 465 (SSL / TLS Seguro — Padrão UOL Pro)",
                "Porta 587 (STARTTLS — Padrão Corporativo)",
                "Personalizado"
            ],
            state="readonly",
            width=36
        )
        self.cmb_smtp_preset.pack(side=tk.LEFT)
        self.cmb_smtp_preset.bind("<<ComboboxSelected>>", self._on_smtp_preset_change)

        self.frame_custom_port = tk.Frame(card_smtp, bg="#FFFFFF")
        self.frame_custom_port.grid(row=3, column=1, sticky=tk.W, pady=(2, 2))

        tk.Label(self.frame_custom_port, text="Porta:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)
        self.ent_smtp_port = ttk.Entry(self.frame_custom_port, width=6)
        self.ent_smtp_port.pack(side=tk.LEFT, padx=(4, 10))

        self.var_smtp_tls = tk.BooleanVar(value=False)
        self.chk_smtp_tls = ttk.Checkbutton(self.frame_custom_port, text="STARTTLS", variable=self.var_smtp_tls)
        self.chk_smtp_tls.pack(side=tk.LEFT, padx=(0, 8))

        self.var_smtp_ssl = tk.BooleanVar(value=True)
        self.chk_smtp_ssl = ttk.Checkbutton(self.frame_custom_port, text="SSL Direto", variable=self.var_smtp_ssl)
        self.chk_smtp_ssl.pack(side=tk.LEFT)

        # Bloco de Autenticação Integrado ao Operador Ativo (sem duplicidade!)
        frame_op_auth = tk.Frame(card_smtp, bg="#F0F7EE", highlightbackground="#D5E5C9", highlightthickness=1, padx=10, pady=7)
        frame_op_auth.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=(6, 4))

        self.lbl_smtp_op_auth = tk.Label(
            frame_op_auth,
            text="👤 Remetente Ativo: Carregando...",
            font=("Segoe UI", 8, "bold"), fg=COLOR_PRIMARY, bg="#F0F7EE", anchor="w"
        )
        self.lbl_smtp_op_auth.pack(fill=tk.X)

        self.lbl_smtp_op_sub = tk.Label(
            frame_op_auth,
            text="O e-mail e a senha de envio são obtidos automaticamente do perfil do operador ativo.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#F0F7EE", anchor="w"
        )
        self.lbl_smtp_op_sub.pack(fill=tk.X, pady=(1, 5))

        frame_smtp_btns = tk.Frame(frame_op_auth, bg="#F0F7EE")
        frame_smtp_btns.pack(fill=tk.X)

        btn_test_smtp = create_btn_secondary(
            frame_smtp_btns, "🔌 Testar Conexão com Servidor", self._test_smtp, pady=2, padx=10
        )
        btn_test_smtp.pack(side=tk.LEFT, padx=(0, 6))

        btn_go_ops = create_btn_secondary(
            frame_smtp_btns, "👤 Gerenciar Operadores & Senhas", self._open_operators_manager, pady=2, padx=10
        )
        btn_go_ops.pack(side=tk.LEFT)

        # Card 3: Modelo de E-mail
        card_tmpl = create_card_frame(parent, padx=12, pady=8)
        card_tmpl.pack(fill=tk.X)

        tk.Label(
            card_tmpl, text="🎨 Personalização do Modelo de E-mail",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 4))

        frame_tmpl_btns = tk.Frame(card_tmpl, bg="#FFFFFF")
        frame_tmpl_btns.pack(fill=tk.X)

        btn_edit_email_tmpl = create_btn_secondary(
            frame_tmpl_btns, "✏ Personalizar Modelo",
            self._open_email_template_editor, pady=2, padx=10
        )
        btn_edit_email_tmpl.pack(side=tk.LEFT, padx=(0, 6))

        btn_preview_tmpl = create_btn_secondary(
            frame_tmpl_btns, "👁 Ver Prévia no Navegador",
            self._preview_email_action, pady=2, padx=10
        )
        btn_preview_tmpl.pack(side=tk.LEFT, padx=(0, 6))

        btn_open_tmpl_file = create_btn_secondary(
            frame_tmpl_btns, "📂 Bloco de Notas",
            self._open_template_in_notepad, pady=2, padx=10
        )
        btn_open_tmpl_file.pack(side=tk.LEFT)

    def _build_tab_operators(self, parent):
        lbl_sec = tk.Label(
            parent, text="Perfis de Operadores & Assinaturas Corporativas",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            parent,
            text="Alterne o operador ativo ou gerencie múltiplos perfis. Cada operador possui seu próprio e-mail e assinatura corporativa nos relatórios.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, wraplength=540, justify=tk.LEFT
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 10))

        # Card de Resumo do Operador Ativo Moderno
        self.frame_active_card = tk.Frame(
            parent, bg="#FFFFFF", highlightbackground="#D5E5C9", highlightthickness=1,
            padx=14, pady=12
        )
        self.frame_active_card.pack(fill=tk.X, pady=(0, 10))

        # Linha 1: Seletor
        f_sel = tk.Frame(self.frame_active_card, bg="#FFFFFF")
        f_sel.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            f_sel, text="👤 Selecionar Operador Ativo:", font=("Segoe UI", 9, "bold"),
            fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.cmb_tab_operator = ttk.Combobox(f_sel, state="readonly", width=36)
        self.cmb_tab_operator.pack(side=tk.LEFT)
        self.cmb_tab_operator.bind("<<ComboboxSelected>>", self._on_tab_operator_change)

        div_op = tk.Frame(self.frame_active_card, height=1, bg="#E8F1E4")
        div_op.pack(fill=tk.X, pady=(4, 8))

        # Detalhes do Perfil
        f_prof = tk.Frame(self.frame_active_card, bg="#FFFFFF")
        f_prof.pack(fill=tk.X)

        f_name_row = tk.Frame(f_prof, bg="#FFFFFF")
        f_name_row.pack(fill=tk.X, anchor=tk.W)

        self.lbl_op_name = tk.Label(
            f_name_row, text="", font=("Segoe UI", 10, "bold"),
            fg=COLOR_TEXT_MAIN, bg="#FFFFFF"
        )
        self.lbl_op_name.pack(side=tk.LEFT, padx=(0, 8))

        self.badge_op_status = tk.Label(
            f_name_row, text="", font=("Segoe UI", 8, "bold"),
            padx=6, pady=1
        )
        self.badge_op_status.pack(side=tk.LEFT)

        self.lbl_op_role = tk.Label(
            f_prof, text="", font=("Segoe UI", 8),
            fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        self.lbl_op_role.pack(anchor=tk.W, pady=(2, 6))

        # Badges de E-mail e WhatsApp
        f_badges = tk.Frame(f_prof, bg="#FFFFFF")
        f_badges.pack(fill=tk.X, anchor=tk.W)

        self.lbl_op_email = tk.Label(
            f_badges, text="", font=("Segoe UI", 8),
            fg=COLOR_TEXT_MAIN, bg="#F0F5EC", padx=8, pady=3,
            highlightbackground="#D5E2CF", highlightthickness=1
        )
        self.lbl_op_email.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_op_wpp = tk.Label(
            f_badges, text="", font=("Segoe UI", 8),
            fg=COLOR_PRIMARY, bg="#EBF4E5", padx=8, pady=3,
            highlightbackground="#C5DCBA", highlightthickness=1
        )
        self.lbl_op_wpp.pack(side=tk.LEFT)

        # Botão para abrir o gerenciador completo
        btn_open_mgr = create_btn_secondary(
            parent,
            "👥 Abrir Gerenciador de Operadores (Cadastrar / Editar)",
            self._open_operators_manager,
            pady=5
        )
        btn_open_mgr.pack(anchor=tk.W, pady=(4, 0))

        self._refresh_tab_operators()

    def _refresh_tab_operators(self):
        cfg = load_config()
        ops = get_operators(cfg)
        active = get_active_operator(cfg)
        self.ops_map = {f"{op.get('name')} ({op.get('role', 'Operador')})": op.get('id') for op in ops}
        names = list(self.ops_map.keys())
        self.cmb_tab_operator["values"] = names

        curr_key = None
        for name, op_id in self.ops_map.items():
            if active and op_id == active.get("id"):
                curr_key = name
                break
        if curr_key:
            self.cmb_tab_operator.set(curr_key)
        elif names:
            self.cmb_tab_operator.set(names[0])

        if active:
            wpp = active.get("whatsapp_phone") or active.get("phone", "Não informado")
            has_k = "Chave API Ativa" if (active.get("whatsapp_apikey") or cfg.get("webhook_whatsapp_apikey")) else "Sem chave API"

            self.lbl_op_name.config(text=active.get('name', 'Operador'))
            self.lbl_op_role.config(text=f"Função: {active.get('role', 'Técnico')} • Condomínio Praça Pamplona")

            if active.get('is_default'):
                self.badge_op_status.config(text="★ Operador Padrão (Agendamento Automático)", fg="#225E1A", bg="#E0F0D8")
            else:
                self.badge_op_status.config(text="● Operador Secundário", fg=COLOR_TEXT_MUTED, bg="#F0F4EC")

            self.lbl_op_email.config(text=f"✉ {active.get('email', 'Sem e-mail')}")
            self.lbl_op_wpp.config(text=f"📲 WhatsApp: {wpp} ({has_k})")

        self._update_operator_context_labels()

    def _on_tab_operator_change(self, event=None):
        val = self.cmb_tab_operator.get()
        if val in self.ops_map:
            set_active_operator(self.ops_map[val])
            self._refresh_tab_operators()

    def _open_operators_manager(self):
        OperatorsDialog(self, on_change_callback=self._refresh_tab_operators)

    def _update_operator_context_labels(self):
        """Atualiza os cards dinâmicos do operador ativo na Aba E-mail e Aba Webhooks."""
        cfg = load_config()
        active = get_active_operator(cfg)
        if not active:
            return

        # 1. Atualizar card na Aba E-mail
        if hasattr(self, "lbl_smtp_op_auth"):
            op_name = active.get("name", "Operador")
            op_email = active.get("email") or active.get("smtp_user") or "E-mail não cadastrado"
            has_pwd = bool(active.get("smtp_password"))
            pwd_txt = "●●●●●● (Senha salva)" if has_pwd else "⚠ Senha SMTP pendente"
            self.lbl_smtp_op_auth.config(
                text=f"👤 Remetente Ativo: {op_name} <{op_email}>",
                fg=COLOR_PRIMARY
            )
            if hasattr(self, "lbl_smtp_op_sub"):
                self.lbl_smtp_op_sub.config(
                    text=f"Autenticação SMTP: {pwd_txt} • Assinatura automática do operador inclusa nos e-mails.",
                    fg=COLOR_TEXT_MAIN if has_pwd else "#C0392B"
                )

        # 2. Atualizar card na Aba Webhooks
        if hasattr(self, "lbl_wh_wpp_contact"):
            from core.webhook_notifier import get_whatsapp_recipients
            recipients = get_whatsapp_recipients(self.config)
            if recipients:
                names_txt = ", ".join(f"{r['name']} ({r['phone'][-9:]})" for r in recipients)
                total_ops = len(recipients)
                plural = "operadores cadastrados" if total_ops > 1 else "operador cadastrado"
                self.lbl_wh_wpp_contact.config(
                    text=f"✔ Envio ativo para {total_ops} {plural}:\n{names_txt}",
                    fg="#2A6320"
                )
            else:
                self.lbl_wh_wpp_contact.config(
                    text="⚠ Nenhum operador com telefone e Chave API CallMeBot configurados.",
                    fg="#C0392B"
                )

    def _build_tab_webhooks(self, parent):
        """Constrói a aba de configuração de Webhooks para Teams, Discord, Slack e Telegram com design moderno de cards."""
        lbl_sec = tk.Label(
            parent, text="Notificações em Tempo Real (Webhooks & WhatsApp)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            parent,
            text="Envie alertas automáticos do fechamento mensal para seu WhatsApp ou canais de equipe (Microsoft Teams, Discord, Slack ou Telegram).",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, wraplength=540, justify=tk.LEFT
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 8))

        # Card 1: Ativação com Badge
        card_toggle = create_card_frame(parent, padx=12, pady=8)
        card_toggle.pack(fill=tk.X, pady=(0, 8))

        self.var_webhook_enabled = tk.BooleanVar(value=False)
        self.chk_webhook_enabled = ttk.Checkbutton(
            card_toggle, text="Ativar Notificações via Webhook ao concluir relatório",
            variable=self.var_webhook_enabled, command=self._update_wh_status_badge
        )
        self.chk_webhook_enabled.pack(side=tk.LEFT)

        self.lbl_wh_badge = tk.Label(
            card_toggle, text="○ INATIVO", font=("Segoe UI", 8, "bold"),
            fg=COLOR_TEXT_MUTED, bg="#F0F0F0", padx=8, pady=2
        )
        self.lbl_wh_badge.pack(side=tk.RIGHT)

        # Card 2: Configuração da Plataforma
        self.card_wh_config = create_card_frame(parent, padx=14, pady=10)
        self.card_wh_config.pack(fill=tk.X, pady=(0, 8))
        self.card_wh_config.columnconfigure(1, weight=1)

        tk.Label(
            self.card_wh_config, text="Configuração da Plataforma",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        tk.Label(self.card_wh_config, text="Plataforma:", font=("Segoe UI", 9), bg="#FFFFFF").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.wh_platform_names = [
            "WhatsApp (CallMeBot Grátis / Notificação Direta)",
            "Microsoft Teams (Incoming Webhook)",
            "Discord (Canal de Alertas)",
            "Slack (Incoming Webhook)",
            "Telegram (Bot API)",
            "Webhook Genérico / WhatsApp API (JSON POST)"
        ]
        self.wh_platform_keys = ["whatsapp", "teams", "discord", "slack", "telegram", "generic"]
        self.cmb_wh_platform = ttk.Combobox(
            self.card_wh_config, values=self.wh_platform_names, state="readonly", width=34
        )
        self.cmb_wh_platform.grid(row=1, column=1, sticky=tk.EW, pady=3)
        self.cmb_wh_platform.bind("<<ComboboxSelected>>", self._on_wh_platform_change)

        # Container dinâmico do WhatsApp integrado ao Operador Ativo (sem duplicidade!)
        self.frame_wh_whatsapp = tk.Frame(self.card_wh_config, bg="#F0F7EE", highlightbackground="#D5E5C9", highlightthickness=1, padx=10, pady=8)

        lbl_wpp_head = tk.Label(
            self.frame_wh_whatsapp,
            text="📲 Destinatário dos Alertas: Operador Ativo",
            font=("Segoe UI", 8, "bold"), fg=COLOR_PRIMARY, bg="#F0F7EE", anchor="w"
        )
        lbl_wpp_head.pack(fill=tk.X)

        self.lbl_wh_wpp_info = tk.Label(
            self.frame_wh_whatsapp,
            text="As mensagens de fechamento e alertas de vazamento são enviadas diretamente para o WhatsApp cadastrado no perfil do operador.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#F0F7EE", anchor="w", wraplength=490, justify=tk.LEFT
        )
        self.lbl_wh_wpp_info.pack(fill=tk.X, pady=(1, 4))

        self.lbl_wh_wpp_contact = tk.Label(
            self.frame_wh_whatsapp,
            text="Operador: Carregando...",
            font=("Segoe UI", 8, "bold"), fg=COLOR_TEXT_MAIN, bg="#F0F7EE", anchor="w"
        )
        self.lbl_wh_wpp_contact.pack(fill=tk.X, pady=(0, 6))

        frame_wpp_btns = tk.Frame(self.frame_wh_whatsapp, bg="#F0F7EE")
        frame_wpp_btns.pack(fill=tk.X)

        btn_manage_wpp = create_btn_secondary(
            frame_wpp_btns, "👤 Alterar no Cadastro do Operador", self._open_operators_manager, pady=2, padx=10
        )
        btn_manage_wpp.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_whatsapp_help = tk.Button(
            frame_wpp_btns,
            text="📲 Como Ativar a Chave Grátis (30 seg)",
            command=self._open_callmebot_help,
            bg="#FFFFFF", fg=COLOR_PRIMARY, activebackground=COLOR_ACCENT,
            font=("Segoe UI", 8, "bold"), relief="flat", bd=1,
            highlightbackground="#C5DCBA", highlightthickness=1,
            padx=8, pady=2, cursor="hand2", takefocus=False
        )
        self.btn_whatsapp_help.pack(side=tk.LEFT)
        bind_button_hover(self.btn_whatsapp_help, "#FFFFFF", "#EBF3E6")

        # URL Webhook (Teams, Discord, Slack, Genérico)
        self.lbl_wh_url = tk.Label(self.card_wh_config, text="URL do Webhook:", font=("Segoe UI", 9), bg="#FFFFFF")
        self.ent_wh_url = ttk.Entry(self.card_wh_config, width=36)

        # Campos Telegram
        self.lbl_wh_tele_token = tk.Label(self.card_wh_config, text="Token do Bot:", font=("Segoe UI", 9), bg="#FFFFFF")
        self.ent_wh_tele_token = ttk.Entry(self.card_wh_config, width=36)

        self.lbl_wh_tele_chat = tk.Label(self.card_wh_config, text="Chat ID / Grupo:", font=("Segoe UI", 9), bg="#FFFFFF")
        self.ent_wh_tele_chat = ttk.Entry(self.card_wh_config, width=36)

        # Separador interno sutil
        sep_wh = tk.Frame(self.card_wh_config, height=1, bg="#E6EFE2")
        sep_wh.grid(row=6, column=0, columnspan=2, sticky=tk.EW, pady=(8, 6))

        # Opções adicionais
        self.var_wh_scheduled = tk.BooleanVar(value=True)
        chk_wh_sch = ttk.Checkbutton(
            self.card_wh_config, text="Disparar também em execuções automáticas do Agendador (dia 29)",
            variable=self.var_wh_scheduled
        )
        chk_wh_sch.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(2, 2))

        self.var_wh_anomalies = tk.BooleanVar(value=True)
        chk_wh_anom = ttk.Checkbutton(
            self.card_wh_config, text="Destacar alertas de suspeita de vazamento / anomalia na mensagem",
            variable=self.var_wh_anomalies
        )
        chk_wh_anom.grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(2, 2))

        # Card 3: Disparo de Teste
        card_test = create_card_frame(parent, padx=12, pady=8)
        card_test.pack(fill=tk.X)

        self.btn_test_webhook = create_btn_secondary(
            card_test, "🔔 Enviar Mensagem de Teste", self._test_webhook_action, pady=3, padx=12
        )
        self.btn_test_webhook.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_wh_test_status = tk.Label(
            card_test, text="", font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        self.lbl_wh_test_status.pack(side=tk.LEFT)

    def _update_wh_status_badge(self):
        if hasattr(self, "lbl_wh_badge"):
            if self.var_webhook_enabled.get():
                self.lbl_wh_badge.config(text="● ATIVO", fg="#2A6320", bg="#E8F4E5")
            else:
                self.lbl_wh_badge.config(text="○ INATIVO", fg=COLOR_TEXT_MUTED, bg="#F0F0F0")

    def _open_callmebot_help(self):
        url = "https://api.whatsapp.com/send?phone=34623758418&text=I%20allow%20callmebot%20to%20call%20me"
        webbrowser.open(url)
        messagebox.showinfo(
            "Como Ativar no WhatsApp (Grátis)",
            "Passo a passo rápido para receber no seu WhatsApp:\n\n"
            "1. Uma janela do WhatsApp foi aberta com o bot oficial CallMeBot (+34 623 75 84 18).\n"
            "2. Envie a mensagem pré-digitada: 'I allow callmebot to call me'.\n"
            "3. O bot responderá em segundos com sua Chave API (ApiKey: XXXXXX).\n"
            "4. Cadastre a Chave recebida no perfil do Operador (aba 'Operadores').\n"
            "5. Clique em '🔔 Enviar Mensagem de Teste' e pronto!\n\n"
            "100% gratuito e sem necessidade de cadastro.",
            parent=self
        )

    def _on_wh_platform_change(self, event=None):
        idx = self.cmb_wh_platform.current()
        key = self.wh_platform_keys[idx] if idx >= 0 else "whatsapp"

        # Esconde todos os campos dinâmicos
        self.lbl_wh_url.grid_remove()
        self.ent_wh_url.grid_remove()
        self.lbl_wh_tele_token.grid_remove()
        self.ent_wh_tele_token.grid_remove()
        self.lbl_wh_tele_chat.grid_remove()
        self.ent_wh_tele_chat.grid_remove()
        if hasattr(self, "frame_wh_whatsapp"):
            self.frame_wh_whatsapp.grid_remove()

        if key == "whatsapp":
            if hasattr(self, "frame_wh_whatsapp"):
                self.frame_wh_whatsapp.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(6, 4))
        elif key == "telegram":
            self.lbl_wh_tele_token.grid(row=2, column=0, sticky=tk.W, pady=3)
            self.ent_wh_tele_token.grid(row=2, column=1, sticky=tk.EW, pady=3)
            self.lbl_wh_tele_chat.grid(row=3, column=0, sticky=tk.W, pady=3)
            self.ent_wh_tele_chat.grid(row=3, column=1, sticky=tk.EW, pady=3)
        else:
            self.lbl_wh_url.grid(row=2, column=0, sticky=tk.W, pady=3)
            self.ent_wh_url.grid(row=2, column=1, sticky=tk.EW, pady=3)

    def _test_webhook_action(self):
        idx = self.cmb_wh_platform.current()
        platform = self.wh_platform_keys[idx] if idx >= 0 else "whatsapp"
        url = self.ent_wh_url.get().strip() if hasattr(self, "ent_wh_url") else ""
        token = self.ent_wh_tele_token.get().strip() if hasattr(self, "ent_wh_tele_token") else ""
        chat_id = self.ent_wh_tele_chat.get().strip() if hasattr(self, "ent_wh_tele_chat") else ""

        if platform == "whatsapp":
            from core.webhook_notifier import get_whatsapp_recipients
            recipients = get_whatsapp_recipients(self.config)
            if not recipients:
                messagebox.showwarning(
                    "WhatsApp Não Configurado",
                    "Nenhum operador possui telefone e Chave API CallMeBot cadastrada.\n\n"
                    "Clique em 'Alterar no Cadastro do Operador' para configurar o WhatsApp antes de testar.",
                    parent=self
                )
                return
            recip_names = ", ".join(r["name"] for r in recipients)
            self.lbl_wh_test_status.config(text=f"Enviando teste para {recip_names}...")
        else:
            self.lbl_wh_test_status.config(text="Enviando mensagem de teste...")

        self.btn_test_webhook.config(state=tk.DISABLED)

        def _worker():
            from core.webhook_notifier import send_test_webhook
            active_op = get_active_operator(self.config)
            op_name = active_op.get("name", "CompaSSS") if active_op else "CompaSSS"
            ok, msg = send_test_webhook(platform, url, token, chat_id, operator_name=op_name, config=self.config)
            def _ui():
                self.btn_test_webhook.config(state=tk.NORMAL)
                self.lbl_wh_test_status.config(text=msg)
                if ok:
                    messagebox.showinfo("Notificação OK", f"Mensagem de teste enviada com sucesso!\n\n{msg}", parent=self)
                else:
                    messagebox.showerror("Falha no Envio", f"Não foi possível enviar a mensagem:\n\n{msg}", parent=self)
            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _build_tab_backup(self, parent):
        """Constrói a aba de Backup e Restauração de configurações."""
        lbl_sec = tk.Label(
            parent, text="Central de Backup & Restauração de Dados",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            parent,
            text="Gere cópias completas de todos os parâmetros, operadores e modelos para transferir entre computadores ou recuperar com segurança.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, wraplength=540, justify=tk.LEFT
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 8))

        # Card 1: Exportar Backup
        card_exp = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#D5E2CF", highlightthickness=1, padx=14, pady=10)
        card_exp.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            card_exp, text="💾 1. Exportar Backup Completo",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 3))

        lbl_exp_info = tk.Label(
            card_exp,
            text="Cria um arquivo comprimido (.zip) contendo config.json, todos os perfis de operadores cadastrados, senhas salvas e modelo de e-mail.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#FFFFFF", wraplength=520, justify=tk.LEFT
        )
        lbl_exp_info.pack(anchor=tk.W, pady=(0, 6))

        btn_exp = create_btn_primary(
            card_exp, "💾 Criar e Exportar Backup (.zip)",
            self._export_backup_action,
            pady=5
        )
        btn_exp.pack(anchor=tk.W)

        # Card 2: Restaurar Backup
        card_imp = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#D5E2CF", highlightthickness=1, padx=14, pady=10)
        card_imp.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            card_imp, text="📂 2. Restaurar Backup de Configurações",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 3))

        lbl_imp_info = tk.Label(
            card_imp,
            text="Carrega as configurações de um arquivo .zip exportado anteriormente. O sistema cria automaticamente uma cópia de segurança antes de aplicar.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#FFFFFF", wraplength=520, justify=tk.LEFT
        )
        lbl_imp_info.pack(anchor=tk.W, pady=(0, 6))

        btn_imp = create_btn_secondary(
            card_imp, "📂 Selecionar Arquivo de Backup para Restaurar...",
            self._restore_backup_action,
            pady=5
        )
        btn_imp.pack(anchor=tk.W)

        # Card 3: Atalho da pasta de dados
        card_data = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#D5E2CF", highlightthickness=1, padx=14, pady=10)
        card_data.pack(fill=tk.X)

        tk.Label(
            card_data, text="📁 3. Pasta de Arquivos do Sistema",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 3))

        btn_open_data = create_btn_secondary(
            card_data, "📁 Abrir Pasta Raiz do Software no Windows Explorer",
            self._open_data_folder,
            pady=5
        )
        btn_open_data.pack(anchor=tk.W)

    def _export_backup_action(self):
        try:
            from core.backup_manager import create_backup
            default_name = f"SmartHydro_Backup_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.zip"
            target_path = filedialog.asksaveasfilename(
                parent=self,
                title="Salvar Arquivo de Backup",
                initialdir=os.path.join(os.path.expanduser("~"), "Desktop"),
                initialfile=default_name,
                filetypes=[("Backup do SmartHydro (*.zip)", "*.zip")]
            )
            if not target_path:
                return

            saved = create_backup(target_path)
            messagebox.showinfo(
                "Backup Realizado com Sucesso",
                f"Todas as configurações, operadores e modelos foram salvos com sucesso em:\n\n{saved}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Erro ao Gerar Backup", f"Falha ao criar arquivo de backup:\n{e}", parent=self)

    def _restore_backup_action(self):
        zip_file = filedialog.askopenfilename(
            parent=self,
            title="Selecionar Arquivo de Backup (.zip)",
            filetypes=[("Backup do SmartHydro (*.zip)", "*.zip")]
        )
        if not zip_file:
            return

        from core.backup_manager import read_backup_manifest, restore_backup
        ok, manifest, msg = read_backup_manifest(zip_file)
        if not ok:
            messagebox.showerror("Arquivo Inválido", msg, parent=self)
            return

        created = manifest.get("created_at", "Não informada")
        n_ops = manifest.get("operators_count", 0)
        ops_names = ", ".join(manifest.get("operators_names", []))

        confirm = messagebox.askyesno(
            "Confirmar Restauração",
            f"Deseja restaurar as configurações deste backup?\n\n"
            f"• Data do Backup: {created}\n"
            f"• Operadores ({n_ops}): {ops_names}\n"
            f"• Banco: {manifest.get('server', '-')}\n\n"
            f"Uma cópia de segurança do seu estado atual será criada automaticamente.",
            parent=self,
            icon="warning"
        )
        if not confirm:
            return

        res_ok, res_msg, new_cfg = restore_backup(zip_file)
        if res_ok:
            self.config = new_cfg
            self._load_values()
            self._refresh_tab_operators()
            if self.on_save_callback:
                self.on_save_callback(new_cfg)
            messagebox.showinfo("Restauração Concluída", f"{res_msg}\n\nAs telas foram atualizadas com os dados importados!", parent=self)
        else:
            messagebox.showerror("Erro ao Restaurar", res_msg, parent=self)

    def _open_data_folder(self):
        from core.config_manager import get_base_dir
        base = get_base_dir()
        try:
            os.startfile(base)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}", parent=self)

    def _build_tab_dev(self, parent):
        """Constrói a aba com informações de autoria e créditos do desenvolvedor."""
        card = tk.Frame(
            parent, bg="#FFFFFF",
            highlightbackground="#D5E5C9",
            highlightthickness=1,
            padx=18, pady=12
        )
        card.pack(fill=tk.BOTH, expand=True)

        lbl_avatar = tk.Label(card, text="💻", font=("Segoe UI Emoji", 26), bg="#FFFFFF")
        lbl_avatar.pack(pady=(0, 2))

        lbl_title = tk.Label(
            card, text="SmartHydro — Automação de Hidrômetros",
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        )
        lbl_title.pack()

        lbl_sub = tk.Label(
            card, text="Condomínio Praça Pamplona • Telemetria StruxureWare EBO",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_sub.pack(pady=(1, 4))

        badge_ver = tk.Label(
            card, text="★ Versão 3.1 • Edição Executiva CompaSSS",
            font=("Segoe UI", 8, "bold"), fg="#225E1A", bg="#E0F0D8",
            padx=8, pady=2
        )
        badge_ver.pack(pady=(0, 6))

        div = tk.Frame(card, height=1, bg="#D5E5C9")
        div.pack(fill=tk.X, pady=(0, 8))

        frame_info = tk.Frame(card, bg="#FFFFFF")
        frame_info.pack(fill=tk.X, padx=10)
        frame_info.columnconfigure(1, weight=1)

        info_items = [
            ("Desenvolvido por:", "Breno Camargo", True),
            ("E-mail / Contato:", "breno.hsc75@gmail.com", False),
            ("Empresa:", "CompaSSS Tecnologia e Automação", False),
            ("Empreendimento:", "Condomínio Praça Pamplona", False),
            ("Integração BMS:", "Schneider Electric StruxureWare EBO (SQL Server)", False),
            ("Linguagem & Motor:", "Python 3.11 • Tkinter • openpyxl", False),
            ("Versão:", "3.1 (Edição Executiva 2026)", False),
        ]

        for r_idx, (label, val, is_bold) in enumerate(info_items):
            lbl_l = tk.Label(
                frame_info, text=label,
                font=("Segoe UI", 9, "bold"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF",
                anchor="w"
            )
            lbl_l.grid(row=r_idx, column=0, sticky=tk.W, pady=2, padx=(0, 8))

            lbl_v = tk.Label(
                frame_info, text=val,
                font=("Segoe UI", 9, "bold" if is_bold else "normal"),
                fg=COLOR_PRIMARY if is_bold else COLOR_TEXT_MAIN,
                bg="#FFFFFF", anchor="w"
            )
            lbl_v.grid(row=r_idx, column=1, sticky=tk.W, pady=2)

        frame_actions = tk.Frame(card, bg="#FFFFFF")
        frame_actions.pack(pady=(10, 4))

        btn_github = create_btn_secondary(
            frame_actions, "🌐 GitHub",
            lambda: webbrowser.open("https://github.com/breno-camargo/Smarthydro"),
            padx=10, pady=5
        )
        btn_github.pack(side=tk.LEFT, padx=(0, 8))

        btn_copy_email = create_btn_secondary(
            frame_actions, "📋 Copiar E-mail",
            self._copy_dev_email,
            padx=10, pady=5
        )
        btn_copy_email.pack(side=tk.LEFT, padx=(0, 8))

        btn_shortcut = create_btn_secondary(
            frame_actions, "🖥️ Criar Atalho",
            self._create_desktop_shortcut_action,
            padx=10, pady=5
        )
        btn_shortcut.pack(side=tk.LEFT)

        lbl_badge = tk.Label(
            card,
            text="● Sistema Online & Conectado ao StruxureWare EBO",
            font=("Segoe UI", 7, "bold"), fg=COLOR_PRIMARY, bg="#EBF3E6",
            padx=8, pady=2
        )
        lbl_badge.pack(pady=(4, 2))

        lbl_copy = tk.Label(
            card,
            text="© 2026 Breno Camargo — Todos os direitos reservados.",
            font=("Segoe UI", 7, "italic"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_copy.pack(pady=(2, 0))

    def _copy_dev_email(self):
        self.clipboard_clear()
        self.clipboard_append("breno.hsc75@gmail.com")
        messagebox.showinfo(
            "Copiado!",
            "E-mail de contato copiado para a área de transferência:\n\nbreno.hsc75@gmail.com",
            parent=self
        )

    def _show_all_cc(self):
        cc_text = self.ent_cc.get().strip()
        if not cc_text:
            messagebox.showinfo("Cópias (Cc)", "Nenhum e-mail configurado em cópia padrão.", parent=self)
            return
        emails = [e.strip() for e in cc_text.split(";") if e.strip()]
        formatted = "\n".join([f"{idx+1}. {e}" for idx, e in enumerate(emails)])
        messagebox.showinfo(
            "Destinatários em Cópia Padrão (Cc)",
            f"Atualmente há {len(emails)} e-mail(s) cadastrados em cópia:\n\n{formatted}",
            parent=self
        )

    def _restore_default_emails(self):
        """Restaura instantaneamente os destinatários oficiais do Condomínio Praça Pamplona."""
        to_email, cc_emails = get_default_condominio_emails()
        self.ent_recipients.delete(0, tk.END)
        self.ent_recipients.insert(0, to_email)
        self.ent_cc.delete(0, tk.END)
        self.ent_cc.insert(0, cc_emails)
        messagebox.showinfo(
            "E-mails Padrão Restaurados",
            f"✔ Destinatários oficiais do Condomínio Praça Pamplona restaurados com sucesso!\n\n"
            f"Para (Gerente):\n• {to_email}\n\n"
            f"Em Cópia (Cc):\n" + "\n".join(f"• {e.strip()}" for e in cc_emails.split(";") if e.strip()),
            parent=self
        )

    def _on_smtp_preset_change(self, event=None):
        idx = self.cmb_smtp_preset.current()
        if idx == 0:
            # 465 SSL
            self.ent_smtp_port.delete(0, tk.END)
            self.ent_smtp_port.insert(0, "465")
            self.var_smtp_ssl.set(True)
            self.var_smtp_tls.set(False)
            self.frame_custom_port.grid_remove()
        elif idx == 1:
            # 587 STARTTLS
            self.ent_smtp_port.delete(0, tk.END)
            self.ent_smtp_port.insert(0, "587")
            self.var_smtp_ssl.set(False)
            self.var_smtp_tls.set(True)
            self.frame_custom_port.grid_remove()
        else:
            # Personalizado
            self.frame_custom_port.grid()

    def _open_email_template_editor(self):
        EmailTemplateDialog(self)

    def _open_scheduler(self):
        SchedulerDialog(self)

    def _create_desktop_shortcut_action(self):
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            desktop = shell.SpecialFolders("Desktop")
            shortcut_path = os.path.join(desktop, "SmartHydro - CompaSSS.lnk")
            lnk = shell.CreateShortCut(shortcut_path)

            exe_cand = [
                os.path.join(os.path.dirname(os.path.dirname(__file__)), "dist", "RelatorioHidrometros.exe"),
                os.path.join(os.getcwd(), "dist", "RelatorioHidrometros.exe"),
                os.path.join(os.getcwd(), "RelatorioHidrometros.exe"),
            ]
            exe_target = exe_cand[0]
            for c in exe_cand:
                if os.path.exists(c):
                    exe_target = c
                    break

            lnk.TargetPath = os.path.abspath(exe_target)
            lnk.WorkingDirectory = os.path.dirname(os.path.abspath(exe_target))
            ico_cand = [
                os.path.join(os.path.dirname(os.path.dirname(__file__)), "app_icon.ico"),
                os.path.join(os.getcwd(), "app_icon.ico"),
            ]
            for ic in ico_cand:
                if os.path.exists(ic):
                    lnk.IconLocation = os.path.abspath(ic) + ",0"
                    break
            lnk.Description = "SmartHydro - Medição de Água Praça Pamplona (CompaSSS)"
            lnk.save()
            messagebox.showinfo(
                "Atalho Criado",
                f"O atalho 'SmartHydro - CompaSSS' foi criado com sucesso na sua Área de Trabalho!\n\nArquivo:\n{shortcut_path}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Erro ao Criar Atalho", f"Não foi possível criar o atalho:\n{e}", parent=self)

    def _preview_email_action(self):
        try:
            subj_tmpl, body_tmpl = load_email_template()
            mock_context = {
                "mes": "Outubro",
                "ano": "2026",
                "periodo": "29/09/2026 a 28/10/2026",
                "total_m3": "46,40",
                "total_valor": "2.954,75",
                "valor_m3": "63,68",
                "qtd_salas": "289",
                "data_emissao": "01/10/2026 21:30"
            }
            _, final_html = render_email(subj_tmpl, body_tmpl, mock_context)
            final_html = prepare_html_for_preview(final_html)

            with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as f:
                f.write(final_html)
                temp_path = f.name
            webbrowser.open(f"file:///{temp_path}")
        except Exception as e:
            messagebox.showerror("Erro na Prévia", f"Não foi possível abrir a prévia no navegador:\n{e}", parent=self)

    def _open_template_in_notepad(self):
        path = get_email_template_path()
        try:
            subprocess.Popen(["notepad.exe", path])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}", parent=self)

    def _edit_template_action(self):
        try:
            path = open_template_in_excel()
            messagebox.showinfo(
                "Editar Modelo Excel",
                "O arquivo de modelo base foi aberto no Microsoft Excel!\n\n"
                "• Você pode trocar a logo, alterar cores, títulos, fontes ou bordas.\n"
                "• Mantenha o cabeçalho (linha 7) e a linha de exemplo (linha 8).\n"
                "• Ao terminar, basta salvar (Ctrl+S) e fechar o Excel.\n\n"
                "Os próximos relatórios seguirão exatamente as alterações feitas!",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Erro ao Abrir Modelo", f"Não foi possível abrir o arquivo de modelo:\n{e}", parent=self)

    def _toggle_auth(self):
        is_trusted = self.var_trusted.get()
        if is_trusted:
            self.frame_auth.grid_remove()
            self.lbl_trusted_hint.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        else:
            self.lbl_trusted_hint.grid_remove()
            self.frame_auth.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=2)
            self.ent_user.config(state=tk.NORMAL)
            self.ent_pass.config(state=tk.NORMAL)

    def _load_values(self):
        # Banco
        self.ent_server.insert(0, self.config.get("server", "WELLCARE-PC\\SQLEXPRESS"))
        self.ent_database.insert(0, self.config.get("database", "StruxureWareReportsDB"))

        driver = self.config.get("odbc_driver", "ODBC Driver 17 for SQL Server")
        if driver in self.cmb_driver["values"]:
            self.cmb_driver.set(driver)
        elif self.cmb_driver["values"]:
            self.cmb_driver.current(0)

        trusted = self.config.get("trusted_connection", True)
        self.var_trusted.set(trusted)
        self.ent_user.insert(0, self.config.get("db_user", ""))
        self.ent_pass.insert(0, self.config.get("db_password", ""))
        self._toggle_auth()

        self.ent_dir.insert(0, self.config.get("output_directory", ""))

        # Preferências de relatório
        self.var_open_excel.set(self.config.get("open_excel_after_generation", True))
        self.var_sort_consumption.set(self.config.get("sort_by_consumption", True))

        # E-mail
        self.ent_recipients.insert(0, self.config.get("email_recipients", ""))
        self.ent_cc.insert(0, self.config.get("email_cc", ""))

        self.ent_smtp_host.insert(0, self.config.get("smtp_server", "smtps.uhserver.com"))
        port = int(self.config.get("smtp_port", 465))
        use_ssl = bool(self.config.get("smtp_use_ssl", True))
        use_tls = bool(self.config.get("smtp_use_tls", False))

        self.ent_smtp_port.insert(0, str(port))
        self.var_smtp_ssl.set(use_ssl)
        self.var_smtp_tls.set(use_tls)

        if port == 465 and (use_ssl or not use_tls):
            self.cmb_smtp_preset.current(0)
            self.frame_custom_port.grid_remove()
        elif port == 587 and use_tls:
            self.cmb_smtp_preset.current(1)
            self.frame_custom_port.grid_remove()
        else:
            self.cmb_smtp_preset.current(2)
            self.frame_custom_port.grid()

        # Webhook
        self.var_webhook_enabled.set(self.config.get("webhook_enabled", False))
        curr_plat = self.config.get("webhook_platform", "whatsapp").lower()
        if curr_plat in self.wh_platform_keys:
            self.cmb_wh_platform.current(self.wh_platform_keys.index(curr_plat))
        else:
            self.cmb_wh_platform.current(0)
        self.ent_wh_url.delete(0, tk.END)
        self.ent_wh_url.insert(0, self.config.get("webhook_url", ""))
        self.ent_wh_tele_token.delete(0, tk.END)
        self.ent_wh_tele_token.insert(0, self.config.get("webhook_telegram_token", ""))
        self.ent_wh_tele_chat.delete(0, tk.END)
        self.ent_wh_tele_chat.insert(0, self.config.get("webhook_telegram_chat_id", ""))
        self.var_wh_scheduled.set(self.config.get("webhook_notify_scheduled", True))
        self.var_wh_anomalies.set(self.config.get("webhook_notify_anomalies", True))
        self._on_wh_platform_change()
        self._update_wh_status_badge()
        self._update_operator_context_labels()

    def _browse_dir(self):
        selected = filedialog.askdirectory(initialdir=self.ent_dir.get() or os.path.expanduser("~"))
        if selected:
            self.ent_dir.delete(0, tk.END)
            self.ent_dir.insert(0, os.path.abspath(selected))

    def _get_current_inputs_config(self):
        try:
            port_val = int(self.ent_smtp_port.get().strip())
        except ValueError:
            port_val = 465

        idx_plat = self.cmb_wh_platform.current()
        wh_plat = self.wh_platform_keys[idx_plat] if idx_plat >= 0 else "whatsapp"

        updated_cfg = self.config.copy()
        updated_cfg.update({
            "server": self.ent_server.get().strip(),
            "database": self.ent_database.get().strip(),
            "odbc_driver": self.cmb_driver.get().strip(),
            "trusted_connection": self.var_trusted.get(),
            "db_user": self.ent_user.get().strip(),
            "db_password": self.ent_pass.get().strip(),
            "output_directory": self.ent_dir.get().strip(),
            "open_excel_after_generation": self.var_open_excel.get(),
            "sort_by_consumption": self.var_sort_consumption.get(),
            # E-mail
            "email_recipients": self.ent_recipients.get().strip(),
            "email_cc": self.ent_cc.get().strip(),
            "email_send_mode": "smtp",
            "smtp_server": self.ent_smtp_host.get().strip(),
            "smtp_port": port_val,
            "smtp_use_tls": self.var_smtp_tls.get(),
            "smtp_use_ssl": self.var_smtp_ssl.get(),
            "smtp_user": self.config.get("smtp_user", ""),
            "smtp_password": self.config.get("smtp_password", ""),
            # Webhook & WhatsApp
            "webhook_enabled": self.var_webhook_enabled.get(),
            "webhook_platform": wh_plat,
            "webhook_whatsapp_phone": self.config.get("webhook_whatsapp_phone", ""),
            "webhook_whatsapp_apikey": self.config.get("webhook_whatsapp_apikey", ""),
            "webhook_url": self.ent_wh_url.get().strip() if hasattr(self, "ent_wh_url") else "",
            "webhook_telegram_token": self.ent_wh_tele_token.get().strip() if hasattr(self, "ent_wh_tele_token") else "",
            "webhook_telegram_chat_id": self.ent_wh_tele_chat.get().strip() if hasattr(self, "ent_wh_tele_chat") else "",
            "webhook_notify_scheduled": self.var_wh_scheduled.get(),
            "webhook_notify_anomalies": self.var_wh_anomalies.get(),
        })
        return updated_cfg

    def _test_connection(self):
        temp_cfg = self._get_current_inputs_config()

        def _worker():
            ok, msg, _ = test_db_connection(temp_cfg)
            def _ui():
                if ok:
                    messagebox.showinfo("Sucesso na Conexão", msg, parent=self)
                else:
                    messagebox.showerror("Erro de Conexão", msg, parent=self)
            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _test_smtp(self):
        temp_cfg = self._get_current_inputs_config()
        active_op = get_active_operator(self.config)
        if active_op:
            temp_cfg["smtp_user"] = (active_op.get("smtp_user") or active_op.get("email") or "").strip()
            temp_cfg["smtp_password"] = active_op.get("smtp_password", "").strip()

        op_name = active_op.get("name", "Operador") if active_op else "Operador"
        user_email = temp_cfg.get("smtp_user", "")
        if not user_email:
            messagebox.showwarning(
                "Aviso de Autenticação",
                f"O operador ativo '{op_name}' não possui e-mail cadastrado.\n\n"
                f"Clique em 'Gerenciar Operadores & Senhas' para configurar o e-mail de envio.",
                parent=self
            )
            return

        def _worker():
            ok, msg = test_smtp_connection(temp_cfg)
            def _ui():
                if ok:
                    messagebox.showinfo("Sucesso no SMTP", f"Conexão e autenticação SMTP bem-sucedidas para {op_name} ({user_email})!\n\n{msg}", parent=self)
                else:
                    messagebox.showerror("Erro no SMTP", f"Falha na conexão SMTP para {op_name} ({user_email}):\n\n{msg}", parent=self)
            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _save_and_close(self):
        new_cfg = self._get_current_inputs_config()
        if save_config(new_cfg):
            if self.on_save_callback:
                self.on_save_callback(new_cfg)
            messagebox.showinfo("Configurações Salvas", "As configurações foram atualizadas com sucesso!", parent=self)
            self.destroy()
        else:
            messagebox.showerror("Erro", "Não foi possível salvar as configurações no arquivo config.json.", parent=self)
