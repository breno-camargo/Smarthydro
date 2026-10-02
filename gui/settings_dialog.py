import os
import subprocess
import threading
import tempfile
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core.config_manager import load_config, save_config
from core.database import get_available_odbc_drivers, test_db_connection
from core.report_generator import open_template_in_excel
from core.email_sender import (
    test_smtp_connection, get_email_template_path, load_email_template,
    render_email, prepare_html_for_preview, DEFAULT_SUBJECT_TEMPLATE
)
from gui.email_template_dialog import EmailTemplateDialog
from gui.scheduler_dialog import SchedulerDialog
from gui.ui_helpers import apply_window_icon, center_modal

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"

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

        # Centralizar perfeitamente sobre a janela principal (570x530 sobre 640x600 = 35px simétrico)
        center_modal(self, parent, 570, 530)

    def _configure_notebook_style(self):
        self.style = ttk.Style(self)
        self.style.configure("Settings.TNotebook", background=COLOR_BG_LIGHT, borderwidth=0)
        self.style.configure(
            "Settings.TNotebook.Tab",
            font=("Segoe UI", 9, "bold"),
            padding=[14, 6]
        )

    def _build_ui(self):
        container = ttk.Frame(self, padding="14 10 14 10")
        container.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(container, style="Settings.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # ─── ABA 1: BANCO DE DADOS & EXCEL ───
        tab_db = ttk.Frame(self.notebook, padding="14 10 14 10")
        self.notebook.add(tab_db, text=" 🏢 Banco & Relatórios ")
        self._build_tab_db(tab_db)

        # ─── ABA 2: E-MAIL & UOL PRO ───
        tab_email = ttk.Frame(self.notebook, padding="14 10 14 10")
        self.notebook.add(tab_email, text=" ✉ E-mail & UOL Pro ")
        self._build_tab_email(tab_email)

        # ─── ABA 3: DESENVOLVEDOR ───
        tab_dev = ttk.Frame(self.notebook, padding="14 10 14 10")
        self.notebook.add(tab_dev, text=" 💻 Desenvolvedor ")
        self._build_tab_dev(tab_dev)

        # ─── BOTÕES DE AÇÃO INFERIORES ───
        frame_btns = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_btns.pack(fill=tk.X, pady=(2, 0))

        btn_save = tk.Button(
            frame_btns, text="Salvar Alterações", command=self._save_and_close,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=16, pady=5, cursor="hand2", takefocus=False
        )
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = ttk.Button(frame_btns, text="Cancelar", command=self.destroy, takefocus=False)
        btn_cancel.pack(side=tk.RIGHT)

    def _build_tab_db(self, parent):
        parent.columnconfigure(1, weight=1)

        # ─── SEÇÃO 1: CONEXÃO AO SQL SERVER ───
        lbl_sec_db = tk.Label(
            parent, text="Conexão ao SQL Server (StruxureWare)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_db.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        ttk.Label(parent, text="Servidor/Instância:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_server = ttk.Entry(parent, width=36)
        self.ent_server.grid(row=1, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Banco de Dados:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_database = ttk.Entry(parent, width=36)
        self.ent_database.grid(row=2, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Driver ODBC:").grid(row=3, column=0, sticky=tk.W, pady=3)
        drivers = get_available_odbc_drivers()
        if not drivers:
            drivers = ["ODBC Driver 17 for SQL Server", "SQL Server"]
        self.cmb_driver = ttk.Combobox(parent, values=drivers, state="readonly", width=34)
        self.cmb_driver.grid(row=3, column=1, sticky=tk.EW, pady=3)

        # Autenticação
        self.var_trusted = tk.BooleanVar(value=True)
        self.chk_trusted = ttk.Checkbutton(
            parent, text="Usar Autenticação Integrada do Windows (Trusted)",
            variable=self.var_trusted, command=self._toggle_auth
        )
        self.chk_trusted.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=4)

        self.lbl_trusted_hint = tk.Label(
            parent,
            text="🔒 Autenticação integrada do Windows ativa (usuário e senha dispensados).",
            font=("Segoe UI", 8, "italic"),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG_LIGHT
        )

        self.frame_auth = ttk.Frame(parent)
        self.frame_auth.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=2)

        ttk.Label(self.frame_auth, text="Usuário:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ent_user = ttk.Entry(self.frame_auth, width=15)
        self.ent_user.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(self.frame_auth, text="Senha:").grid(row=0, column=2, sticky=tk.W, padx=(12, 0), pady=2)
        self.ent_pass = ttk.Entry(self.frame_auth, width=15, show="*")
        self.ent_pass.grid(row=0, column=3, sticky=tk.W, padx=4, pady=2)

        btn_test_db = ttk.Button(parent, text="🔌 Testar Conexão SQL", command=self._test_connection)
        btn_test_db.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(3, 6))

        # Separador 1
        sep1 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep1.grid(row=7, column=0, columnspan=2, sticky=tk.EW, pady=4)

        # ─── SEÇÃO 2: PASTA PADRÃO & MODELO EXCEL ───
        lbl_sec_rep = tk.Label(
            parent, text="Pasta Padrão & Modelo Excel",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_rep.grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        ttk.Label(parent, text="Pasta de Saída:").grid(row=9, column=0, sticky=tk.W, pady=3)
        frame_dir = ttk.Frame(parent)
        frame_dir.grid(row=9, column=1, sticky=tk.EW, pady=3)
        self.ent_dir = ttk.Entry(frame_dir, width=24)
        self.ent_dir.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frame_dir, text="Alterar...", command=self._browse_dir, width=9).pack(side=tk.LEFT, padx=(4, 0))

        btn_edit_model = ttk.Button(
            parent, text="✏ Abrir Modelo no Excel para Edição (modelo_relatorio.xlsx)",
            command=self._edit_template_action
        )
        btn_edit_model.grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(4, 6))

        # Separador 2
        sep2 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep2.grid(row=11, column=0, columnspan=2, sticky=tk.EW, pady=4)

        # ─── SEÇÃO 3: PREFERÊNCIAS DO RELATÓRIO ───
        lbl_sec_pref = tk.Label(
            parent, text="Preferências Gerais do Relatório",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_pref.grid(row=12, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        self.var_open_excel = tk.BooleanVar(value=True)
        self.chk_open_excel = ttk.Checkbutton(
            parent, text="Abrir planilha no Microsoft Excel automaticamente após geração",
            variable=self.var_open_excel
        )
        self.chk_open_excel.grid(row=13, column=0, columnspan=2, sticky=tk.W, pady=2)

        self.var_sort_consumption = tk.BooleanVar(value=True)
        self.chk_sort_consumption = ttk.Checkbutton(
            parent, text="Ordenar unidades por maior consumo no relatório (ranking decrescente)",
            variable=self.var_sort_consumption
        )
        self.chk_sort_consumption.grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=2)

        btn_scheduler = ttk.Button(
            parent,
            text="⏰ Configurar Agendamento Automático no Windows (Executar dia 29 silencioso)",
            command=self._open_scheduler
        )
        btn_scheduler.grid(row=15, column=0, columnspan=2, sticky=tk.W, pady=(6, 2))

    def _build_tab_email(self, parent):
        parent.columnconfigure(1, weight=1)

        # ─── SEÇÃO 1: DESTINATÁRIOS E MÉTODO ───
        lbl_sec_em = tk.Label(
            parent, text="Destinatários Padrão & Método de Envio",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_em.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        ttk.Label(parent, text="E-mail da Gerente (Para):").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_recipients = ttk.Entry(parent, width=36)
        self.ent_recipients.grid(row=1, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Em Cópia Padrão (Cc):").grid(row=2, column=0, sticky=tk.W, pady=3)
        frame_cc = ttk.Frame(parent)
        frame_cc.grid(row=2, column=1, sticky=tk.EW, pady=3)

        self.ent_cc = ttk.Entry(frame_cc, width=28)
        self.ent_cc.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_view_cc = ttk.Button(
            frame_cc, text="🔍 Ver Todos", width=10, command=self._show_all_cc
        )
        btn_view_cc.pack(side=tk.LEFT, padx=(4, 0))

        lbl_hint_to = tk.Label(
            parent, text="Separe múltiplos e-mails por ponto-e-vírgula (;)",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_hint_to.grid(row=3, column=1, sticky=tk.W, pady=(0, 4))

        # Separador 1
        sep_e1 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep_e1.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=6)

        # ─── SEÇÃO 2: SERVIDOR SMTP (UOL PRO) ───
        lbl_sec_smtp = tk.Label(
            parent, text="Sua Conta de E-mail (UOL Pro / Remetente)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_smtp.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        ttk.Label(parent, text="Servidor SMTP:").grid(row=6, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_host = ttk.Entry(parent, width=36)
        self.ent_smtp_host.grid(row=6, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Porta / Criptografia:").grid(row=7, column=0, sticky=tk.W, pady=3)
        frame_port_sec = ttk.Frame(parent)
        frame_port_sec.grid(row=7, column=1, sticky=tk.W, pady=3)

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

        # Painel para configuração personalizada (oculto quando usando presets padrão)
        self.frame_custom_port = ttk.Frame(parent)
        self.frame_custom_port.grid(row=8, column=1, sticky=tk.W, pady=(2, 4))

        ttk.Label(self.frame_custom_port, text="Porta:").pack(side=tk.LEFT)
        self.ent_smtp_port = ttk.Entry(self.frame_custom_port, width=6)
        self.ent_smtp_port.pack(side=tk.LEFT, padx=(4, 10))

        self.var_smtp_tls = tk.BooleanVar(value=False)
        self.chk_smtp_tls = ttk.Checkbutton(self.frame_custom_port, text="STARTTLS", variable=self.var_smtp_tls)
        self.chk_smtp_tls.pack(side=tk.LEFT, padx=(0, 8))

        self.var_smtp_ssl = tk.BooleanVar(value=True)
        self.chk_smtp_ssl = ttk.Checkbutton(self.frame_custom_port, text="SSL Direto", variable=self.var_smtp_ssl)
        self.chk_smtp_ssl.pack(side=tk.LEFT)

        ttk.Label(parent, text="Seu E-mail (Remetente):").grid(row=9, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_user = ttk.Entry(parent, width=36)
        self.ent_smtp_user.grid(row=9, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Senha do seu E-mail:").grid(row=10, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_pass = ttk.Entry(parent, width=36, show="*")
        self.ent_smtp_pass.grid(row=10, column=1, sticky=tk.EW, pady=3)

        btn_test_smtp = ttk.Button(parent, text="🔌 Testar Conexão com seu E-mail", command=self._test_smtp)
        btn_test_smtp.grid(row=11, column=0, columnspan=2, sticky=tk.W, pady=(4, 6))

        # Separador 2
        sep_e2 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep_e2.grid(row=12, column=0, columnspan=2, sticky=tk.EW, pady=4)

        # ─── SEÇÃO 3: MODELO DE E-MAIL ───
        lbl_sec_tmpl_e = tk.Label(
            parent, text="Personalização do Modelo de E-mail",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_tmpl_e.grid(row=13, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        frame_tmpl_btns = ttk.Frame(parent)
        frame_tmpl_btns.grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))

        btn_edit_email_tmpl = ttk.Button(
            frame_tmpl_btns, text="✏ Personalizar Modelo",
            command=self._open_email_template_editor
        )
        btn_edit_email_tmpl.pack(side=tk.LEFT, padx=(0, 6))

        btn_preview_tmpl = ttk.Button(
            frame_tmpl_btns, text="👁 Ver Prévia no Navegador",
            command=self._preview_email_action
        )
        btn_preview_tmpl.pack(side=tk.LEFT, padx=(0, 6))

        btn_open_tmpl_file = ttk.Button(
            frame_tmpl_btns, text="📂 Bloco de Notas",
            command=self._open_template_in_notepad
        )
        btn_open_tmpl_file.pack(side=tk.LEFT)

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
        lbl_sub.pack(pady=(1, 6))

        div = tk.Frame(card, height=1, bg="#D5E5C9")
        div.pack(fill=tk.X, pady=(0, 8))

        frame_info = tk.Frame(card, bg="#FFFFFF")
        frame_info.pack(fill=tk.X, padx=10)
        frame_info.columnconfigure(1, weight=1)

        info_items = [
            ("Desenvolvido por:", "Breno Camargo", True),
            ("E-mail:", "breno.camargo@compasss.com.br", False),
            ("Empresa:", "CompaSSS Tecnologia e Automação", False),
            ("Empreendimento:", "Condomínio Praça Pamplona", False),
            ("Integração BMS:", "Schneider Electric StruxureWare EBO (SQL Server)", False),
            ("Linguagem & Motor:", "Python 3.11 • Tkinter • openpyxl", False),
            ("Versão:", "2.1 (Edição Executiva 2026)", False),
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

        btn_github = tk.Button(
            frame_actions,
            text="🌐 Repositório no GitHub",
            command=lambda: webbrowser.open("https://github.com/breno-camargo/Smarthydro"),
            bg="#EBF3E6", fg=COLOR_PRIMARY, activebackground=COLOR_ACCENT,
            font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5,
            cursor="hand2", takefocus=False
        )
        btn_github.pack(side=tk.LEFT, padx=(0, 8))

        btn_copy_email = tk.Button(
            frame_actions,
            text="📋 Copiar E-mail",
            command=self._copy_dev_email,
            bg="#EBF3E6", fg=COLOR_PRIMARY, activebackground=COLOR_ACCENT,
            font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=5,
            cursor="hand2", takefocus=False
        )
        btn_copy_email.pack(side=tk.LEFT, padx=(0, 8))

        btn_shortcut = tk.Button(
            frame_actions,
            text="🖥️ Criar Atalho na Área de Trabalho",
            command=self._create_desktop_shortcut_action,
            bg="#EBF3E6", fg=COLOR_PRIMARY, activebackground=COLOR_ACCENT,
            font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=5,
            cursor="hand2", takefocus=False
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
        self.clipboard_append("breno.camargo@compasss.com.br")
        messagebox.showinfo(
            "Copiado!",
            "E-mail de contato copiado para a área de transferência:\n\nbreno.camargo@compasss.com.br",
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

        self.ent_smtp_user.insert(0, self.config.get("smtp_user", ""))
        self.ent_smtp_pass.insert(0, self.config.get("smtp_password", ""))

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
            "smtp_user": self.ent_smtp_user.get().strip(),
            "smtp_password": self.ent_smtp_pass.get().strip(),
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

        def _worker():
            ok, msg = test_smtp_connection(temp_cfg)
            def _ui():
                if ok:
                    messagebox.showinfo("Sucesso no SMTP", msg, parent=self)
                else:
                    messagebox.showerror("Erro no SMTP", msg, parent=self)
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
