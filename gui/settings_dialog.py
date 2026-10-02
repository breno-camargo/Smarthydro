import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core.config_manager import load_config, save_config
from core.database import get_available_odbc_drivers, test_db_connection
from core.report_generator import open_template_in_excel
from core.email_sender import test_smtp_connection, get_email_template_path
from gui.email_template_dialog import EmailTemplateDialog

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MUTED = "#55664C"

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("Configurações do Sistema — CompaSSS")
        self.geometry("560x540")
        self.minsize(530, 500)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self.on_save_callback = on_save_callback
        self.config = load_config()

        self._build_ui()
        self._load_values()

    def _build_ui(self):
        container = ttk.Frame(self, padding="16 12 16 12")
        container.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # ─── ABA 1: BANCO DE DADOS & EXCEL ───
        tab_db = ttk.Frame(self.notebook, padding="14 12 14 12")
        self.notebook.add(tab_db, text=" 🏢 Banco & Relatórios ")
        self._build_tab_db(tab_db)

        # ─── ABA 2: E-MAIL & UOL PRO ───
        tab_email = ttk.Frame(self.notebook, padding="14 12 14 12")
        self.notebook.add(tab_email, text=" ✉ E-mail & UOL Pro ")
        self._build_tab_email(tab_email)

        # ─── BOTÕES DE AÇÃO INFERIORES ───
        frame_btns = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_btns.pack(fill=tk.X, pady=(4, 0))

        btn_save = tk.Button(
            frame_btns, text="Salvar Alterações", command=self._save_and_close,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=16, pady=5, cursor="hand2"
        )
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = ttk.Button(frame_btns, text="Cancelar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT)

    def _build_tab_db(self, parent):
        parent.columnconfigure(1, weight=1)

        # ─── SEÇÃO 1: CONEXÃO AO SQL SERVER ───
        lbl_sec_db = tk.Label(
            parent, text="Conexão ao SQL Server (StruxureWare)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_db.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(parent, text="Servidor/Instância:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.ent_server = ttk.Entry(parent, width=36)
        self.ent_server.grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(parent, text="Banco de Dados:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.ent_database = ttk.Entry(parent, width=36)
        self.ent_database.grid(row=2, column=1, sticky=tk.EW, pady=4)

        ttk.Label(parent, text="Driver ODBC:").grid(row=3, column=0, sticky=tk.W, pady=4)
        drivers = get_available_odbc_drivers()
        if not drivers:
            drivers = ["ODBC Driver 17 for SQL Server", "SQL Server"]
        self.cmb_driver = ttk.Combobox(parent, values=drivers, state="readonly", width=34)
        self.cmb_driver.grid(row=3, column=1, sticky=tk.EW, pady=4)

        # Autenticação
        self.var_trusted = tk.BooleanVar(value=True)
        self.chk_trusted = ttk.Checkbutton(
            parent, text="Usar Autenticação Integrada do Windows (Trusted)",
            variable=self.var_trusted, command=self._toggle_auth
        )
        self.chk_trusted.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.frame_auth = ttk.Frame(parent)
        self.frame_auth.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=2)

        ttk.Label(self.frame_auth, text="Usuário:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ent_user = ttk.Entry(self.frame_auth, width=14)
        self.ent_user.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(self.frame_auth, text="Senha:").grid(row=0, column=2, sticky=tk.W, padx=(10, 0), pady=2)
        self.ent_pass = ttk.Entry(self.frame_auth, width=14, show="*")
        self.ent_pass.grid(row=0, column=3, sticky=tk.W, padx=4, pady=2)

        btn_test_db = ttk.Button(parent, text="🔌 Testar Conexão SQL", command=self._test_connection)
        btn_test_db.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(4, 8))

        # Separador 1
        sep1 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep1.grid(row=7, column=0, columnspan=2, sticky=tk.EW, pady=6)

        # ─── SEÇÃO 2: PREFERÊNCIAS DE RELATÓRIO ───
        lbl_sec_rep = tk.Label(
            parent, text="Pasta Padrão & Modelo Excel",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_rep.grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(parent, text="Pasta de Saída:").grid(row=9, column=0, sticky=tk.W, pady=4)
        frame_dir = ttk.Frame(parent)
        frame_dir.grid(row=9, column=1, sticky=tk.EW, pady=4)
        self.ent_dir = ttk.Entry(frame_dir, width=24)
        self.ent_dir.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frame_dir, text="Alterar...", command=self._browse_dir, width=9).pack(side=tk.LEFT, padx=(4, 0))

        btn_edit_model = ttk.Button(
            parent, text="✏ Abrir Modelo no Excel para Edição (modelo_relatorio.xlsx)",
            command=self._edit_template_action
        )
        btn_edit_model.grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))

    def _build_tab_email(self, parent):
        parent.columnconfigure(1, weight=1)

        # ─── SEÇÃO 1: DESTINATÁRIOS E MÉTODO ───
        lbl_sec_em = tk.Label(
            parent, text="Destinatários Padrão & Método de Envio",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_em.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(parent, text="E-mail da Gerente (Para):").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.ent_recipients = ttk.Entry(parent, width=36)
        self.ent_recipients.grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(parent, text="Em Cópia Padrão (Cc):").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.ent_cc = ttk.Entry(parent, width=36)
        self.ent_cc.grid(row=2, column=1, sticky=tk.EW, pady=4)

        lbl_hint_to = tk.Label(
            parent, text="Separe múltiplos e-mails por ponto-e-vírgula (;)",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_hint_to.grid(row=3, column=1, sticky=tk.W, pady=(0, 6))



        # Separador
        sep_e1 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep_e1.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=8)

        # ─── SEÇÃO 2: SERVIDOR SMTP (UOL PRO) ───
        lbl_sec_smtp = tk.Label(
            parent, text="Sua Conta de E-mail (UOL Pro / Remetente)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_smtp.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(parent, text="Servidor SMTP:").grid(row=7, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_host = ttk.Entry(parent, width=36)
        self.ent_smtp_host.grid(row=7, column=1, sticky=tk.EW, pady=3)

        frame_port = ttk.Frame(parent)
        frame_port.grid(row=8, column=1, sticky=tk.W, pady=3)

        ttk.Label(parent, text="Porta SMTP:").grid(row=8, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_port = ttk.Entry(frame_port, width=8)
        self.ent_smtp_port.pack(side=tk.LEFT)

        self.var_smtp_tls = tk.BooleanVar(value=True)
        self.chk_smtp_tls = ttk.Checkbutton(frame_port, text="STARTTLS (recomendado)", variable=self.var_smtp_tls)
        self.chk_smtp_tls.pack(side=tk.LEFT, padx=(12, 0))

        ttk.Label(parent, text="Seu E-mail (Remetente):").grid(row=9, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_user = ttk.Entry(parent, width=36)
        self.ent_smtp_user.grid(row=9, column=1, sticky=tk.EW, pady=3)

        ttk.Label(parent, text="Senha do seu E-mail:").grid(row=10, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_pass = ttk.Entry(parent, width=36, show="*")
        self.ent_smtp_pass.grid(row=10, column=1, sticky=tk.EW, pady=3)

        btn_test_smtp = ttk.Button(parent, text="🔌 Testar Conexão com seu E-mail", command=self._test_smtp)
        btn_test_smtp.grid(row=11, column=0, columnspan=2, sticky=tk.W, pady=(6, 8))

        # Separador
        sep_e2 = tk.Frame(parent, height=1, bg=COLOR_ACCENT)
        sep_e2.grid(row=12, column=0, columnspan=2, sticky=tk.EW, pady=6)

        # ─── SEÇÃO 3: MODELO DE E-MAIL ───
        lbl_sec_tmpl_e = tk.Label(
            parent, text="Personalização do Modelo de E-mail",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_tmpl_e.grid(row=13, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        frame_tmpl_btns = ttk.Frame(parent)
        frame_tmpl_btns.grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))

        btn_edit_email_tmpl = ttk.Button(
            frame_tmpl_btns, text="✏ Personalizar Modelo de E-mail (Tags & Texto)",
            command=self._open_email_template_editor
        )
        btn_edit_email_tmpl.pack(side=tk.LEFT, padx=(0, 8))

        btn_open_tmpl_file = ttk.Button(
            frame_tmpl_btns, text="📂 Abrir no Bloco de Notas",
            command=self._open_template_in_notepad
        )
        btn_open_tmpl_file.pack(side=tk.LEFT)

    def _open_email_template_editor(self):
        EmailTemplateDialog(self)

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
        state = tk.DISABLED if is_trusted else tk.NORMAL
        self.ent_user.config(state=state)
        self.ent_pass.config(state=state)

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

        # E-mail
        self.ent_recipients.insert(0, self.config.get("email_recipients", ""))
        self.ent_cc.insert(0, self.config.get("email_cc", ""))

        self.ent_smtp_host.insert(0, self.config.get("smtp_server", "smtps.uhserver.com"))
        self.ent_smtp_port.insert(0, str(self.config.get("smtp_port", 465)))
        self.var_smtp_tls.set(self.config.get("smtp_use_tls", False))
        self.ent_smtp_user.insert(0, self.config.get("smtp_user", ""))
        self.ent_smtp_pass.insert(0, self.config.get("smtp_password", ""))

    def _browse_dir(self):
        selected = filedialog.askdirectory(initialdir=self.ent_dir.get() or os.path.expanduser("~"))
        if selected:
            self.ent_dir.delete(0, tk.END)
            self.ent_dir.insert(0, os.path.abspath(selected))

    def _get_current_inputs_config(self):
        mode_val = "smtp"

        try:
            port_val = int(self.ent_smtp_port.get().strip())
        except ValueError:
            port_val = 465

        # Preserva integralmente chaves existentes (recent_reports, report_send_log, etc.)
        updated_cfg = self.config.copy()
        updated_cfg.update({
            "server": self.ent_server.get().strip(),
            "database": self.ent_database.get().strip(),
            "odbc_driver": self.cmb_driver.get().strip(),
            "trusted_connection": self.var_trusted.get(),
            "db_user": self.ent_user.get().strip(),
            "db_password": self.ent_pass.get().strip(),
            "output_directory": self.ent_dir.get().strip(),
            # E-mail
            "email_recipients": self.ent_recipients.get().strip(),
            "email_cc": self.ent_cc.get().strip(),
            "email_send_mode": mode_val,
            "smtp_server": self.ent_smtp_host.get().strip(),
            "smtp_port": port_val,
            "smtp_use_tls": self.var_smtp_tls.get(),
            "smtp_use_ssl": False,
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
