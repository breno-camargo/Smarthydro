import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from core.config_manager import load_config, save_config
from core.database import get_available_odbc_drivers, test_db_connection
from core.report_generator import open_template_in_excel

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("Configurações do Sistema e Conexão")
        self.geometry("540x490")
        self.minsize(520, 460)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self.on_save_callback = on_save_callback
        self.config = load_config()

        self._build_ui()
        self._load_values()

    def _build_ui(self):
        container = ttk.Frame(self, padding="20 16 20 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── SEÇÃO 1: CONEXÃO AO SQL SERVER ───
        lbl_sec_db = tk.Label(
            container, text="Conexão ao SQL Server (StruxureWare)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_db.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(container, text="Servidor/Instância:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.ent_server = ttk.Entry(container, width=38)
        self.ent_server.grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Banco de Dados:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.ent_database = ttk.Entry(container, width=38)
        self.ent_database.grid(row=2, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Driver ODBC:").grid(row=3, column=0, sticky=tk.W, pady=4)
        drivers = get_available_odbc_drivers()
        if not drivers:
            drivers = ["ODBC Driver 17 for SQL Server", "SQL Server"]
        self.cmb_driver = ttk.Combobox(container, values=drivers, state="readonly", width=36)
        self.cmb_driver.grid(row=3, column=1, sticky=tk.EW, pady=4)

        # Autenticação
        self.var_trusted = tk.BooleanVar(value=True)
        self.chk_trusted = ttk.Checkbutton(
            container, text="Usar Autenticação Integrada do Windows (Trusted)", 
            variable=self.var_trusted, command=self._toggle_auth
        )
        self.chk_trusted.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.frame_auth = ttk.Frame(container)
        self.frame_auth.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=2)

        ttk.Label(self.frame_auth, text="Usuário SQL:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ent_user = ttk.Entry(self.frame_auth, width=16)
        self.ent_user.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(self.frame_auth, text="Senha:").grid(row=0, column=2, sticky=tk.W, padx=(10, 0), pady=2)
        self.ent_pass = ttk.Entry(self.frame_auth, width=16, show="*")
        self.ent_pass.grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)

        # Separador 1
        sep1 = tk.Frame(container, height=1, bg=COLOR_ACCENT)
        sep1.grid(row=6, column=0, columnspan=2, sticky=tk.EW, pady=10)

        # ─── SEÇÃO 2: PREFERÊNCIAS DE RELATÓRIO ───
        lbl_sec_rep = tk.Label(
            container, text="Preferências de Relatório",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_rep.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(container, text="Pasta Padrão de Saída:").grid(row=8, column=0, sticky=tk.W, pady=4)
        frame_dir = ttk.Frame(container)
        frame_dir.grid(row=8, column=1, sticky=tk.EW, pady=4)
        self.ent_dir = ttk.Entry(frame_dir, width=28)
        self.ent_dir.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frame_dir, text="Alterar...", command=self._browse_dir, width=10).pack(side=tk.LEFT, padx=(5, 0))

        # Separador 2
        sep2 = tk.Frame(container, height=1, bg=COLOR_ACCENT)
        sep2.grid(row=9, column=0, columnspan=2, sticky=tk.EW, pady=10)

        # ─── SEÇÃO 3: MODELO DO RELATÓRIO EXCEL ───
        lbl_sec_tmpl = tk.Label(
            container, text="Personalização do Modelo Excel (Layout & Logo)",
            font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_sec_tmpl.grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        lbl_tmpl_desc = tk.Label(
            container,
            text="Abra o arquivo base no Excel para trocar o logotipo, ajustar cores, cabeçalhos\nou fontes. Os relatórios gerados seguirão o modelo salvo.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, justify=tk.LEFT
        )
        lbl_tmpl_desc.grid(row=11, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        btn_edit_model = ttk.Button(
            container, text="✏ Abrir Modelo no Excel para Edição",
            command=self._edit_template_action
        )
        btn_edit_model.grid(row=12, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))

        # Separador 3
        sep3 = tk.Frame(container, height=1, bg=COLOR_ACCENT)
        sep3.grid(row=13, column=0, columnspan=2, sticky=tk.EW, pady=10)

        # ─── BOTÕES DE AÇÃO INFERIORES ───
        frame_btns = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_btns.grid(row=14, column=0, columnspan=2, sticky=tk.EW, pady=(6, 0))

        btn_test = ttk.Button(frame_btns, text="🔌 Testar Conexão", command=self._test_connection)
        btn_test.pack(side=tk.LEFT)

        btn_save = tk.Button(
            frame_btns, text="Salvar Alterações", command=self._save_and_close,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=14, pady=5, cursor="hand2"
        )
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = ttk.Button(frame_btns, text="Cancelar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT)

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

    def _browse_dir(self):
        selected = filedialog.askdirectory(initialdir=self.ent_dir.get() or os.path.expanduser("~"))
        if selected:
            self.ent_dir.delete(0, tk.END)
            self.ent_dir.insert(0, os.path.abspath(selected))

    def _get_current_inputs_config(self):
        return {
            "server": self.ent_server.get().strip(),
            "database": self.ent_database.get().strip(),
            "odbc_driver": self.cmb_driver.get().strip(),
            "trusted_connection": self.var_trusted.get(),
            "db_user": self.ent_user.get().strip(),
            "db_password": self.ent_pass.get().strip(),
            "output_directory": self.ent_dir.get().strip(),
            "billing_cycle_type": self.config.get("billing_cycle_type", "ciclo_29_28"),
            "default_m3_price": self.config.get("default_m3_price", 63.68),
            "open_excel_after_generation": self.config.get("open_excel_after_generation", True),
            "sort_by_consumption": self.config.get("sort_by_consumption", True)
        }

    def _test_connection(self):
        temp_cfg = self._get_current_inputs_config()
        ok, msg, drv = test_db_connection(temp_cfg)
        if ok:
            messagebox.showinfo("Sucesso na Conexão", msg, parent=self)
        else:
            messagebox.showerror("Erro de Conexão", msg, parent=self)

    def _save_and_close(self):
        new_cfg = self._get_current_inputs_config()
        if save_config(new_cfg):
            if self.on_save_callback:
                self.on_save_callback(new_cfg)
            messagebox.showinfo("Configurações Salvas", "As configurações foram atualizadas com sucesso!", parent=self)
            self.destroy()
        else:
            messagebox.showerror("Erro", "Não foi possível salvar as configurações no arquivo config.json.", parent=self)
