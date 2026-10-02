import tkinter as tk
from tkinter import ttk, messagebox
import re

from core.config_manager import (
    load_config, get_operators, get_active_operator,
    set_active_operator, save_operator, delete_operator
)
from gui.ui_helpers import apply_window_icon, center_modal

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"


class OperatorsDialog(tk.Toplevel):
    def __init__(self, parent, on_change_callback=None):
        super().__init__(parent)
        self.title("Gerenciamento de Operadores & Assinaturas — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        self.on_change_callback = on_change_callback
        self.editing_op_id = None

        self._build_ui()
        self._load_operators_list()
        center_modal(self, parent, 580, 525)

    def _build_ui(self):
        main_box = ttk.Frame(self, padding="12 10 12 10")
        main_box.pack(fill=tk.BOTH, expand=True)

        # ─── HEADER ───
        lbl_head = tk.Label(
            main_box,
            text="👤 Perfis de Operador & Assinaturas de E-mail",
            font=("Segoe UI", 11, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_head.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            main_box,
            text="Cadastre os operadores que utilizam o sistema. Cada perfil possui seu próprio e-mail e assinatura corporativa.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, wraplength=540, justify=tk.LEFT
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 8))

        # ─── TABELA DE OPERADORES ───
        frame_table = tk.Frame(main_box, bg=COLOR_BG_LIGHT)
        frame_table.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columns = ("padrao", "nome", "cargo", "email", "telefone")
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=4, selectmode="browse")
        self.tree.heading("padrao", text="Status")
        self.tree.heading("nome", text="Nome Completo")
        self.tree.heading("cargo", text="Cargo / Função")
        self.tree.heading("email", text="E-mail de Envio")
        self.tree.heading("telefone", text="Telefone / Ramal")

        self.tree.column("padrao", width=65, anchor=tk.CENTER)
        self.tree.column("nome", width=120, anchor=tk.W)
        self.tree.column("cargo", width=115, anchor=tk.W)
        self.tree.column("email", width=145, anchor=tk.W)
        self.tree.column("telefone", width=95, anchor=tk.W)

        sb = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_select_operator)

        # Botões de ação da tabela
        frame_tbtns = tk.Frame(main_box, bg=COLOR_BG_LIGHT)
        frame_tbtns.pack(fill=tk.X, pady=(0, 8))

        btn_new = ttk.Button(frame_tbtns, text="➕ Novo Operador", command=self._start_new_operator)
        btn_new.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_set_default = ttk.Button(frame_tbtns, text="⭐ Definir como Padrão", command=self._set_as_default)
        self.btn_set_default.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_delete = ttk.Button(frame_tbtns, text="🗑️ Excluir", command=self._delete_selected)
        self.btn_delete.pack(side=tk.LEFT)

        # ─── FORMULÁRIO DE EDIÇÃO / CADASTRO ───
        self.frame_form = ttk.LabelFrame(main_box, text="  Dados do Operador Selecionado  ", padding="10 8 10 8")
        self.frame_form.pack(fill=tk.X, pady=(0, 8))

        # Grid 2 colunas
        lbl_n = ttk.Label(self.frame_form, text="Nome Completo:", font=("Segoe UI", 9, "bold"))
        lbl_n.grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_nome = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_nome.grid(row=0, column=1, sticky=tk.W, pady=3, padx=(4, 12))

        lbl_c = ttk.Label(self.frame_form, text="Cargo / Função:", font=("Segoe UI", 9, "bold"))
        lbl_c.grid(row=0, column=2, sticky=tk.W, pady=3)
        self.ent_cargo = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_cargo.grid(row=0, column=3, sticky=tk.W, pady=3, padx=(4, 0))

        lbl_e = ttk.Label(self.frame_form, text="E-mail Pessoal/Corp:", font=("Segoe UI", 9, "bold"))
        lbl_e.grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_email = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_email.grid(row=1, column=1, sticky=tk.W, pady=3, padx=(4, 12))

        lbl_t = ttk.Label(self.frame_form, text="Telefone / Contato:", font=("Segoe UI", 9, "bold"))
        lbl_t.grid(row=1, column=2, sticky=tk.W, pady=3)
        self.ent_tel = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_tel.grid(row=1, column=3, sticky=tk.W, pady=3, padx=(4, 0))

        lbl_u = ttk.Label(self.frame_form, text="Usuário SMTP:", font=("Segoe UI", 9))
        lbl_u.grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_smtp_user = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_smtp_user.grid(row=2, column=1, sticky=tk.W, pady=3, padx=(4, 12))

        lbl_p = ttk.Label(self.frame_form, text="Senha SMTP:", font=("Segoe UI", 9))
        lbl_p.grid(row=2, column=2, sticky=tk.W, pady=3)
        self.ent_smtp_pwd = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21, show="●")
        self.ent_smtp_pwd.grid(row=2, column=3, sticky=tk.W, pady=3, padx=(4, 0))

        self.var_default = tk.BooleanVar(value=False)
        self.chk_def = ttk.Checkbutton(
            self.frame_form,
            text="Operador Padrão (usado em agendamentos automáticos e inicialização)",
            variable=self.var_default
        )
        self.chk_def.grid(row=3, column=0, columnspan=4, sticky=tk.W, pady=(4, 4))

        frame_fbtns = tk.Frame(self.frame_form, bg=COLOR_BG_LIGHT)
        frame_fbtns.grid(row=4, column=0, columnspan=4, sticky=tk.E, pady=(2, 0))

        self.btn_save_op = tk.Button(
            frame_fbtns,
            text="💾 Salvar Operador",
            command=self._save_operator_form,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=12, pady=3, cursor="hand2"
        )
        self.btn_save_op.pack(side=tk.RIGHT)

        # ─── FOOTER ───
        frame_foot = tk.Frame(main_box, bg=COLOR_BG_LIGHT)
        frame_foot.pack(fill=tk.X, side=tk.BOTTOM, pady=(6, 0))

        btn_close = ttk.Button(frame_foot, text="Fechar", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _load_operators_list(self, select_id=None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        cfg = load_config()
        operators = get_operators(cfg)
        active_op = get_active_operator(cfg)
        active_id = active_op.get("id") if active_op else None

        target_iid = None
        for op in operators:
            op_id = op.get("id")
            is_def = op.get("is_default", False)
            status_txt = "★ Padrão" if is_def else ("● Ativo" if op_id == active_id else "")
            iid = self.tree.insert(
                "", tk.END,
                values=(status_txt, op.get("name", ""), op.get("role", ""), op.get("email", ""), op.get("phone", "")),
                tags=(op_id,)
            )
            if select_id and op_id == select_id:
                target_iid = iid
            elif not select_id and (op_id == active_id or is_def) and target_iid is None:
                target_iid = iid

        if target_iid:
            self.tree.selection_set(target_iid)
            self.tree.focus(target_iid)
            self._fill_form_with_selection()
        elif operators:
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self._fill_form_with_selection()

    def _on_select_operator(self, event=None):
        self._fill_form_with_selection()

    def _fill_form_with_selection(self):
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        tags = self.tree.item(item_id, "tags")
        if not tags:
            return
        op_id = tags[0]
        self.editing_op_id = op_id

        cfg = load_config()
        operators = get_operators(cfg)
        target = None
        for op in operators:
            if op.get("id") == op_id:
                target = op
                break
        if not target:
            return

        self.ent_nome.delete(0, tk.END)
        self.ent_nome.insert(0, target.get("name", ""))

        self.ent_cargo.delete(0, tk.END)
        self.ent_cargo.insert(0, target.get("role", ""))

        self.ent_email.delete(0, tk.END)
        self.ent_email.insert(0, target.get("email", ""))

        self.ent_tel.delete(0, tk.END)
        self.ent_tel.insert(0, target.get("phone", ""))

        self.ent_smtp_user.delete(0, tk.END)
        self.ent_smtp_user.insert(0, target.get("smtp_user", ""))

        self.ent_smtp_pwd.delete(0, tk.END)
        self.ent_smtp_pwd.insert(0, target.get("smtp_password", ""))

        self.var_default.set(bool(target.get("is_default", False)))
        self.frame_form.config(text=f"  Editando: {target.get('name', 'Operador')}  ")

    def _start_new_operator(self):
        self.editing_op_id = None
        self.tree.selection_remove(self.tree.selection())

        self.ent_nome.delete(0, tk.END)
        self.ent_cargo.delete(0, tk.END)
        self.ent_cargo.insert(0, "Técnico de Sistemas Prediais")
        self.ent_email.delete(0, tk.END)
        self.ent_tel.delete(0, tk.END)
        self.ent_tel.insert(0, "+55 11 ")
        self.ent_smtp_user.delete(0, tk.END)
        self.ent_smtp_pwd.delete(0, tk.END)
        self.var_default.set(False)

        self.frame_form.config(text="  Novo Operador  ")
        self.ent_nome.focus_set()

    def _save_operator_form(self):
        nome = self.ent_nome.get().strip()
        email = self.ent_email.get().strip()
        role = self.ent_cargo.get().strip() or "Técnico de Sistemas"
        phone = self.ent_tel.get().strip()
        smtp_u = self.ent_smtp_user.get().strip() or email
        smtp_p = self.ent_smtp_pwd.get().strip()
        is_def = self.var_default.get()

        if not nome:
            messagebox.showwarning("Aviso", "Por favor, informe o Nome Completo do operador.", parent=self)
            self.ent_nome.focus_set()
            return
        if not email or "@" not in email:
            messagebox.showwarning("Aviso", "Por favor, informe um endereço de e-mail válido.", parent=self)
            self.ent_email.focus_set()
            return

        op_data = {
            "name": nome,
            "role": role,
            "email": email,
            "phone": phone,
            "smtp_user": smtp_u,
            "smtp_password": smtp_p,
            "is_default": is_def
        }
        if self.editing_op_id:
            op_data["id"] = self.editing_op_id

        saved_id = save_operator(op_data)
        self._load_operators_list(select_id=saved_id)

        if self.on_change_callback:
            try:
                self.on_change_callback()
            except Exception:
                pass

        messagebox.showinfo("Sucesso", f"Operador '{nome}' salvo com sucesso!", parent=self)

    def _set_as_default(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Informação", "Selecione um operador na lista acima.", parent=self)
            return
        item_id = sel[0]
        op_id = self.tree.item(item_id, "tags")[0]

        cfg = load_config()
        ops = get_operators(cfg)
        target = None
        for op in ops:
            if op.get("id") == op_id:
                op["is_default"] = True
                target = op
            else:
                op["is_default"] = False

        cfg["operators"] = ops
        cfg["active_operator_id"] = op_id
        from core.config_manager import save_config
        save_config(cfg)

        self._load_operators_list(select_id=op_id)
        if self.on_change_callback:
            try:
                self.on_change_callback()
            except Exception:
                pass

        messagebox.showinfo(
            "Operador Padrão",
            f"'{target.get('name')}' foi definido como o operador padrão do sistema.",
            parent=self
        )

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Informação", "Selecione um operador na lista acima para excluir.", parent=self)
            return
        item_id = sel[0]
        op_id = self.tree.item(item_id, "tags")[0]

        cfg = load_config()
        ops = get_operators(cfg)
        if len(ops) <= 1:
            messagebox.showwarning("Aviso", "Não é possível excluir o único operador cadastrado no sistema.", parent=self)
            return

        target_name = ""
        for op in ops:
            if op.get("id") == op_id:
                target_name = op.get("name", "Operador")
                break

        if not messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o operador '{target_name}'?", parent=self):
            return

        ok, msg = delete_operator(op_id)
        if ok:
            self._load_operators_list()
            if self.on_change_callback:
                try:
                    self.on_change_callback()
                except Exception:
                    pass
            messagebox.showinfo("Sucesso", msg, parent=self)
        else:
            messagebox.showerror("Erro", msg, parent=self)
