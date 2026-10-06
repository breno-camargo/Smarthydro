import tkinter as tk
from tkinter import ttk, messagebox
import re
import threading

from core.config_manager import (
    load_config, get_operators, get_active_operator,
    set_active_operator, save_operator, delete_operator
)
from gui.ui_helpers import (
    apply_window_icon, center_modal, setup_common_styles,
    create_btn_primary, create_btn_secondary, create_btn_danger,
    COLOR_PRIMARY, COLOR_PRIMARY_HOVER, COLOR_ACCENT, COLOR_BG_LIGHT,
    COLOR_TEXT_MAIN, COLOR_TEXT_MUTED, COLOR_CARD_BG, COLOR_CARD_BORDER
)


class OperatorsDialog(tk.Toplevel):
    def __init__(self, parent, on_change_callback=None):
        super().__init__(parent)
        self.title("Gerenciamento de Operadores & Assinaturas — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        setup_common_styles(ttk.Style(self))
        self.on_change_callback = on_change_callback
        self.editing_op_id = None

        self._build_ui()
        self._load_operators_list()
        center_modal(self, parent, 610, 560)

    def _build_ui(self):
        main_box = ttk.Frame(self, padding="12 10 12 10")
        main_box.pack(fill=tk.BOTH, expand=True)

        # ─── HEADER ───
        lbl_head = tk.Label(
            main_box,
            text="👤 Perfis de Operador, WhatsApp & Assinaturas",
            font=("Segoe UI", 11, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_head.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            main_box,
            text="Cadastre os operadores do sistema. Cada perfil possui seu próprio e-mail, número/chave de WhatsApp para alertas automáticos e assinatura corporativa.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, wraplength=570, justify=tk.LEFT
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
        self.tree.heading("telefone", text="WhatsApp / Telefone")

        self.tree.column("padrao", width=65, anchor=tk.CENTER)
        self.tree.column("nome", width=120, anchor=tk.W)
        self.tree.column("cargo", width=115, anchor=tk.W)
        self.tree.column("email", width=145, anchor=tk.W)
        self.tree.column("telefone", width=115, anchor=tk.W)

        sb = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("default_op", background="#EBF5E7")
        self.tree.tag_configure("active_op", background="#F2F8ED")

        self.tree.bind("<<TreeviewSelect>>", self._on_select_operator)

        # Botões de ação da tabela
        frame_tbtns = tk.Frame(main_box, bg=COLOR_BG_LIGHT)
        frame_tbtns.pack(fill=tk.X, pady=(0, 8))

        btn_new = create_btn_secondary(frame_tbtns, "➕ Novo Operador", self._start_new_operator, pady=4)
        btn_new.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_set_default = create_btn_secondary(frame_tbtns, "⭐ Definir como Padrão", self._set_as_default, pady=4)
        self.btn_set_default.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_delete = create_btn_danger(frame_tbtns, "🗑️ Excluir", self._delete_selected, pady=4)
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

        lbl_t = ttk.Label(self.frame_form, text="WhatsApp (com DDD):", font=("Segoe UI", 9, "bold"))
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

        lbl_k = ttk.Label(self.frame_form, text="Chave CallMeBot Wpp:", font=("Segoe UI", 8))
        lbl_k.grid(row=3, column=0, sticky=tk.W, pady=3)
        self.ent_wpp_apikey = ttk.Entry(self.frame_form, font=("Segoe UI", 9), width=21)
        self.ent_wpp_apikey.grid(row=3, column=1, sticky=tk.W, pady=3, padx=(4, 12))

        self.btn_test_wpp = create_btn_secondary(
            self.frame_form, "📲 Testar WhatsApp", self._test_operator_whatsapp, pady=2
        )
        self.btn_test_wpp.grid(row=3, column=2, columnspan=2, sticky=tk.W, pady=3)

        self.var_default = tk.BooleanVar(value=False)
        self.chk_def = ttk.Checkbutton(
            self.frame_form,
            text="Operador Padrão (usado em agendamentos automáticos e inicialização)",
            variable=self.var_default
        )
        self.chk_def.grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=(4, 4))

        frame_fbtns = tk.Frame(self.frame_form, bg=COLOR_BG_LIGHT)
        frame_fbtns.grid(row=5, column=0, columnspan=4, sticky=tk.E, pady=(2, 0))

        self.btn_save_op = create_btn_primary(
            frame_fbtns, "💾 Salvar Operador", self._save_operator_form, pady=4
        )
        self.btn_save_op.pack(side=tk.RIGHT)

        # ─── FOOTER ───
        frame_foot = tk.Frame(main_box, bg=COLOR_BG_LIGHT)
        frame_foot.pack(fill=tk.X, side=tk.BOTTOM, pady=(6, 0))

        btn_close = create_btn_secondary(frame_foot, "Fechar", self.destroy, pady=5)
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
            row_tag = "default_op" if is_def else ("active_op" if op_id == active_id else "normal_op")
            phone_disp = op.get("whatsapp_phone") or op.get("phone", "")
            iid = self.tree.insert(
                "", tk.END,
                values=(status_txt, op.get("name", ""), op.get("role", ""), op.get("email", ""), phone_disp),
                tags=(op_id, row_tag)
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
        self.ent_tel.insert(0, target.get("whatsapp_phone") or target.get("phone", ""))

        self.ent_wpp_apikey.delete(0, tk.END)
        self.ent_wpp_apikey.insert(0, target.get("whatsapp_apikey", ""))

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
        self.ent_wpp_apikey.delete(0, tk.END)
        self.ent_smtp_user.delete(0, tk.END)
        self.ent_smtp_pwd.delete(0, tk.END)
        self.var_default.set(False)

        self.frame_form.config(text="  Novo Operador  ")
        self.ent_nome.focus_set()

    def _test_operator_whatsapp(self):
        nome = self.ent_nome.get().strip() or "Operador"
        phone = self.ent_tel.get().strip()
        apikey = self.ent_wpp_apikey.get().strip()
        cfg = load_config()
        if not apikey:
            apikey = cfg.get("webhook_whatsapp_apikey", "").strip()

        if not phone:
            messagebox.showwarning("WhatsApp", "Por favor, digite o número do WhatsApp com DDD.", parent=self)
            self.ent_tel.focus_set()
            return
        if not apikey:
            messagebox.showwarning(
                "WhatsApp",
                "Nenhuma Chave API CallMeBot informada para este operador nem nas configurações globais.\n\n"
                "Para ativar gratuitamente no WhatsApp, envie 'I allow callmebot to call me' para o número oficial +34 623 75 84 18.",
                parent=self
            )
            self.ent_wpp_apikey.focus_set()
            return

        self.btn_test_wpp.config(state=tk.DISABLED)
        def _worker():
            from core.webhook_notifier import send_test_webhook
            ok, msg = send_test_webhook("whatsapp", "", "", "", phone, apikey, operator_name=nome)
            def _ui():
                self.btn_test_wpp.config(state=tk.NORMAL)
                if ok:
                    messagebox.showinfo("WhatsApp Enviado", f"Mensagem de teste enviada com sucesso para o WhatsApp de '{nome}' ({phone})!\n\n{msg}", parent=self)
                else:
                    messagebox.showerror("Falha no WhatsApp", f"Não foi possível enviar a mensagem para o WhatsApp de '{nome}':\n\n{msg}", parent=self)
            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _save_operator_form(self):
        nome = self.ent_nome.get().strip()
        email = self.ent_email.get().strip()
        role = self.ent_cargo.get().strip() or "Técnico de Sistemas"
        phone = self.ent_tel.get().strip()
        wpp_key = self.ent_wpp_apikey.get().strip()
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
            "whatsapp_phone": phone,
            "whatsapp_apikey": wpp_key,
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
