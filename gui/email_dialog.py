import os
import threading
import tempfile
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from core.config_manager import (
    load_config, save_config, get_operators, get_active_operator,
    get_default_condominio_emails, DEFAULT_CONDOMINIO_EMAIL_TO, DEFAULT_CONDOMINIO_EMAIL_CC
)
from core.email_sender import (
    load_email_template, extract_report_summary, render_email,
    parse_recipients, send_email_smtp, prepare_html_for_preview
)
from gui.email_template_dialog import EmailTemplateDialog
from gui.ui_helpers import (
    apply_window_icon, center_modal, setup_common_styles,
    create_btn_primary, create_btn_secondary, ModernProgressBar,
    create_card_frame, create_modern_badge, create_tooltip, bind_button_hover,
    COLOR_PRIMARY, COLOR_PRIMARY_HOVER, COLOR_BG_LIGHT, COLOR_TEXT_MUTED, COLOR_TEXT_MAIN
)

class SendEmailDialog(tk.Toplevel):
    def __init__(self, parent, xlsx_path, pdf_path=None):
        super().__init__(parent)
        self.title("Enviar Relatório por E-mail — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        setup_common_styles(ttk.Style(self))

        self.xlsx_path = xlsx_path
        self.pdf_path = pdf_path if pdf_path and os.path.exists(pdf_path) else (
            os.path.splitext(xlsx_path)[0] + ".pdf" if os.path.exists(os.path.splitext(xlsx_path)[0] + ".pdf") else None
        )

        self.config = load_config()
        self.operators = get_operators(self.config)
        self.active_operator = get_active_operator(self.config)
        self.summary = extract_report_summary(self.xlsx_path)
        self.subj_template, self.body_template = load_email_template()

        # Renderizar assunto e corpo com os dados desta medição e operador
        self.rendered_subj, self.rendered_body = render_email(
            self.subj_template, self.body_template, self.summary, operator=self.active_operator
        )

        # Destinatários e modo de teste
        self.is_test_mode = False
        to_cfg = self.config.get("email_recipients", "")
        # Se o salvo for o e-mail de teste pessoal, restaura o oficial do condomínio como padrão oficial
        if to_cfg and "zangari" in to_cfg.lower():
            self.saved_official_to = to_cfg
        else:
            self.saved_official_to = DEFAULT_CONDOMINIO_EMAIL_TO

        cc_cfg = self.config.get("email_cc", "")
        self.saved_official_cc = cc_cfg if cc_cfg else DEFAULT_CONDOMINIO_EMAIL_CC
        self.test_email = self.config.get("test_email", "breno.hsc75@gmail.com")

        self._build_ui()
        center_modal(self, parent, 660, 660)

    def _build_ui(self):
        container = ttk.Frame(self, padding="18 16 18 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── 1. CABEÇALHO MODERNO ───
        frame_head = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_head.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            frame_head, text="✉️ Disparo de Relatório Mensal",
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        ).pack(anchor=tk.W)

        tk.Label(
            frame_head, text="Condomínio Praça Pamplona • Envio Seguro via UOL Pro (CompaSSS)",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        ).pack(anchor=tk.W)

        # ─── 2. CARD DE MÉTRICAS DO RELATÓRIO ───
        card_metrics = create_card_frame(container, padx=10, pady=8)
        card_metrics.pack(fill=tk.X, pady=(0, 10))

        f_name = os.path.basename(self.xlsx_path)
        p_txt = self.summary.get("periodo") or "Período automático"
        tot_m3 = self.summary.get("total_m3") or "0,00"
        tot_rs = self.summary.get("total_valor") or "0,00"

        f_kpi_row = tk.Frame(card_metrics, bg="#FFFFFF")
        f_kpi_row.pack(fill=tk.X)
        for i in range(4):
            f_kpi_row.columnconfigure(i, weight=1)

        f_k1 = tk.Frame(f_kpi_row, bg="#F6FAF4", highlightbackground="#D5E5C9", highlightthickness=1, padx=6, pady=4)
        f_k1.grid(row=0, column=0, sticky=tk.EW, padx=2)
        tk.Label(f_k1, text="📄 RELATÓRIO", font=("Segoe UI", 7, "bold"), fg=COLOR_PRIMARY, bg="#F6FAF4").pack(anchor="w")
        clean_name = f_name[:-5] if f_name.lower().endswith(".xlsx") else f_name
        tk.Label(f_k1, text=clean_name[:20] + "..." if len(clean_name) > 20 else clean_name, font=("Segoe UI", 8, "bold"), fg=COLOR_TEXT_MAIN, bg="#F6FAF4").pack(anchor="w")

        f_k2 = tk.Frame(f_kpi_row, bg="#F6FAF4", highlightbackground="#D5E5C9", highlightthickness=1, padx=6, pady=4)
        f_k2.grid(row=0, column=1, sticky=tk.EW, padx=2)
        tk.Label(f_k2, text="📅 PERÍODO", font=("Segoe UI", 7, "bold"), fg=COLOR_PRIMARY, bg="#F6FAF4").pack(anchor="w")
        tk.Label(f_k2, text=p_txt, font=("Segoe UI", 8, "bold"), fg=COLOR_TEXT_MAIN, bg="#F6FAF4").pack(anchor="w")

        f_k3 = tk.Frame(f_kpi_row, bg="#F6FAF4", highlightbackground="#D5E5C9", highlightthickness=1, padx=6, pady=4)
        f_k3.grid(row=0, column=2, sticky=tk.EW, padx=2)
        tk.Label(f_k3, text="💧 CONSUMO", font=("Segoe UI", 7, "bold"), fg=COLOR_PRIMARY, bg="#F6FAF4").pack(anchor="w")
        tk.Label(f_k3, text=f"{tot_m3} m³", font=("Segoe UI", 8, "bold"), fg=COLOR_TEXT_MAIN, bg="#F6FAF4").pack(anchor="w")

        f_k4 = tk.Frame(f_kpi_row, bg="#F6FAF4", highlightbackground="#D5E5C9", highlightthickness=1, padx=6, pady=4)
        f_k4.grid(row=0, column=3, sticky=tk.EW, padx=2)
        tk.Label(f_k4, text="💰 TOTAL RATEADO", font=("Segoe UI", 7, "bold"), fg=COLOR_PRIMARY, bg="#F6FAF4").pack(anchor="w")
        tk.Label(f_k4, text=f"R$ {tot_rs}", font=("Segoe UI", 8, "bold"), fg=COLOR_PRIMARY, bg="#F6FAF4").pack(anchor="w")

        # ─── 3. CARD DE DESTINATÁRIOS & CONFIGURAÇÃO DO ENVIO ───
        card_dest = create_card_frame(container, padx=14, pady=10)
        card_dest.pack(fill=tk.X, pady=(0, 10))

        # Título do Card e Seletor de Modo
        f_dest_hdr = tk.Frame(card_dest, bg="#FFFFFF")
        f_dest_hdr.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            f_dest_hdr, text="📬 Destinatários do Relatório",
            font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(side=tk.LEFT)

        # Botões do Seletor de Modo (Oficial vs Teste)
        f_mode_btns = tk.Frame(f_dest_hdr, bg="#FFFFFF")
        f_mode_btns.pack(side=tk.RIGHT)

        self.btn_mode_official = create_btn_secondary(
            f_mode_btns, "🏢 Modo Oficial", self._set_official_mode, pady=3, padx=10
        )
        self.btn_mode_official.pack(side=tk.LEFT, padx=(0, 6))
        create_tooltip(self.btn_mode_official, "Preenche a lista oficial do condomínio (Gerente e equipe Zangari/CompaSSS)")

        self.btn_mode_test = create_btn_secondary(
            f_mode_btns, "🧪 Modo Teste (Para Mim)", self._set_test_mode, pady=3, padx=10
        )
        self.btn_mode_test.pack(side=tk.LEFT)
        create_tooltip(self.btn_mode_test, "Envia exclusivamente para o seu e-mail pessoal para testes, sem afetar o condomínio")

        # Banner Indicador do Modo Ativo
        self.lbl_mode_badge = tk.Label(
            card_dest,
            text="🏢 MODO OFICIAL ATIVO — Relatório será enviado para a gerência e equipe do condomínio.",
            font=("Segoe UI", 8, "bold"), fg="#2A6320", bg="#E8F4E5",
            highlightbackground="#B8DCB2", highlightthickness=1, padx=10, pady=4
        )
        self.lbl_mode_badge.pack(fill=tk.X, pady=(0, 8))

        # Remetente (Operador)
        f_rem = tk.Frame(card_dest, bg="#FFFFFF")
        f_rem.pack(fill=tk.X, pady=(0, 6))

        tk.Label(f_rem, text="Remetente:", font=("Segoe UI", 8, "bold"), bg="#FFFFFF", fg=COLOR_TEXT_MAIN, width=12, anchor="w").pack(side=tk.LEFT)
        self.ops_map = {f"{op.get('name')} <{op.get('email')}>": op for op in self.operators}
        op_names = list(self.ops_map.keys())
        self.cmb_op = ttk.Combobox(f_rem, values=op_names, state="readonly", font=("Segoe UI", 9))
        self.cmb_op.pack(side=tk.LEFT, fill=tk.X, expand=True)

        curr_op_name = None
        for name, op in self.ops_map.items():
            if self.active_operator and op.get("id") == self.active_operator.get("id"):
                curr_op_name = name
                break
        if curr_op_name:
            self.cmb_op.set(curr_op_name)
        elif op_names:
            self.cmb_op.set(op_names[0])
        self.cmb_op.bind("<<ComboboxSelected>>", self._on_operator_change)

        # Destinatário Principal (Para)
        f_to = tk.Frame(card_dest, bg="#FFFFFF")
        f_to.pack(fill=tk.X, pady=(0, 6))
        tk.Label(f_to, text="Para (Gerente):", font=("Segoe UI", 8, "bold"), bg="#FFFFFF", fg=COLOR_TEXT_MAIN, width=12, anchor="w").pack(side=tk.LEFT)
        self.ent_to = ttk.Entry(f_to, font=("Segoe UI", 9))
        self.ent_to.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ent_to.insert(0, self.saved_official_to)

        # Em Cópia (Cc)
        f_cc = tk.Frame(card_dest, bg="#FFFFFF")
        f_cc.pack(fill=tk.X, pady=(0, 4))
        tk.Label(f_cc, text="Em Cópia (Cc):", font=("Segoe UI", 8, "bold"), bg="#FFFFFF", fg=COLOR_TEXT_MAIN, width=12, anchor="w").pack(side=tk.LEFT)
        self.ent_cc = ttk.Entry(f_cc, font=("Segoe UI", 9))
        self.ent_cc.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ent_cc.insert(0, self.saved_official_cc)

        lbl_to_hint = tk.Label(
            card_dest, text="Dica: Separe múltiplos e-mails por ponto-e-vírgula (;)",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_to_hint.pack(anchor=tk.W, padx=(95, 0), pady=(0, 6))

        # Assunto
        f_subj = tk.Frame(card_dest, bg="#FFFFFF")
        f_subj.pack(fill=tk.X, pady=(0, 2))
        tk.Label(f_subj, text="Assunto:", font=("Segoe UI", 8, "bold"), bg="#FFFFFF", fg=COLOR_TEXT_MAIN, width=12, anchor="w").pack(side=tk.LEFT)
        self.ent_subj = ttk.Entry(f_subj, font=("Segoe UI", 9))
        self.ent_subj.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ent_subj.insert(0, self.rendered_subj)

        # ─── 4. CARD DE ANEXOS INCLUSOS ───
        card_att = create_card_frame(container, padx=14, pady=8)
        card_att.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            card_att, text="📎 Arquivos Anexos Inclusos",
            font=("Segoe UI", 8, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        ).pack(anchor=tk.W, pady=(0, 4))

        self.var_att_xlsx = tk.BooleanVar(value=True)
        chk_xlsx = ttk.Checkbutton(
            card_att,
            text=f"📊 Planilha Excel (.xlsx)  •  {os.path.basename(self.xlsx_path)}",
            variable=self.var_att_xlsx
        )
        chk_xlsx.pack(anchor=tk.W, pady=2)

        self.var_att_pdf = tk.BooleanVar(value=bool(self.pdf_path))
        pdf_label = (
            f"📕 Relatório Oficial em PDF (.pdf)  •  {os.path.basename(self.pdf_path)}"
            if self.pdf_path else "📕 Relatório Oficial em PDF (não encontrado)"
        )
        chk_pdf = ttk.Checkbutton(
            card_att,
            text=pdf_label,
            variable=self.var_att_pdf,
            state=tk.NORMAL if self.pdf_path else tk.DISABLED
        )
        chk_pdf.pack(anchor=tk.W, pady=2)

        # ─── 5. PROGRESSO E STATUS ───
        self.lbl_status = tk.Label(
            container, text="", font=("Segoe UI", 9, "bold"),
            fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        self.lbl_status.pack(anchor=tk.W, pady=(0, 2))

        self.prog_bar = ModernProgressBar(container, height=10)
        self.prog_bar.pack(fill=tk.X, pady=(0, 10))

        # ─── 6. BOTÕES DE AÇÃO INFERIORES ───
        frame_bottom = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_bottom.pack(fill=tk.X)

        btn_preview = create_btn_secondary(frame_bottom, "👁 Prévia HTML", self._preview_email, pady=6)
        btn_preview.pack(side=tk.LEFT, padx=(0, 8))

        btn_edit_template = create_btn_secondary(frame_bottom, "✏ Editar Modelo", self._edit_template, pady=6)
        btn_edit_template.pack(side=tk.LEFT)

        self.btn_send = create_btn_primary(
            frame_bottom, "✉ Enviar Agora", self._start_send,
            padx=18, pady=6
        )
        self.btn_send.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = create_btn_secondary(frame_bottom, "Cancelar", self.destroy, padx=14, pady=6)
        btn_cancel.pack(side=tk.RIGHT)

    def _set_official_mode(self):
        """Restaura os destinatários oficiais do Condomínio Praça Pamplona."""
        self.is_test_mode = False
        to_email, cc_emails = get_default_condominio_emails()
        to_val = self.saved_official_to or to_email
        cc_val = self.saved_official_cc or cc_emails

        self.ent_to.delete(0, tk.END)
        self.ent_to.insert(0, to_val)
        self.ent_cc.delete(0, tk.END)
        self.ent_cc.insert(0, cc_val)

        self.lbl_mode_badge.config(
            text="🏢 MODO OFICIAL ATIVO — Relatório será enviado para a gerência e equipe do condomínio.",
            fg="#2A6320", bg="#E8F4E5", highlightbackground="#B8DCB2"
        )

    def _set_test_mode(self):
        """Ativa o modo de teste enviando apenas para o endereço de teste pessoal."""
        self.is_test_mode = True
        curr_to = self.ent_to.get().strip()
        curr_cc = self.ent_cc.get().strip()
        if curr_to and curr_to != self.test_email:
            self.saved_official_to = curr_to
            self.saved_official_cc = curr_cc

        self.ent_to.delete(0, tk.END)
        self.ent_to.insert(0, self.test_email)
        self.ent_cc.delete(0, tk.END)

        self.lbl_mode_badge.config(
            text="🧪 MODO TESTE ATIVO — O relatório será enviado exclusivamente para seu e-mail pessoal.",
            fg="#9C4410", bg="#FEF3EB", highlightbackground="#F8D7BE"
        )

    def _on_operator_change(self, event=None):
        val = self.cmb_op.get()
        if val in self.ops_map:
            self.active_operator = self.ops_map[val]
            self.rendered_subj, self.rendered_body = render_email(
                self.subj_template, self.body_template, self.summary, operator=self.active_operator
            )
            self.ent_subj.delete(0, tk.END)
            self.ent_subj.insert(0, self.rendered_subj)

    def _preview_email(self):
        """Abre no navegador uma prévia exata do e-mail com os dados preenchidos."""
        try:
            temp_path = os.path.join(tempfile.gettempdir(), "previa_envio_hidrometros.html")
            preview_content = prepare_html_for_preview(self.rendered_body)
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(preview_content)
            webbrowser.open(f"file:///{temp_path}")
        except Exception as e:
            messagebox.showerror("Erro na Prévia", f"Não foi possível gerar a prévia:\n{e}", parent=self)

    def _edit_template(self):
        """Abre o editor de modelo de e-mail e recarrega os dados ao salvar."""
        def _on_save(new_subj, new_body):
            self.subj_template = new_subj
            self.body_template = new_body
            self.rendered_subj, self.rendered_body = render_email(
                self.subj_template, self.body_template, self.summary
            )
            self.ent_subj.delete(0, tk.END)
            self.ent_subj.insert(0, self.rendered_subj)

        EmailTemplateDialog(self, on_save_callback=_on_save)

    def _start_send(self):
        raw_to = self.ent_to.get().strip()
        raw_cc = self.ent_cc.get().strip()
        recipients = parse_recipients(raw_to)
        cc_recipients = parse_recipients(raw_cc)

        if not recipients:
            messagebox.showerror(
                "Destinatário Obrigatório",
                "Por favor, informe ao menos um endereço de e-mail no campo 'Para (Gerente)'.",
                parent=self
            )
            return

        # Anexos selecionados
        attachments = []
        if self.var_att_xlsx.get() and os.path.exists(self.xlsx_path):
            attachments.append(self.xlsx_path)
        if self.var_att_pdf.get() and self.pdf_path and os.path.exists(self.pdf_path):
            attachments.append(self.pdf_path)

        if not attachments:
            if not messagebox.askyesno(
                "Sem Anexos",
                "Nenhum anexo foi selecionado. Deseja enviar o e-mail mesmo sem anexar os arquivos?",
                parent=self
            ):
                return

        subject = self.ent_subj.get().strip() or self.rendered_subj

        # Confirmação explícita
        to_list_str = "\n".join(f"  • {e}" for e in recipients)
        cc_list_str = "\n".join(f"  • {e}" for e in cc_recipients) if cc_recipients else "  (Nenhum)"
        att_list_str = "\n".join(f"  • {os.path.basename(a)}" for a in attachments) if attachments else "  (Nenhum anexo)"

        if self.is_test_mode:
            confirm_msg = (
                f"🧪 CONFIRMAÇÃO — MODO DE TESTE ATIVO\n\n"
                f"O e-mail será enviado APENAS para seu endereço pessoal de teste:\n"
                f"{to_list_str}\n\n"
                f"Nenhum destinatário do condomínio receberá esta mensagem.\n"
                f"As configurações oficiais de produção NÃO serão alteradas.\n\n"
                f"Arquivo(s) em Anexo:\n{att_list_str}\n\n"
                f"Deseja disparar o e-mail de teste agora?"
            )
        else:
            confirm_msg = (
                f"🏢 CONFIRMAÇÃO DE DISPARO OFICIAL\n\n"
                f"Destinatário(s) Principal(is):\n{to_list_str}\n\n"
                f"Em Cópia (Cc):\n{cc_list_str}\n\n"
                f"Arquivo(s) em Anexo:\n{att_list_str}\n\n"
                f"Deseja confirmar e disparar o e-mail agora?"
            )

        if not messagebox.askyesno("Confirmar Destinatários", confirm_msg, parent=self, default=messagebox.YES):
            return

        # Salvar preferências: se modo teste, guarda o test_email; se oficial, atualiza email_recipients
        if self.is_test_mode:
            self.config["test_email"] = raw_to
            save_config(self.config)
        else:
            self.config["email_recipients"] = raw_to
            self.config["email_cc"] = raw_cc
            self.config["email_send_mode"] = "smtp"
            save_config(self.config)

        # Iniciar envio em thread secundária
        self.btn_send.config(state=tk.DISABLED)
        self.prog_bar.start(10)
        status_txt = "Enviando e-mail de teste..." if self.is_test_mode else "Enviando e-mail via UOL Pro..."
        self.lbl_status.config(text=status_txt)

        threading.Thread(
            target=self._send_worker,
            args=(recipients, cc_recipients, subject, self.rendered_body, attachments),
            daemon=True
        ).start()

    def _send_worker(self, recipients, cc_recipients, subject, body, attachments):
        ok, msg = send_email_smtp(self.config, recipients, subject, body, attachments, cc_addrs=cc_recipients, operator=self.active_operator)

        def _ui_done():
            self.prog_bar.stop()
            self.btn_send.config(state=tk.NORMAL)
            if ok:
                self.lbl_status.config(text="Sucesso!")
                try:
                    f_name = os.path.basename(self.xlsx_path)
                    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")
                    send_log = self.config.get("report_send_log", {})
                    send_log[f_name] = now_str
                    self.config["report_send_log"] = send_log
                    save_config(self.config)
                except Exception:
                    pass

                messagebox.showinfo("Envio de E-mail", msg, parent=self)
                self.destroy()
            else:
                self.lbl_status.config(text="Erro no envio.")
                messagebox.showerror("Falha no Envio", msg, parent=self)

        self.after(0, _ui_done)
