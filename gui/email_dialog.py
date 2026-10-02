import os
import threading
import tempfile
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from core.config_manager import load_config, save_config
from core.email_sender import (
    load_email_template, extract_report_summary, render_email,
    parse_recipients, send_email_smtp, prepare_html_for_preview
)
from gui.email_template_dialog import EmailTemplateDialog

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MUTED = "#55664C"

class SendEmailDialog(tk.Toplevel):
    def __init__(self, parent, xlsx_path, pdf_path=None):
        super().__init__(parent)
        self.title("Enviar Relatório por E-mail — CompaSSS")
        self.geometry("580x520")
        self.minsize(540, 470)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self.xlsx_path = xlsx_path
        self.pdf_path = pdf_path if pdf_path and os.path.exists(pdf_path) else (
            os.path.splitext(xlsx_path)[0] + ".pdf" if os.path.exists(os.path.splitext(xlsx_path)[0] + ".pdf") else None
        )

        self.config = load_config()
        self.summary = extract_report_summary(self.xlsx_path)
        self.subj_template, self.body_template = load_email_template()

        # Renderizar assunto e corpo com os dados desta medição
        self.rendered_subj, self.rendered_body = render_email(
            self.subj_template, self.body_template, self.summary
        )

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self, padding="18 16 18 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── TÍTULO & DETALHES DO RELATÓRIO ───
        lbl_head = tk.Label(
            container, text="Envio de Relatório Mensal",
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_head.pack(anchor=tk.W, pady=(0, 2))

        f_name = os.path.basename(self.xlsx_path)
        p_txt = self.summary.get("periodo") or "Período automático"
        tot_m3 = self.summary.get("total_m3") or "0,00"
        tot_rs = self.summary.get("total_valor") or "0,00"

        info_txt = f"Relatório: {f_name}\nPeríodo: {p_txt}  •  Consumo: {tot_m3} m³  •  Total: R$ {tot_rs}"
        lbl_info = tk.Label(
            container, text=info_txt,
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, justify=tk.LEFT
        )
        lbl_info.pack(anchor=tk.W, pady=(0, 12))

        # ─── DESTINATÁRIOS ───
        lbl_to = ttk.Label(container, text="Para (Gerente):", font=("Segoe UI", 9, "bold"))
        lbl_to.pack(anchor=tk.W, pady=(0, 2))

        self.ent_to = ttk.Entry(container, font=("Segoe UI", 9))
        self.ent_to.pack(fill=tk.X, pady=(0, 4))
        def_recipients = self.config.get("email_recipients", "")
        if def_recipients:
            self.ent_to.insert(0, def_recipients)

        lbl_cc = ttk.Label(container, text="Em Cópia (Cc):", font=("Segoe UI", 9, "bold"))
        lbl_cc.pack(anchor=tk.W, pady=(4, 2))

        self.ent_cc = ttk.Entry(container, font=("Segoe UI", 9))
        self.ent_cc.pack(fill=tk.X, pady=(0, 2))
        def_cc = self.config.get("email_cc", "")
        if def_cc:
            self.ent_cc.insert(0, def_cc)

        lbl_to_hint = tk.Label(
            container, text="Separe múltiplos e-mails por ponto-e-vírgula (;) ou vírgula (,)",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_to_hint.pack(anchor=tk.W, pady=(0, 10))

        # ─── ASSUNTO ───
        lbl_subj = ttk.Label(container, text="Assunto:", font=("Segoe UI", 9, "bold"))
        lbl_subj.pack(anchor=tk.W, pady=(0, 2))

        self.ent_subj = ttk.Entry(container, font=("Segoe UI", 9))
        self.ent_subj.pack(fill=tk.X, pady=(0, 12))
        self.ent_subj.insert(0, self.rendered_subj)

        # ─── ANEXOS INCLUSOS ───
        frame_att = ttk.LabelFrame(container, text=" Arquivos Anexados ", padding="10 8 10 8")
        frame_att.pack(fill=tk.X, pady=(0, 12))

        self.var_att_xlsx = tk.BooleanVar(value=True)
        chk_xlsx = ttk.Checkbutton(
            frame_att,
            text=f"Planilha Excel (.xlsx)  —  {os.path.basename(self.xlsx_path)}",
            variable=self.var_att_xlsx
        )
        chk_xlsx.pack(anchor=tk.W, pady=2)

        self.var_att_pdf = tk.BooleanVar(value=bool(self.pdf_path))
        pdf_label = (
            f"Relatório Oficial em PDF (.pdf)  —  {os.path.basename(self.pdf_path)}"
            if self.pdf_path else "Relatório Oficial em PDF (não encontrado)"
        )
        chk_pdf = ttk.Checkbutton(
            frame_att,
            text=pdf_label,
            variable=self.var_att_pdf,
            state=tk.NORMAL if self.pdf_path else tk.DISABLED
        )
        chk_pdf.pack(anchor=tk.W, pady=2)



        # ─── BARRA DE STATUS / PROGRESSO ───
        self.lbl_status = tk.Label(
            container, text="", font=("Segoe UI", 9, "bold"),
            fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        self.lbl_status.pack(anchor=tk.W, pady=(0, 4))

        self.prog_bar = ttk.Progressbar(container, mode="indeterminate")
        self.prog_bar.pack(fill=tk.X, pady=(0, 10))

        # ─── BOTÕES DE AÇÃO ───
        frame_bottom = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_bottom.pack(fill=tk.X, pady=(4, 0))

        btn_preview = ttk.Button(frame_bottom, text="👁 Prévia", command=self._preview_email)
        btn_preview.pack(side=tk.LEFT, padx=(0, 6))

        btn_edit_template = ttk.Button(frame_bottom, text="✏ Editar Modelo", command=self._edit_template)
        btn_edit_template.pack(side=tk.LEFT)

        self.btn_send = tk.Button(
            frame_bottom, text="✉ Enviar Agora", command=self._start_send,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=16, pady=5, cursor="hand2"
        )
        self.btn_send.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = ttk.Button(frame_bottom, text="Cancelar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT)

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
            messagebox.showwarning(
                "Sem Anexos",
                "Nenhum anexo foi selecionado. Deseja enviar o e-mail mesmo sem anexar os arquivos?",
                parent=self
            )

        subject = self.ent_subj.get().strip() or self.rendered_subj

        # Confirmação explícita dos destinatários antes do disparo real para total segurança
        to_list_str = "\n".join(f"  • {e}" for e in recipients)
        cc_list_str = "\n".join(f"  • {e}" for e in cc_recipients) if cc_recipients else "  (Nenhum)"
        att_list_str = "\n".join(f"  • {os.path.basename(a)}" for a in attachments) if attachments else "  (Nenhum anexo)"

        confirm_msg = (
            f"Por favor, confira os destinatários antes do disparo:\n\n"
            f"Destinatário(s) Principal(is):\n{to_list_str}\n\n"
            f"Em Cópia (Cc):\n{cc_list_str}\n\n"
            f"Arquivo(s) em Anexo:\n{att_list_str}\n\n"
            f"Deseja confirmar e disparar o e-mail agora?"
        )
        if not messagebox.askyesno("Confirmar Destinatários", confirm_msg, parent=self, default=messagebox.YES):
            return

        # Salvar destinatários no config para as próximas vezes
        self.config["email_recipients"] = raw_to
        self.config["email_cc"] = raw_cc
        self.config["email_send_mode"] = "smtp"
        save_config(self.config)

        # Iniciar envio em thread secundária
        self.btn_send.config(state=tk.DISABLED)
        self.prog_bar.start(10)
        self.lbl_status.config(text="Enviando e-mail via UOL Pro...")

        threading.Thread(
            target=self._send_worker,
            args=(recipients, cc_recipients, subject, self.rendered_body, attachments),
            daemon=True
        ).start()

    def _send_worker(self, recipients, cc_recipients, subject, body, attachments):
        ok, msg = send_email_smtp(self.config, recipients, subject, body, attachments, cc_addrs=cc_recipients)

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
