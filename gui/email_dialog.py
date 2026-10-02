import os
import threading
import tempfile
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

from core.config_manager import load_config, save_config
from core.email_sender import (
    load_email_template, extract_report_summary, render_email,
    parse_recipients, open_in_outlook, send_email_smtp
)
from gui.email_template_dialog import EmailTemplateDialog

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"

class SendEmailDialog(tk.Toplevel):
    def __init__(self, parent, xlsx_path, pdf_path=None):
        super().__init__(parent)
        self.title("Enviar Relatório por E-mail — CompaSSS")
        self.geometry("580x560")
        self.minsize(540, 500)
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
        lbl_to = ttk.Label(container, text="Destinatários (Para):", font=("Segoe UI", 9, "bold"))
        lbl_to.pack(anchor=tk.W, pady=(0, 2))

        self.ent_to = ttk.Entry(container, font=("Segoe UI", 9))
        self.ent_to.pack(fill=tk.X, pady=(0, 2))
        def_recipients = self.config.get("email_recipients", "")
        if def_recipients:
            self.ent_to.insert(0, def_recipients)

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

        # ─── MODO DE ENVIO ───
        frame_mode = ttk.LabelFrame(container, text=" Método de Envio ", padding="10 8 10 8")
        frame_mode.pack(fill=tk.X, pady=(0, 12))

        saved_mode = self.config.get("email_send_mode", "outlook")
        self.var_mode = tk.StringVar(value=saved_mode)

        rb_outlook = ttk.Radiobutton(
            frame_mode,
            text="Microsoft Outlook (Abre a mensagem com os anexos pronta para enviar)",
            variable=self.var_mode,
            value="outlook"
        )
        rb_outlook.pack(anchor=tk.W, pady=2)

        rb_smtp = ttk.Radiobutton(
            frame_mode,
            text="Envio Direto via Servidor (SMTP / UOL Pro configurado)",
            variable=self.var_mode,
            value="smtp"
        )
        rb_smtp.pack(anchor=tk.W, pady=2)

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
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(self.rendered_body)
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
        recipients = parse_recipients(raw_to)
        if not recipients:
            messagebox.showerror(
                "Destinatário Obrigatório",
                "Por favor, informe ao menos um endereço de e-mail válido para envio.",
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
        mode = self.var_mode.get()

        # Salvar destinatários e modo no config para as próximas vezes
        self.config["email_recipients"] = raw_to
        self.config["email_send_mode"] = mode
        save_config(self.config)

        # Iniciar envio em thread secundária
        self.btn_send.config(state=tk.DISABLED)
        self.prog_bar.start(10)
        self.lbl_status.config(text="Processando envio...")

        threading.Thread(
            target=self._send_worker,
            args=(recipients, subject, self.rendered_body, attachments, mode),
            daemon=True
        ).start()

    def _send_worker(self, recipients, subject, body, attachments, mode):
        if mode == "outlook":
            ok, msg = open_in_outlook(recipients, subject, body, attachments)
        else:
            ok, msg = send_email_smtp(self.config, recipients, subject, body, attachments)

        def _ui_done():
            self.prog_bar.stop()
            self.btn_send.config(state=tk.NORMAL)
            if ok:
                self.lbl_status.config(text="Sucesso!")
                messagebox.showinfo("Envio de E-mail", msg, parent=self)
                self.destroy()
            else:
                self.lbl_status.config(text="Erro no envio.")
                messagebox.showerror("Falha no Envio", msg, parent=self)

        self.after(0, _ui_done)
