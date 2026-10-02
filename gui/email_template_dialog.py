import os
import subprocess
import threading
import tempfile
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

from core.email_sender import (
    load_email_template, save_email_template, get_email_template_path,
    DEFAULT_HTML_BODY, DEFAULT_SUBJECT_TEMPLATE, render_email, prepare_html_for_preview
)

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MUTED = "#55664C"

class EmailTemplateDialog(tk.Toplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("Personalizar Modelo de E-mail — CompaSSS")
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self.on_save_callback = on_save_callback
        self._build_ui()
        self._load_template()
        self._center_on_parent(parent, 760, 650)

    def _center_on_parent(self, parent, w, h):
        self.update_idletasks()
        try:
            p_x = parent.winfo_rootx()
            p_y = parent.winfo_rooty()
            p_w = parent.winfo_width()
            p_h = parent.winfo_height()
            x = p_x + (p_w - w) // 2
            y = p_y + (p_h - h) // 2
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        except Exception:
            self.geometry(f"{w}x{h}")

    def _build_ui(self):
        container = ttk.Frame(self, padding="18 16 18 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── CABEÇALHO ───
        lbl_title = tk.Label(
            container, text="Modelo de Envio de E-mail",
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_title.pack(anchor=tk.W, pady=(0, 2))

        lbl_desc = tk.Label(
            container,
            text="Personalize o assunto e a mensagem padrão dos relatórios enviados. "
                 "As variáveis entre chaves { } serão substituídas automaticamente pelos dados reais de cada mês.",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT, justify=tk.LEFT
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 10))

        # ─── CAMPO ASSUNTO ───
        frame_subj = ttk.Frame(container)
        frame_subj.pack(fill=tk.X, pady=(0, 8))

        lbl_s = ttk.Label(frame_subj, text="Assunto:", font=("Segoe UI", 9, "bold"))
        lbl_s.pack(side=tk.LEFT, padx=(0, 8))

        self.ent_subject = ttk.Entry(frame_subj, font=("Segoe UI", 9))
        self.ent_subject.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ─── BARRA DE TAGS DINÂMICAS ───
        frame_tags = ttk.LabelFrame(container, text=" Inserir Variáveis Dinâmicas ", padding="8 6 8 6")
        frame_tags.pack(fill=tk.X, pady=(0, 10))

        tags = [
            ("{MES}", "Mês (ex: Outubro)"),
            ("{ANO}", "Ano (ex: 2026)"),
            ("{PERIODO}", "Período (29/09 a 28/10)"),
            ("{TOTAL_M3}", "Consumo Total m³"),
            ("{TOTAL_VALOR}", "Valor Total R$"),
            ("{VALOR_M3}", "Tarifa do m³"),
            ("{QTD_SALAS}", "Total de Salas"),
            ("{DATA_EMISSAO}", "Data/Hora Atual"),
        ]

        row_f = ttk.Frame(frame_tags)
        row_f.pack(fill=tk.X)

        for tag, hint in tags:
            btn = ttk.Button(
                row_f, text=tag,
                command=lambda t=tag: self._insert_tag(t),
                width=13
            )
            btn.pack(side=tk.LEFT, padx=3, pady=2)

        # ─── ÁREA DE TEXTO / HTML ───
        lbl_body = ttk.Label(container, text="Conteúdo da Mensagem (HTML):", font=("Segoe UI", 9, "bold"))
        lbl_body.pack(anchor=tk.W, pady=(0, 4))

        frame_editor = ttk.Frame(container)
        frame_editor.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.scroll_y = ttk.Scrollbar(frame_editor, orient=tk.VERTICAL)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.scroll_x = ttk.Scrollbar(frame_editor, orient=tk.HORIZONTAL)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.txt_content = tk.Text(
            frame_editor,
            font=("Consolas", 9),
            wrap=tk.NONE,
            yscrollcommand=self.scroll_y.set,
            xscrollcommand=self.scroll_x.set,
            undo=True
        )
        self.txt_content.pack(fill=tk.BOTH, expand=True)
        self.scroll_y.config(command=self.txt_content.yview)
        self.scroll_x.config(command=self.txt_content.xview)

        # ─── BOTÕES DE AÇÃO INFERIORES ───
        frame_actions = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_actions.pack(fill=tk.X)

        btn_preview = ttk.Button(frame_actions, text="👁 Visualizar Prévia", command=self._preview_in_browser)
        btn_preview.pack(side=tk.LEFT, padx=(0, 6))

        btn_notepad = ttk.Button(frame_actions, text="📝 Abrir no Bloco de Notas", command=self._open_in_notepad)
        btn_notepad.pack(side=tk.LEFT, padx=(0, 6))

        btn_restore = ttk.Button(frame_actions, text="↺ Restaurar Padrão", command=self._restore_default)
        btn_restore.pack(side=tk.LEFT)

        btn_save = tk.Button(
            frame_actions, text="💾 Salvar Modelo", command=self._save_template,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=14, pady=5, cursor="hand2"
        )
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = ttk.Button(frame_actions, text="Fechar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT)

    def _insert_tag(self, tag):
        """Insere a tag dinâmica na posição atual do cursor."""
        try:
            # Se o foco estiver no campo assunto, insere lá
            if self.focus_get() == self.ent_subject:
                self.ent_subject.insert(tk.INSERT, tag)
            else:
                self.txt_content.insert(tk.INSERT, tag)
                self.txt_content.focus_set()
        except Exception:
            self.txt_content.insert(tk.INSERT, tag)

    def _load_template(self):
        subject, content = load_email_template()
        self.ent_subject.delete(0, tk.END)
        self.ent_subject.insert(0, subject)

        self.txt_content.delete("1.0", tk.END)
        self.txt_content.insert("1.0", content)

    def _save_template(self):
        subject = self.ent_subject.get().strip() or DEFAULT_SUBJECT_TEMPLATE
        content = self.txt_content.get("1.0", tk.END).strip()

        if save_email_template(subject, content):
            if self.on_save_callback:
                self.on_save_callback(subject, content)
            messagebox.showinfo(
                "Modelo Salvo",
                f"O modelo de e-mail foi salvo com sucesso em:\n{get_email_template_path()}",
                parent=self
            )
            self.destroy()
        else:
            messagebox.showerror("Erro ao Salvar", "Não foi possível salvar as alterações no arquivo.", parent=self)

    def _restore_default(self):
        resp = messagebox.askyesno(
            "Restaurar Padrão",
            "Deseja realmente restaurar o modelo visual padrão do CompaSSS?\nSuas edições não salvas serão substituídas.",
            parent=self
        )
        if resp:
            self.ent_subject.delete(0, tk.END)
            self.ent_subject.insert(0, DEFAULT_SUBJECT_TEMPLATE)
            self.txt_content.delete("1.0", tk.END)
            self.txt_content.insert("1.0", DEFAULT_HTML_BODY.strip())

    def _preview_in_browser(self):
        """Gera uma prévia HTML renderizada com dados fictícios de exemplo e abre no navegador."""
        subject = self.ent_subject.get().strip() or DEFAULT_SUBJECT_TEMPLATE
        content = self.txt_content.get("1.0", tk.END)

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

        r_subj, r_body = render_email(subject, content, mock_context)
        preview_body = prepare_html_for_preview(r_body)

        try:
            temp_path = os.path.join(tempfile.gettempdir(), "previa_modelo_email.html")
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(preview_body)
            webbrowser.open(f"file:///{temp_path}")
        except Exception as e:
            messagebox.showerror("Erro na Prévia", f"Falha ao gerar prévia: {e}", parent=self)

    def _open_in_notepad(self):
        """Salva temporariamente e abre o arquivo no Bloco de Notas para edição externa sem travar a interface."""
        path = get_email_template_path()
        try:
            def _watch():
                proc = subprocess.Popen(["notepad.exe", path])
                proc.wait()
                self.after(0, self._load_template)

            threading.Thread(target=_watch, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir no Bloco de Notas: {e}", parent=self)
