import tkinter as tk
from tkinter import ttk, messagebox

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MUTED = "#55664C"


class AnomalyDialog(tk.Toplevel):
    def __init__(self, parent, anomalies, periodo_str=""):
        super().__init__(parent)
        self.title("🔍 Auditoria Interna de Medição — CompaSSS")
        self.geometry("720x460")
        self.minsize(640, 380)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self.anomalies = anomalies
        self.periodo_str = periodo_str

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self, padding="18 16 18 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ─── CABEÇALHO ───
        frame_header = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_header.pack(fill=tk.X, pady=(0, 10))

        vazamentos = [a for a in self.anomalies if "Vazamento" in a["tipo"]]
        parados = [a for a in self.anomalies if "Travado" in a["tipo"]]

        titulo_texto = f"Auditoria de Consumo — {len(self.anomalies)} Alerta(s) Encontrado(s)"
        lbl_title = tk.Label(
            frame_header, text=titulo_texto,
            font=("Segoe UI", 12, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_title.pack(anchor=tk.W)

        sub_desc = []
        if vazamentos:
            sub_desc.append(f"{len(vazamentos)} suspeita(s) de vazamento")
        if parados:
            sub_desc.append(f"{len(parados)} possível(is) medidor(es) travado(s)")

        lbl_sub = tk.Label(
            frame_header,
            text=f"Período: {self.periodo_str} • " + " • ".join(sub_desc),
            font=("Segoe UI", 9), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_sub.pack(anchor=tk.W, pady=(2, 0))

        # ─── TABELA DE ANOMALIAS ───
        frame_table = ttk.Frame(container)
        frame_table.pack(fill=tk.BOTH, expand=True, pady=(6, 10))

        columns = ("sala", "consumo", "media", "variacao", "diagnostico")
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("sala", text="Unidade / Sala", anchor=tk.W)
        self.tree.heading("consumo", text="Consumo Atual", anchor=tk.E)
        self.tree.heading("media", text="Média (11m)", anchor=tk.E)
        self.tree.heading("variacao", text="Variação", anchor=tk.CENTER)
        self.tree.heading("diagnostico", text="Diagnóstico Sugerido", anchor=tk.W)

        self.tree.column("sala", width=170, minwidth=140, anchor=tk.W)
        self.tree.column("consumo", width=105, minwidth=90, anchor=tk.E)
        self.tree.column("media", width=100, minwidth=85, anchor=tk.E)
        self.tree.column("variacao", width=95, minwidth=80, anchor=tk.CENTER)
        self.tree.column("diagnostico", width=210, minwidth=170, anchor=tk.W)

        scroll_y = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags visuais para linhas
        self.tree.tag_configure("vazamento_alta", background="#FFEBEE", foreground="#C62828")
        self.tree.tag_configure("vazamento_media", background="#FFF3E0", foreground="#E65100")
        self.tree.tag_configure("travado", background="#F5F5F5", foreground="#616161")

        for item in self.anomalies:
            c_m3 = f"{item['consumo']:.2f} m³".replace(".", ",")
            m_m3 = f"{item['media']:.2f} m³".replace(".", ",") if item['media'] > 0 else "0,00 m³"

            tag = "travado"
            if "Vazamento" in item["tipo"]:
                tag = "vazamento_alta" if item.get("severidade") == "alta" else "vazamento_media"

            self.tree.insert("", tk.END, values=(
                item["sala"],
                c_m3,
                m_m3,
                item["variacao"],
                item["tipo"]
            ), tags=(tag,))

        # ─── NOTA INFORMATIVA INFERIOR ───
        frame_note = tk.Frame(container, bg="#EBF3E6", bd=1, relief="solid")
        frame_note.pack(fill=tk.X, pady=(0, 10), ipady=6, ipadx=8)

        lbl_note = tk.Label(
            frame_note,
            text="ℹ️ Uso exclusivo da Automação Predial: Este diagnóstico é apenas para sua conferência prévia.\nEle NÃO é adicionado à planilha do Excel nem ao relatório oficial em PDF.",
            font=("Segoe UI", 8), fg="#2D4F1E", bg="#EBF3E6", justify=tk.LEFT
        )
        lbl_note.pack(anchor=tk.W)

        # ─── BOTÕES DE AÇÃO ───
        frame_bottom = tk.Frame(container, bg=COLOR_BG_LIGHT)
        frame_bottom.pack(fill=tk.X)

        btn_copy = ttk.Button(
            frame_bottom, text="📋 Copiar Lista de Suspeitas",
            command=self._copy_to_clipboard
        )
        btn_copy.pack(side=tk.LEFT)

        btn_close = tk.Button(
            frame_bottom, text="✓ Entendido / Fechar",
            command=self.destroy,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=16, pady=5, cursor="hand2"
        )
        btn_close.pack(side=tk.RIGHT)

    def _copy_to_clipboard(self):
        """Copia os dados da auditoria em formato texto legível para colar no WhatsApp ou bloco de notas."""
        linhas = [
            f"🔍 *Auditoria de Medição — Praça Pamplona ({self.periodo_str})*",
            "--------------------------------------------------",
        ]
        for a in self.anomalies:
            linhas.append(f"• *{a['sala']}*: Atual={a['consumo']:.2f} m³ | Média={a['media']:.2f} m³ ({a['variacao']}) ➔ {a['tipo']}")
        linhas.append("--------------------------------------------------")
        linhas.append("Fonte: Telemetria StruxureWare EBO / CompaSSS")

        texto_final = "\n".join(linhas)
        self.clipboard_clear()
        self.clipboard_append(texto_final)
        messagebox.showinfo(
            "Copiado",
            "A lista de salas suspeitas foi copiada para a área de transferência!\nVocê pode colar no WhatsApp da manutenção ou onde desejar.",
            parent=self
        )
