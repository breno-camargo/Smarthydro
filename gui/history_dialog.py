import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core.config_manager import load_config
from core.history_manager import fetch_annual_history, export_annual_history_excel
from gui.ui_helpers import apply_window_icon, center_modal

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_PEAK = "#D9534F"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"
COLOR_CARD_BG = "#FFFFFF"
COLOR_CARD_BORDER = "#D5E5C9"


class AnnualHistoryDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Histórico Anual de Telemetria — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        self.history_data = None
        self.tooltip_window = None

        self._build_ui()
        center_modal(self, parent, 610, 580)
        self._load_data_async()

    def _build_ui(self):
        self.main_box = ttk.Frame(self, padding="12 10 12 10")
        self.main_box.pack(fill=tk.BOTH, expand=True)

        # ─── HEADER ───
        frame_head = tk.Frame(self.main_box, bg=COLOR_BG_LIGHT)
        frame_head.pack(fill=tk.X, pady=(0, 6))

        lbl_title = tk.Label(
            frame_head, text="📈 Histórico Anual de Consumo de Água (12 Meses)",
            font=("Segoe UI", 11, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        )
        lbl_title.pack(anchor=tk.W)

        self.lbl_sub = tk.Label(
            frame_head, text="Condomínio Praça Pamplona • Telemetria Schneider Electric EcoStruxure EBO",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        self.lbl_sub.pack(anchor=tk.W, pady=(1, 4))

        # ─── CARDS DE KPIS ───
        self.frame_kpis = tk.Frame(self.main_box, bg=COLOR_BG_LIGHT)
        self.frame_kpis.pack(fill=tk.X, pady=(0, 8))

        self.card_total = self._create_kpi_card(self.frame_kpis, "💧 Consumo Anual", "Carregando...", "Soma 12 ciclos")
        self.card_total.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        self.card_fat = self._create_kpi_card(self.frame_kpis, "💰 Faturamento Anual", "Carregando...", "Tarifa base")
        self.card_fat.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

        self.card_media = self._create_kpi_card(self.frame_kpis, "📊 Média Mensal", "Carregando...", "Média/mês")
        self.card_media.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

        self.card_pico = self._create_kpi_card(self.frame_kpis, "⚡ Mês de Pico", "Carregando...", "Maior medição")
        self.card_pico.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        # ─── GRÁFICO VISUAL (CANVAS) ───
        frame_chart_box = tk.LabelFrame(
            self.main_box, text="  Evolução Visual do Consumo Mensal (m³)  ",
            bg=COLOR_BG_LIGHT, fg=COLOR_PRIMARY, font=("Segoe UI", 9, "bold"),
            padx=6, pady=4
        )
        frame_chart_box.pack(fill=tk.X, pady=(0, 8))

        self.canvas_w = 560
        self.canvas_h = 135
        self.canvas = tk.Canvas(
            frame_chart_box, width=self.canvas_w, height=self.canvas_h,
            bg="#FFFFFF", highlightthickness=1, highlightbackground=COLOR_CARD_BORDER
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # ─── TABELA DE MEDIÇÕES ───
        frame_table_box = tk.LabelFrame(
            self.main_box, text="  Detalhamento dos Ciclos de Faturamento  ",
            bg=COLOR_BG_LIGHT, fg=COLOR_PRIMARY, font=("Segoe UI", 9, "bold"),
            padx=6, pady=4
        )
        frame_table_box.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columns = ("mes", "periodo", "consumo", "diff_m3", "diff_pct", "faturamento")
        self.tree = ttk.Treeview(frame_table_box, columns=columns, show="headings", height=4)
        self.tree.heading("mes", text="Mês / Ano")
        self.tree.heading("periodo", text="Período de Medição")
        self.tree.heading("consumo", text="Consumo (m³)")
        self.tree.heading("diff_m3", text="Variação (m³)")
        self.tree.heading("diff_pct", text="Variação (%)")
        self.tree.heading("faturamento", text="Valor Faturado (R$)")

        self.tree.column("mes", width=65, anchor=tk.CENTER)
        self.tree.column("periodo", width=155, anchor=tk.CENTER)
        self.tree.column("consumo", width=80, anchor=tk.E)
        self.tree.column("diff_m3", width=75, anchor=tk.E)
        self.tree.column("diff_pct", width=65, anchor=tk.E)
        self.tree.column("faturamento", width=95, anchor=tk.E)

        # Tags visuais de alta/baixa
        self.tree.tag_configure("up", foreground="#C9302C")
        self.tree.tag_configure("down", foreground="#2D7A1E")
        self.tree.tag_configure("neutral", foreground=COLOR_TEXT_MAIN)

        sb = ttk.Scrollbar(frame_table_box, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ─── PROGRESS BAR / STATUS ───
        self.lbl_status = tk.Label(
            self.main_box, text="Conectando ao SQL Server e extraindo dados dos 12 meses...",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        self.lbl_status.pack(anchor=tk.W, pady=(0, 4))

        self.prog_bar = ttk.Progressbar(self.main_box, mode="determinate", maximum=12)
        self.prog_bar.pack(fill=tk.X, pady=(0, 8))

        # ─── FOOTER & AÇÕES ───
        frame_foot = tk.Frame(self.main_box, bg=COLOR_BG_LIGHT)
        frame_foot.pack(fill=tk.X, side=tk.BOTTOM)

        btn_close = ttk.Button(frame_foot, text="Fechar", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

        self.btn_export = tk.Button(
            frame_foot, text="📥 Exportar Excel (.xlsx)", command=self._export_excel,
            bg=COLOR_PRIMARY, fg="white", activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white", font=("Segoe UI", 9, "bold"), relief="flat",
            padx=10, pady=3, cursor="hand2", state=tk.DISABLED
        )
        self.btn_export.pack(side=tk.RIGHT, padx=(0, 6))

        btn_refresh = ttk.Button(frame_foot, text="🔄 Atualizar", command=self._load_data_async)
        btn_refresh.pack(side=tk.LEFT)

    def _create_kpi_card(self, parent, title, val, sub):
        card = tk.Frame(
            parent, bg=COLOR_CARD_BG,
            highlightbackground=COLOR_CARD_BORDER, highlightthickness=1,
            padx=6, pady=5
        )
        lbl_t = tk.Label(card, text=title, font=("Segoe UI", 8, "bold"), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG)
        lbl_t.pack(anchor=tk.W)

        lbl_v = tk.Label(card, text=val, font=("Segoe UI", 10, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD_BG)
        lbl_v.pack(anchor=tk.W, pady=(1, 1))

        lbl_s = tk.Label(card, text=sub, font=("Segoe UI", 7), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG)
        lbl_s.pack(anchor=tk.W)

        card.lbl_val = lbl_v
        card.lbl_sub = lbl_s
        return card

    def _load_data_async(self):
        self.btn_export.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.lbl_status.config(text="Consultando histórico dos últimos 12 meses no SQL Server...")

        def _worker():
            try:
                def _cb(current, total, msg):
                    self.after(0, lambda: self._update_progress(current, total, msg))

                data = fetch_annual_history(progress_callback=_cb)
                self.after(0, lambda: self._on_data_loaded(data))
            except Exception as e:
                self.after(0, lambda: self._on_data_error(str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _update_progress(self, current, total, msg):
        self.prog_bar["maximum"] = total
        self.prog_bar["value"] = current
        self.lbl_status.config(text=msg)

    def _on_data_loaded(self, data):
        self.history_data = data
        kpis = data.get("kpis", {})
        months = data.get("months", [])

        # 1. Atualizar cards
        tot_m3_fmt = f"{kpis.get('total_consumo_m3', 0):,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
        tot_rs_fmt = f"R$ {kpis.get('total_faturamento_rs', 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        med_m3_fmt = f"{kpis.get('media_mensal_m3', 0):,.1f} m³/mês".replace(",", "X").replace(".", ",").replace("X", ".")
        pico_fmt = f"{kpis.get('pico_mes', '-')} ({kpis.get('pico_m3', 0):,.1f} m³)".replace(",", "X").replace(".", ",").replace("X", ".")

        self.card_total.lbl_val.config(text=tot_m3_fmt)
        self.card_fat.lbl_val.config(text=tot_rs_fmt)
        self.card_media.lbl_val.config(text=med_m3_fmt)
        self.card_pico.lbl_val.config(text=pico_fmt)

        self.card_fat.lbl_sub.config(text=f"Tarifa: R$ {kpis.get('tarifa_m3', 63.68):.2f}/m³")
        self.card_pico.lbl_sub.config(text=kpis.get('pico_periodo', ''))
        self.lbl_sub.config(text=f"Condomínio Praça Pamplona • Período: {kpis.get('periodo_geral', '')} • 12 Ciclos Consolidados")

        # 2. Desenhar Gráfico no Canvas
        self._render_chart(months, kpis)

        # 3. Preencher Tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, m in enumerate(reversed(months)):
            diff_m3 = m["diff_m3"]
            diff_pct = m["diff_pct"]
            if idx == len(months) - 1:
                # Primeiro mês não tem variação anterior
                str_diff_m3 = "-"
                str_diff_pct = "-"
                tag = "neutral"
            else:
                str_diff_m3 = f"{diff_m3:+.1f} m³".replace(".", ",")
                str_diff_pct = f"{diff_pct:+.1f}%".replace(".", ",")
                tag = "up" if diff_m3 > 0 else ("down" if diff_m3 < 0 else "neutral")

            consumo_fmt = f"{m['consumo_m3']:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
            valor_fmt = f"R$ {m['valor_rs']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            self.tree.insert(
                "", tk.END,
                values=(m["mes_label"], m["periodo"], consumo_fmt, str_diff_m3, str_diff_pct, valor_fmt),
                tags=(tag,)
            )

        self.lbl_status.config(text="✓ Dados consolidados com sucesso direto do Schneider EBO SQL Server.")
        self.btn_export.config(state=tk.NORMAL)

    def _on_data_error(self, err_msg):
        self.lbl_status.config(text=f"Erro ao carregar histórico: {err_msg}")
        messagebox.showerror("Erro", f"Não foi possível carregar os dados históricos:\n\n{err_msg}", parent=self)

    def _render_chart(self, months, kpis):
        self.canvas.delete("all")
        if not months:
            return

        w = self.canvas.winfo_width() or self.canvas_w
        h = self.canvas.winfo_height() or self.canvas_h

        pad_left = 38
        pad_right = 16
        pad_top = 20
        pad_bottom = 22

        chart_w = w - pad_left - pad_right
        chart_h = h - pad_top - pad_bottom

        max_val = max(m["consumo_m3"] for m in months) or 100.0
        # Arredondar topo para múltiplo de 200
        upper_limit = (int(max_val // 200) + 1) * 200

        # Linhas de grade horizontais sutis
        for i in range(4):
            y_val = upper_limit * (i / 3)
            y_pos = pad_top + chart_h - (chart_h * (y_val / upper_limit))
            self.canvas.create_line(pad_left, y_pos, w - pad_right, y_pos, fill="#EBF3E6", width=1)
            self.canvas.create_text(pad_left - 6, y_pos, text=f"{int(y_val)}", font=("Segoe UI", 7), fill=COLOR_TEXT_MUTED, anchor=tk.E)

        # Linha pontilhada da média anual
        media_val = kpis.get("media_mensal_m3", 0)
        if media_val > 0:
            y_media = pad_top + chart_h - (chart_h * (media_val / upper_limit))
            self.canvas.create_line(pad_left, y_media, w - pad_right, y_media, fill=COLOR_ACCENT, width=1.5, dash=(4, 3))
            self.canvas.create_text(w - pad_right, y_media - 8, text=f"Média: {media_val:.0f} m³", font=("Segoe UI", 7, "bold"), fill=COLOR_PRIMARY, anchor=tk.E)

        # Desenhar Barras
        n = len(months)
        slot_w = chart_w / n
        bar_w = slot_w * 0.58

        pico_val = kpis.get("pico_m3", 0)

        for i, m in enumerate(months):
            val = m["consumo_m3"]
            b_height = (val / upper_limit) * chart_h
            x_center = pad_left + (i + 0.5) * slot_w
            x1 = x_center - bar_w / 2
            x2 = x_center + bar_w / 2
            y2 = pad_top + chart_h
            y1 = y2 - b_height

            is_pico = (val == pico_val and val > 0)
            bar_color = COLOR_PEAK if is_pico else COLOR_PRIMARY
            bar_hover_color = "#C9302C" if is_pico else COLOR_PRIMARY_HOVER

            bar_id = self.canvas.create_rectangle(x1, y1, x2, y2, fill=bar_color, outline=bar_color, width=1)

            # Rótulo do valor no topo da barra
            val_txt = f"{int(round(val))}"
            self.canvas.create_text(x_center, y1 - 8, text=val_txt, font=("Segoe UI", 7, "bold" if is_pico else "normal"), fill=bar_color, anchor=tk.CENTER)

            # Rótulo do mês na base
            self.canvas.create_text(x_center, y2 + 10, text=m["mes_label"].split("/")[0], font=("Segoe UI", 7, "bold"), fill=COLOR_TEXT_MAIN, anchor=tk.CENTER)

    def _export_excel(self):
        if not self.history_data:
            return

        cfg = load_config()
        out_dir = cfg.get("output_directory", os.path.expanduser("~"))
        default_file = os.path.join(out_dir, "Historico_Anual_Consumo_Praca_Pamplona.xlsx")

        target = filedialog.asksaveasfilename(
            parent=self,
            title="Salvar Histórico Anual em Excel",
            initialdir=out_dir,
            initialfile="Historico_Anual_Consumo_Praca_Pamplona.xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")]
        )
        if not target:
            return

        try:
            saved_path = export_annual_history_excel(self.history_data, target)
            if messagebox.askyesno(
                "Exportação Concluída",
                f"Relatório anual gerado com sucesso em:\n\n{saved_path}\n\nDeseja abrir o arquivo no Excel agora?",
                parent=self
            ):
                try:
                    os.startfile(saved_path)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Erro ao Exportar", f"Falha ao gerar planilha Excel:\n\n{e}", parent=self)
