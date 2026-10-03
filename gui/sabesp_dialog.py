"""
Diálogo de Importação e Análise de Fatura da Sabesp para SmartHydro.
Permite selecionar um PDF local ou buscar da caixa de entrada via IMAP,
exibe os dados extraídos e aplica automaticamente datas e tarifas na tela principal.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

from core.sabesp_parser import parse_sabesp_pdf, search_sabesp_in_email
from gui.ui_helpers import (
    apply_window_icon, center_modal, setup_common_styles,
    create_btn_primary, create_btn_secondary,
    COLOR_PRIMARY, COLOR_PRIMARY_HOVER, COLOR_ACCENT, COLOR_BG_LIGHT,
    COLOR_TEXT_MAIN, COLOR_TEXT_MUTED, COLOR_CARD_BG, COLOR_CARD_BORDER
)


class SabespImportDialog(tk.Toplevel):
    def __init__(self, parent, config: dict, on_apply_callback=None, on_generate_callback=None, target_month=None, target_year=None):
        super().__init__(parent)
        self.title("Fatura Sabesp — Preenchimento Inteligente")
        self.geometry("630x530")
        self.minsize(580, 480)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        setup_common_styles(ttk.Style(self))

        self.config = config
        self.on_apply_callback = on_apply_callback
        self.on_generate_callback = on_generate_callback
        self.target_month = target_month
        self.target_year = target_year
        self.sabesp_data = None

        self._center_window()
        self._build_ui()
        self._auto_check_existing_or_email()

    def _center_window(self):
        self.update_idletasks()
        w = 630
        h = 530
        x = max(0, self.master.winfo_x() + (self.master.winfo_width() - w) // 2)
        y = max(0, self.master.winfo_y() + (self.master.winfo_height() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _auto_check_existing_or_email(self):
        """Verifica se já existe uma fatura recente baixada no cache para pré-carregar."""
        from core.config_manager import get_base_dir
        cache_dir = os.path.join(get_base_dir(), "temp_sabesp")
        if os.path.exists(cache_dir):
            pdfs = [os.path.join(cache_dir, f) for f in os.listdir(cache_dir) if f.lower().endswith(".pdf")]
            if pdfs:
                latest = max(pdfs, key=os.path.getmtime)
                self._process_pdf(latest)

    def _build_ui(self):
        # Top banner
        header = tk.Frame(self, bg=COLOR_PRIMARY, height=55)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        lbl_icon = tk.Label(header, text="📄", font=("Segoe UI", 22), bg=COLOR_PRIMARY, fg="white")
        lbl_icon.pack(side=tk.LEFT, padx=(15, 8))

        title_frame = tk.Frame(header, bg=COLOR_PRIMARY)
        title_frame.pack(side=tk.LEFT, fill=tk.Y, pady=6)
        lbl_title = tk.Label(title_frame, text="Leitor Inteligente de Fatura Sabesp", font=("Segoe UI", 12, "bold"), bg=COLOR_PRIMARY, fg="white")
        lbl_title.pack(anchor=tk.W)
        lbl_sub = tk.Label(title_frame, text="Conecta na caixa de e-mail e extrai datas e tarifas da concessionária", font=("Segoe UI", 8), bg=COLOR_PRIMARY, fg="#D8E8D0")
        lbl_sub.pack(anchor=tk.W)

        # Content Frame
        container = tk.Frame(self, bg=COLOR_BG_LIGHT, padx=15, pady=12)
        container.pack(fill=tk.BOTH, expand=True)

        # Action Buttons frame (Select File or Search in Email)
        f_actions = tk.LabelFrame(container, text="Origem da Conta de Água", font=("Segoe UI", 9, "bold"), bg=COLOR_BG_LIGHT, fg=COLOR_TEXT_MAIN, padx=10, pady=8)
        f_actions.pack(fill=tk.X, pady=(0, 10))

        btn_email = create_btn_primary(
            f_actions,
            "📩 Buscar no E-mail (IMAP)",
            self._on_search_email,
            pady=6
        )
        btn_email.pack(side=tk.LEFT, padx=(0, 10))

        btn_file = create_btn_secondary(
            f_actions,
            "📁 Escolher PDF Local...",
            self._on_select_file,
            pady=6
        )
        btn_file.pack(side=tk.LEFT)

        self.lbl_loading = ttk.Label(f_actions, text="", font=("Segoe UI", 8, "italic"), foreground=COLOR_TEXT_MUTED)
        self.lbl_loading.pack(side=tk.LEFT, padx=15)

        # Result Card Frame
        self.f_card = tk.LabelFrame(container, text="Dados Reconhecidos da Sabesp", font=("Segoe UI", 9, "bold"), bg=COLOR_BG_LIGHT, fg=COLOR_TEXT_MAIN, padx=12, pady=10)
        self.f_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Initial prompt inside card
        self.lbl_empty = ttk.Label(
            self.f_card,
            text="Nenhuma fatura carregada ainda.\nClique em 'Buscar no E-mail' acima ou selecione um PDF.",
            font=("Segoe UI", 9), justify=tk.CENTER, foreground=COLOR_TEXT_MUTED
        )
        self.lbl_empty.pack(expand=True, pady=30)

        # Data grid (hidden until loaded)
        self.f_grid = tk.Frame(self.f_card, bg=COLOR_BG_LIGHT)

        # Divisor suave antes da barra inferior
        div_foot = tk.Frame(self, height=1, bg=COLOR_CARD_BORDER)
        div_foot.pack(fill=tk.X, side=tk.BOTTOM)

        # Bottom Bar
        bottom_bar = tk.Frame(self, bg=COLOR_BG_LIGHT)
        bottom_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=14, pady=10)

        self.btn_cancel = create_btn_secondary(bottom_bar, "Fechar", self.destroy, pady=6)
        self.btn_cancel.pack(side=tk.RIGHT)

        self.btn_apply_generate = create_btn_primary(
            bottom_bar,
            "🚀 Aplicar e Gerar Relatório",
            self._on_apply_and_generate,
            pady=6, state=tk.DISABLED
        )
        self.btn_apply_generate.pack(side=tk.RIGHT, padx=(0, 8))

        self.btn_apply = create_btn_secondary(
            bottom_bar,
            "✔ Apenas Aplicar Parâmetros",
            self._on_apply,
            pady=6, state=tk.DISABLED
        )
        self.btn_apply.pack(side=tk.RIGHT, padx=(0, 8))

        self.var_tarifa_escolhida = tk.StringVar(value="faixa")

    def _on_select_file(self):
        f = filedialog.askopenfilename(
            parent=self,
            title="Selecione o PDF da Conta de Água da Sabesp",
            filetypes=[("Fatura Sabesp (PDF)", "*.pdf"), ("Todos os Arquivos", "*.*")]
        )
        if not f:
            return
        self._process_pdf(f)

    def _on_search_email(self):
        self.lbl_loading.config(text="Buscando faturas na caixa de entrada...")
        self.update_idletasks()

        def _worker():
            ok, msg, pdf_path = search_sabesp_in_email(
                self.config,
                target_month=self.target_month,
                target_year=self.target_year
            )
            self.after(0, lambda: self._on_email_search_done(ok, msg, pdf_path))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_email_search_done(self, ok, msg, pdf_path):
        self.lbl_loading.config(text="")
        if not ok or not pdf_path:
            messagebox.showinfo("Busca de E-mail", msg, parent=self)
            return
        self._process_pdf(pdf_path)

    def _process_pdf(self, pdf_path: str):
        ok, msg, data = parse_sabesp_pdf(pdf_path)
        if not ok:
            messagebox.showerror("Erro ao Processar Fatura", msg, parent=self)
            return

        self.sabesp_data = data
        self._render_data_card(data)
        self.btn_apply.config(state=tk.NORMAL)
        self.btn_apply_generate.config(state=tk.NORMAL)

    def _render_data_card(self, d: dict):
        self.lbl_empty.pack_forget()
        self.f_grid.pack(fill=tk.BOTH, expand=True)

        for widget in self.f_grid.winfo_children():
            widget.destroy()

        def add_row(parent, row, label, value, is_bold=False, highlight=False):
            fg_color = COLOR_PRIMARY if highlight else COLOR_TEXT_MAIN
            font_spec = ("Segoe UI", 9, "bold") if (is_bold or highlight) else ("Segoe UI", 9)

            lbl = ttk.Label(parent, text=label, font=("Segoe UI", 9), foreground=COLOR_TEXT_MUTED)
            lbl.grid(row=row, column=0, sticky=tk.W, pady=3, padx=(0, 10))

            val = tk.Label(parent, text=value, font=font_spec, fg=fg_color, bg=COLOR_BG_LIGHT)
            val.grid(row=row, column=1, sticky=tk.W, pady=3)

        # 1. Unidade / Cliente
        add_row(self.f_grid, 0, "Condomínio:", f"{d.get('cliente')} (Hidrômetro: {d.get('hidrometro') or 'Geral'})", is_bold=True)

        # 2. Período Sabesp vs Período do Rateio
        p_sabesp = f"{d.get('leitura_anterior')} a {d.get('leitura_atual')} ({d.get('dias_faturamento')} dias)"
        add_row(self.f_grid, 1, "Leitura Sabesp:", p_sabesp)

        p_rateio = f"{d.get('periodo_rateio_ini')} a {d.get('periodo_rateio_fim')} (Ciclo das Salas)"
        add_row(self.f_grid, 2, "Período do Relatório:", p_rateio, is_bold=True, highlight=True)

        # 3. Consumo e Valor Total Sabesp
        m3_sabesp = f"{d.get('consumo_sabesp_m3'):,.1f} m³".replace(".", ",")
        val_sabesp = f"R$ {d.get('valor_total_fatura'):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        add_row(self.f_grid, 3, "Consumo Geral Sabesp:", f"{m3_sabesp}  |  Total da Fatura: {val_sabesp}", is_bold=True)

        # 4. Divisor de tarifas
        sep = ttk.Separator(self.f_grid, orient=tk.HORIZONTAL)
        sep.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=8)

        # 5. Opções de Tarifa para o Rateio
        lbl_opt = ttk.Label(self.f_grid, text="Tarifa do m³ para aplicar no rateio:", font=("Segoe UI", 9, "bold"))
        lbl_opt.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tar_faixa_txt = f"Tarifa da Faixa Sabesp (> 50 m³): R$ {d.get('tarifa_faixa'):.2f} (Água R$ {d.get('tarifa_agua'):.2f} + Esgoto R$ {d.get('tarifa_esgoto'):.2f}) [Padrão Pamplona]"
        r_faixa = tk.Radiobutton(
            self.f_grid,
            text=tar_faixa_txt,
            variable=self.var_tarifa_escolhida,
            value="faixa",
            font=("Segoe UI", 9, "bold"),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG_LIGHT,
            activebackground=COLOR_BG_LIGHT
        )
        r_faixa.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=2)

        tar_media_txt = f"Custo Médio Efetivo da Fatura: R$ {d.get('tarifa_media'):.2f} (Total Fatura ÷ Total m³ Sabesp)"
        r_media = tk.Radiobutton(
            self.f_grid,
            text=tar_media_txt,
            variable=self.var_tarifa_escolhida,
            value="media",
            font=("Segoe UI", 9),
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_LIGHT,
            activebackground=COLOR_BG_LIGHT
        )
        r_media.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=2)

        # 6. Informações de Vencimento e Origem
        venc_str = f"Vencimento: {d.get('vencimento') or 'N/I'}  |  Próxima Leitura: {d.get('proxima_leitura') or 'N/I'}"
        lbl_venc = ttk.Label(self.f_grid, text=venc_str, font=("Segoe UI", 8, "italic"), foreground=COLOR_TEXT_MUTED)
        lbl_venc.grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))

    def _on_apply(self):
        if not self.sabesp_data:
            return

        choice = self.var_tarifa_escolhida.get()
        if choice == "media":
            selected_rate = self.sabesp_data.get("tarifa_media")
        else:
            selected_rate = self.sabesp_data.get("tarifa_faixa")

        if self.on_apply_callback:
            self.on_apply_callback(self.sabesp_data, selected_rate)

        self.destroy()

    def _on_apply_and_generate(self):
        self._on_apply()
        if self.on_generate_callback:
            self.on_generate_callback()

