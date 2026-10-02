import os
import sys
import tkinter as tk
from tkinter import ttk
import webbrowser

from core.config_manager import get_base_dir

COLOR_PRIMARY = "#3D6B24"
COLOR_PRIMARY_HOVER = "#2D501A"
COLOR_ACCENT = "#90C671"
COLOR_BG_LIGHT = "#F6F9F2"
COLOR_TEXT_MAIN = "#1B2A12"
COLOR_TEXT_MUTED = "#55664C"
COLOR_CARD_BORDER = "#D5E5C9"


class AboutDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sobre o Desenvolvedor — CompaSSS")
        self.geometry("500x520")
        self.minsize(480, 500)
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        self._center_window(parent)
        self._build_ui()

    def _center_window(self, parent):
        self.update_idletasks()
        try:
            p_x = parent.winfo_rootx()
            p_y = parent.winfo_rooty()
            p_w = parent.winfo_width()
            p_h = parent.winfo_height()
            w = 500
            h = 520
            x = p_x + (p_w - w) // 2
            y = p_y + (p_h - h) // 2
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _build_ui(self):
        container = tk.Frame(self, bg=COLOR_BG_LIGHT, padx=22, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        # ─── CABEÇALHO COM CARD ───
        card = tk.Frame(
            container, bg="#FFFFFF",
            highlightbackground=COLOR_CARD_BORDER,
            highlightthickness=1,
            padx=20, pady=18
        )
        card.pack(fill=tk.BOTH, expand=True)

        # Ícone de Desenvolvedor
        lbl_avatar = tk.Label(
            card, text="👨‍💻", font=("Segoe UI Emoji", 34),
            bg="#FFFFFF"
        )
        lbl_avatar.pack(pady=(0, 4))

        # Nome do Software
        lbl_app_name = tk.Label(
            card, text="SmartHydro",
            font=("Segoe UI", 16, "bold"), fg=COLOR_PRIMARY, bg="#FFFFFF"
        )
        lbl_app_name.pack()

        lbl_app_sub = tk.Label(
            card, text="Automação de Telemetria e Rateio de Hidrômetros",
            font=("Segoe UI", 9, "bold"), fg=COLOR_TEXT_MAIN, bg="#FFFFFF"
        )
        lbl_app_sub.pack(pady=(2, 6))

        lbl_version = tk.Label(
            card, text="Versão 2.1 • CompaSSS Automação Predial",
            font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_version.pack(pady=(0, 12))

        # Divisória
        div = tk.Frame(card, height=1, bg=COLOR_CARD_BORDER)
        div.pack(fill=tk.X, pady=(0, 14))

        # ─── INFORMAÇÕES DO DESENVOLVEDOR ───
        frame_info = tk.Frame(card, bg="#FFFFFF")
        frame_info.pack(fill=tk.X)
        frame_info.columnconfigure(1, weight=1)

        info_items = [
            ("Desenvolvedor:", "Breno Camargo", True),
            ("E-mail:", "breno.camargo@compasss.com.br", False),
            ("Empresa:", "CompaSSS Tecnologia e Automação", False),
            ("Empreendimento:", "Condomínio Praça Pamplona", False),
            ("Integração BMS:", "StruxureWare EBO (SQL Server)", False),
            ("Linguagem:", "Python 3.11 & Tkinter Nativo", False),
        ]

        for r_idx, (label, val, is_bold) in enumerate(info_items):
            lbl_title = tk.Label(
                frame_info, text=label,
                font=("Segoe UI", 9, "bold"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF",
                anchor="w"
            )
            lbl_title.grid(row=r_idx, column=0, sticky=tk.W, pady=3, padx=(0, 10))

            lbl_val = tk.Label(
                frame_info, text=val,
                font=("Segoe UI", 9, "bold" if is_bold else "normal"),
                fg=COLOR_PRIMARY if is_bold else COLOR_TEXT_MAIN,
                bg="#FFFFFF", anchor="w"
            )
            lbl_val.grid(row=r_idx, column=1, sticky=tk.W, pady=3)

        # ─── BOTÃO DO GITHUB ───
        btn_github = tk.Button(
            card,
            text="🌐 Ver Repositório no GitHub (breno-camargo/Smarthydro)",
            command=lambda: webbrowser.open("https://github.com/breno-camargo/Smarthydro"),
            bg="#EBF3E6", fg=COLOR_PRIMARY, activebackground=COLOR_ACCENT,
            font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=6,
            cursor="hand2"
        )
        btn_github.pack(pady=(16, 6))

        # ─── RODAPÉ / DIREITOS ───
        lbl_rights = tk.Label(
            card,
            text="© 2026 Breno Camargo — Todos os direitos reservados.",
            font=("Segoe UI", 7, "italic"), fg=COLOR_TEXT_MUTED, bg="#FFFFFF"
        )
        lbl_rights.pack(pady=(6, 0))

        # ─── BOTÃO FECHAR ───
        btn_close = ttk.Button(container, text="Fechar", command=self.destroy, width=12)
        btn_close.pack(pady=(12, 0))
