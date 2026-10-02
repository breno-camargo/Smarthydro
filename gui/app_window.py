import os
import sys
import threading
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from PIL import Image, ImageTk

from core.config_manager import load_config, save_config, get_base_dir, get_recent_reports, format_report_filename, get_report_output_folder
from core.database import test_db_connection
from cli.runner import execute_extraction
from gui.settings_dialog import SettingsDialog
from gui.email_dialog import SendEmailDialog
from gui.anomaly_dialog import AnomalyDialog

# Cores institucionais CompaSSS
COLOR_PRIMARY = "#3D6B24"       # Verde escuro institucional
COLOR_PRIMARY_HOVER = "#2D501A" # Verde escuro ao passar o mouse
COLOR_ACCENT = "#90C671"        # Verde da logo CompaSSS
COLOR_BG_LIGHT = "#F6F9F2"      # Fundo suave esverdeado
COLOR_TEXT_MAIN = "#1B2A12"     # Texto principal escuro
COLOR_TEXT_MUTED = "#55664C"    # Texto secundário

class AppHidrometrosWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("CompaSSS — Medição de Água Praça Pamplona")
        self.root.geometry("640x700")
        self.root.minsize(600, 660)
        self.root.configure(bg=COLOR_BG_LIGHT)

        self._set_window_icon()

        # Configurações do tema e estilos
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._configure_styles()
        self.config = load_config()

        self._build_ui()
        self._set_default_dates()
        self._refresh_history()

    def _set_window_icon(self):
        try:
            if sys.platform == "win32":
                try:
                    import ctypes
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("smarthydro.hidrometros.relatorio.1.0")
                except Exception:
                    pass

            # Candidatos para o arquivo .ico nativo do Windows
            ico_candidates = [
                os.path.join(get_base_dir(), "app_icon.ico"),
                os.path.join(get_base_dir(), "icon.ico"),
            ]
            if hasattr(sys, '_MEIPASS'):
                ico_candidates.insert(0, os.path.join(sys._MEIPASS, "app_icon.ico"))
                ico_candidates.insert(1, os.path.join(sys._MEIPASS, "icon.ico"))

            for p in ico_candidates:
                if os.path.exists(p):
                    try:
                        self.root.iconbitmap(default=p)
                        return
                    except Exception:
                        try:
                            self.root.iconbitmap(p)
                            return
                        except Exception:
                            pass

            # Fallback com imagem PNG apenas se iconbitmap falhar
            img_candidates = [
                os.path.join(get_base_dir(), "logo_final.png"),
                os.path.join(get_base_dir(), "gui_logo.png"),
            ]
            if hasattr(sys, '_MEIPASS'):
                img_candidates.insert(0, os.path.join(sys._MEIPASS, "logo_final.png"))

            for p in img_candidates:
                if os.path.exists(p):
                    try:
                        img = Image.open(p).resize((32, 32), Image.Resampling.LANCZOS)
                        self._icon_photo = ImageTk.PhotoImage(img)
                        self.root.iconphoto(False, self._icon_photo)
                        return
                    except Exception:
                        pass
        except Exception:
            pass

    def _configure_styles(self):
        self.style.configure(".", background=COLOR_BG_LIGHT, font=("Segoe UI", 9))
        self.style.configure("TLabel", background=COLOR_BG_LIGHT, foreground=COLOR_TEXT_MAIN)
        self.style.configure("TLabelframe", background=COLOR_BG_LIGHT, bordercolor=COLOR_ACCENT)
        self.style.configure("TLabelframe.Label", background=COLOR_BG_LIGHT, foreground=COLOR_PRIMARY, font=("Segoe UI", 10, "bold"))
        self.style.configure("TCheckbutton", background=COLOR_BG_LIGHT, foreground=COLOR_TEXT_MAIN)
        self.style.configure("TProgressbar", troughcolor="#E3EDD8", background=COLOR_ACCENT)

        # Estilo dos botões secundários e do histórico
        self.style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=6)
        self.style.configure("History.TButton", font=("Segoe UI", 9), padding=(6, 2))
        self.style.configure("Delete.TButton", font=("Segoe UI", 9), padding=(6, 2))

    def _build_ui(self):
        main_container = ttk.Frame(self.root, padding="20 12 20 12")
        main_container.pack(fill=tk.BOTH, expand=True)

        # ─── TOP BAR (LOGO + TÍTULO + CONFIGURAÇÕES) ───
        frame_top = tk.Frame(main_container, bg=COLOR_BG_LIGHT)
        frame_top.pack(fill=tk.X, pady=(0, 10))

        # Tentar carregar logo CompaSSS
        self.logo_img = self._load_logo_image()
        if self.logo_img:
            lbl_logo = tk.Label(frame_top, image=self.logo_img, bg=COLOR_BG_LIGHT)
            lbl_logo.pack(side=tk.LEFT, padx=(0, 14))

        title_box = tk.Frame(frame_top, bg=COLOR_BG_LIGHT)
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        lbl_title = tk.Label(title_box, text="Medição de Hidrômetros", font=("Segoe UI", 15, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT)
        lbl_title.pack(anchor=tk.W)
        lbl_sub = tk.Label(title_box, text="Condomínio Praça Pamplona  •  StruxureWare EBO", font=("Segoe UI", 9), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT)
        lbl_sub.pack(anchor=tk.W)

        # Botões de Configurações e Desenvolvedor no canto superior direito
        btn_settings = ttk.Button(frame_top, text="⚙ Configurações", command=self._open_settings, style="Secondary.TButton")
        btn_settings.pack(side=tk.RIGHT, anchor=tk.NE, pady=4)

        btn_about = ttk.Button(frame_top, text="👨‍💻 Desenvolvedor", command=self._open_about, style="Secondary.TButton")
        btn_about.pack(side=tk.RIGHT, anchor=tk.NE, pady=4, padx=(0, 6))

        # Linha divisória verde suave
        div = tk.Frame(main_container, height=2, bg=COLOR_ACCENT)
        div.pack(fill=tk.X, pady=(0, 16))

        # ─── PARÂMETROS DE EXTRAÇÃO (CARD PRINCIPAL) ───
        frame_card = ttk.LabelFrame(main_container, text="  Parâmetros do Relatório  ", padding="16 14 16 14")
        frame_card.pack(fill=tk.X, pady=(0, 14))
        frame_card.columnconfigure(1, weight=1)

        # Data Inicial com Mini Calendário DateEntry
        lbl_ini = ttk.Label(frame_card, text="Data Inicial:", font=("Segoe UI", 9, "bold"))
        lbl_ini.grid(row=0, column=0, sticky=tk.W, pady=8)

        self.cal_inicio = DateEntry(
            frame_card, width=14, font=("Segoe UI", 9),
            background=COLOR_PRIMARY, foreground="white",
            headersbackground=COLOR_PRIMARY, headersforeground="white",
            selectbackground=COLOR_ACCENT, selectforeground="black",
            date_pattern="dd/mm/yyyy", locale="pt_BR", borderwidth=1
        )
        self.cal_inicio.grid(row=0, column=1, sticky=tk.W, pady=8, padx=(10, 0))

        # Data Final com Mini Calendário DateEntry
        lbl_fim = ttk.Label(frame_card, text="Data Final:", font=("Segoe UI", 9, "bold"))
        lbl_fim.grid(row=1, column=0, sticky=tk.W, pady=8)

        self.cal_fim = DateEntry(
            frame_card, width=14, font=("Segoe UI", 9),
            background=COLOR_PRIMARY, foreground="white",
            headersbackground=COLOR_PRIMARY, headersforeground="white",
            selectbackground=COLOR_ACCENT, selectforeground="black",
            date_pattern="dd/mm/yyyy", locale="pt_BR", borderwidth=1
        )
        self.cal_fim.grid(row=1, column=1, sticky=tk.W, pady=8, padx=(10, 0))

        # Botão rápido para preencher o ciclo fechado (dia 29 ao dia 28)
        btn_ciclo = ttk.Button(
            frame_card, text="⚡ Preencher Ciclo Atual", command=self._apply_closed_cycle
        )
        btn_ciclo.grid(row=0, column=2, rowspan=2, padx=(18, 0), pady=8, sticky=tk.W)

        # Valor do m³ (R$)
        lbl_val = ttk.Label(frame_card, text="Valor do m³ (R$):", font=("Segoe UI", 9, "bold"))
        lbl_val.grid(row=2, column=0, sticky=tk.W, pady=8)

        self.ent_valor = ttk.Entry(frame_card, font=("Segoe UI", 9), width=16)
        self.ent_valor.insert(0, str(self.config.get("default_m3_price", "63.68")))
        self.ent_valor.grid(row=2, column=1, sticky=tk.W, pady=8, padx=(10, 0))

        # Pasta de Saída
        lbl_dir = ttk.Label(frame_card, text="Salvar em:", font=("Segoe UI", 9, "bold"))
        lbl_dir.grid(row=3, column=0, sticky=tk.W, pady=8)

        frame_out = tk.Frame(frame_card, bg=COLOR_BG_LIGHT)
        frame_out.grid(row=3, column=1, sticky=tk.EW, pady=8, padx=(10, 0))

        self.lbl_pasta = ttk.Entry(frame_out, font=("Segoe UI", 8))
        self.lbl_pasta.insert(0, self.config.get("output_directory", ""))
        self.lbl_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_browse = ttk.Button(frame_out, text="Alterar...", width=9, command=self._browse_output_dir)
        btn_browse.pack(side=tk.LEFT, padx=(6, 0))

        # Opção: Abrir planilha automaticamente
        self.var_open_excel = tk.BooleanVar(value=self.config.get("open_excel_after_generation", True))
        chk_open = ttk.Checkbutton(
            main_container,
            text="Abrir planilha no Excel automaticamente após a geração",
            variable=self.var_open_excel,
            command=self._update_open_excel_pref
        )
        chk_open.pack(anchor=tk.W, pady=(0, 4))

        # Opção: Ordenar por maior consumo
        self.var_sort_desc = tk.BooleanVar(value=self.config.get("sort_by_consumption", True))
        chk_sort = ttk.Checkbutton(
            main_container,
            text="Ordenar da sala de maior consumo para a de menor consumo",
            variable=self.var_sort_desc,
            command=self._update_sort_pref
        )
        chk_sort.pack(anchor=tk.W, pady=(0, 6))

        # ─── BOTÕES DE AÇÃO INFERIORES (DOCK NO BOTTOM PRIMEIRO PARA NÃO SER CORTADO) ───
        frame_actions = tk.Frame(main_container, bg=COLOR_BG_LIGHT)
        frame_actions.pack(fill=tk.X, side=tk.BOTTOM, pady=(8, 0))

        btn_open_folder = ttk.Button(
            frame_actions, text="Abrir Pasta", command=self._open_output_folder,
            style="Secondary.TButton"
        )
        btn_open_folder.pack(side=tk.LEFT, padx=(0, 8))

        btn_test = ttk.Button(
            frame_actions, text="Testar Conexão", command=self._test_connection_action,
            style="Secondary.TButton"
        )
        btn_test.pack(side=tk.LEFT)

        # Botão Principal Verde CompaSSS
        self.btn_gerar = tk.Button(
            frame_actions,
            text="✔ Gerar Relatório Excel",
            command=self._start_processing,
            bg=COLOR_PRIMARY,
            fg="white",
            activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            cursor="hand2"
        )
        self.btn_gerar.pack(side=tk.RIGHT)

        # ─── BARRA DE PROGRESSO E STATUS (DOCK NO BOTTOM) ───
        self.lbl_status = tk.Label(
            main_container,
            text="Pronto para gerar relatório. Selecione o período acima.",
            font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        self.lbl_status.pack(fill=tk.X, side=tk.BOTTOM, pady=(4, 8))

        self.prog_bar = ttk.Progressbar(main_container, mode="determinate", maximum=100)
        self.prog_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 4))
        self.prog_bar["value"] = 0

        # Rodapé com créditos de desenvolvimento (clicável)
        lbl_credits = tk.Label(
            main_container,
            text="Desenvolvido por Breno Camargo  •  CompaSSS Automação",
            font=("Segoe UI", 7),
            fg="#7E9672",
            bg=COLOR_BG_LIGHT,
            cursor="hand2"
        )
        lbl_credits.pack(side=tk.BOTTOM, pady=(0, 2))
        lbl_credits.bind("<Button-1>", lambda e: self._open_about())

        # ─── HISTÓRICO DE RELATÓRIOS RECENTES (COMPACTO) ───
        self.frame_history_card = ttk.LabelFrame(main_container, text="  Últimos Relatórios Gerados  ", padding="12 8 12 8")
        self.frame_history_card.pack(fill=tk.X, side=tk.TOP, pady=(6, 12))

        self.frame_history_list = tk.Frame(self.frame_history_card, bg=COLOR_BG_LIGHT)
        self.frame_history_list.pack(fill=tk.X, expand=True)

    def _load_logo_image(self):
        """Carrega a logo escrita institucional (igual à da planilha Excel) para o cabeçalho superior esquerdo."""
        candidates = [
            os.path.join(get_base_dir(), "gui_logo.png"),
            os.path.join(get_base_dir(), "excel_logo_0.png"),
            os.path.join(get_base_dir(), "logo_final.png"),
        ]
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
            candidates.insert(0, os.path.join(base, "gui_logo.png"))
            candidates.insert(1, os.path.join(base, "excel_logo_0.png"))

        for path in candidates:
            if os.path.exists(path):
                try:
                    img = Image.open(path)
                    # Redimensionar suavemente para altura de 40px mantendo proporção original do logo escrito
                    target_h = 40
                    target_w = int(img.width * (target_h / img.height))
                    img_resized = img.resize((target_w, target_h), Image.Resampling.LANCZOS)
                    return ImageTk.PhotoImage(img_resized)
                except Exception:
                    pass
        return None

    def _set_default_dates(self):
        """Define datas padrão para o ciclo anterior (ex: dia 29 ao dia 28)."""
        today = date.today()
        # Se hoje for antes do dia 28, o ciclo anterior terminou no dia 28 do mês passado
        if today.day < 28:
            # mês passado
            m_fim = today.month - 1 if today.month > 1 else 12
            y_fim = today.year if today.month > 1 else today.year - 1
        else:
            m_fim = today.month
            y_fim = today.year

        d_fim = date(y_fim, m_fim, 28)
        m_ini = m_fim - 1 if m_fim > 1 else 12
        y_ini = y_fim if m_fim > 1 else y_fim - 1
        d_ini = date(y_ini, m_ini, 29)

        try:
            self.cal_inicio.set_date(d_ini)
            self.cal_fim.set_date(d_fim)
        except Exception:
            pass

    def _apply_closed_cycle(self):
        """Preenche o último ciclo fechado (dia 29 a 28) com confirmação no status."""
        self._set_default_dates()
        try:
            ini_str = self.cal_inicio.get_date().strftime("%d/%m/%Y")
            fim_str = self.cal_fim.get_date().strftime("%d/%m/%Y")
            self.lbl_status.config(text=f"Datas ajustadas para o ciclo fechado: {ini_str} a {fim_str}")
        except Exception:
            pass

    def _browse_output_dir(self):
        curr = self.lbl_pasta.get().strip() or os.path.expanduser("~")
        selected = filedialog.askdirectory(initialdir=curr)
        if selected:
            self.lbl_pasta.delete(0, tk.END)
            self.lbl_pasta.insert(0, os.path.abspath(selected))
            self.config["output_directory"] = os.path.abspath(selected)
            save_config(self.config)

    def _update_open_excel_pref(self):
        self.config["open_excel_after_generation"] = self.var_open_excel.get()
        save_config(self.config)

    def _update_sort_pref(self):
        self.config["sort_by_consumption"] = self.var_sort_desc.get()
        save_config(self.config)

    def _open_settings(self):
        SettingsDialog(self.root, on_save_callback=self._on_settings_saved)

    def _open_about(self):
        """Abre diretamente a aba com informações de autoria e créditos do desenvolvedor."""
        SettingsDialog(self.root, on_save_callback=self._on_settings_saved, initial_tab=2)

    def _on_settings_saved(self, new_cfg):
        self.config = new_cfg
        self.lbl_pasta.delete(0, tk.END)
        self.lbl_pasta.insert(0, self.config.get("output_directory", ""))
        self.ent_valor.delete(0, tk.END)
        self.ent_valor.insert(0, str(self.config.get("default_m3_price", 63.68)))

    def _refresh_history(self, force=False):
        """Atualiza a lista visual dos relatórios gerados recentemente sem travamentos."""
        # Cancela timer anterior para evitar acúmulo de callbacks
        if hasattr(self, "_history_timer_id") and self._history_timer_id is not None:
            try:
                self.root.after_cancel(self._history_timer_id)
            except Exception:
                pass
            self._history_timer_id = None

        self.config = load_config()
        recent = get_recent_reports()
        send_log = self.config.get("report_send_log", {})

        # Cria assinatura de estado para evitar reconstrução desnecessária de widgets da tela
        current_sig = tuple(
            (it.get("path"), it.get("gerado_em"), send_log.get(it.get("filename", os.path.basename(it.get("path", "")))))
            for it in recent[:3]
        )

        if not force and hasattr(self, "_last_history_sig") and self._last_history_sig == current_sig:
            # Estado idêntico; agenda próxima checagem em 30s sem recriar widgets
            self._history_timer_id = self.root.after(30000, lambda: self._refresh_history(force=False))
            return

        self._last_history_sig = current_sig

        for child in self.frame_history_list.winfo_children():
            child.destroy()

        if not recent:
            lbl_empty = tk.Label(
                self.frame_history_list,
                text="Nenhum relatório recente gerado nesta máquina ainda.",
                font=("Segoe UI", 8, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
            )
            lbl_empty.pack(anchor=tk.W, pady=4, padx=4)
            self._history_timer_id = self.root.after(30000, lambda: self._refresh_history(force=False))
            return

        for idx, item in enumerate(recent[:3]):
            f_path = item.get("path", "")
            f_name = item.get("filename", os.path.basename(f_path))
            dt_ger = item.get("gerado_em", "")

            # Padronizar nome: retirar .xlsx e permitir exibição do nome completo sem cortes
            clean_name = f_name[:-5] if f_name.lower().endswith(".xlsx") else f_name
            max_len = 38
            disp_name = clean_name[:max_len] + "..." if len(clean_name) > max_len else clean_name

            row_frame = tk.Frame(self.frame_history_list, bg=COLOR_BG_LIGHT)
            row_frame.pack(fill=tk.X, pady=3, padx=2)

            # Empacotar botões de ação à direita na ordem visual correta:
            # 1º pack: E-mail (fica na ponta direita)
            # 2º pack: PDF (ao lado do E-mail, se existir)
            # 3º pack: Excel (ao lado esquerdo do PDF)
            # 4º pack: Excluir (ao lado esquerdo do Excel)
            pdf_p = item.get("pdf_path", "")
            has_pdf = item.get("has_pdf") and os.path.exists(pdf_p)

            btn_email = ttk.Button(
                row_frame, text="E-mail", width=6,
                command=lambda x=f_path, p=pdf_p: self._send_email_action(x, p),
                style="History.TButton"
            )
            btn_email.pack(side=tk.RIGHT, padx=(4, 0))

            if has_pdf:
                btn_pdf = ttk.Button(
                    row_frame, text="PDF", width=5,
                    command=lambda p=pdf_p: self._open_specific_file(p),
                    style="History.TButton"
                )
                btn_pdf.pack(side=tk.RIGHT, padx=(4, 0))

            btn_open = ttk.Button(
                row_frame, text="Excel", width=6,
                command=lambda p=f_path: self._open_specific_file(p),
                style="History.TButton"
            )
            btn_open.pack(side=tk.RIGHT, padx=(4, 0))

            btn_del = ttk.Button(
                row_frame, text="Excluir", width=6,
                command=lambda x=f_path, p=pdf_p: self._delete_specific_file(x, p),
                style="Delete.TButton"
            )
            btn_del.pack(side=tk.RIGHT, padx=(4, 0))

            # Lado esquerdo: Nome em negrito com espaço amplo para não cortar nenhum mês
            lbl_left = tk.Frame(row_frame, bg=COLOR_BG_LIGHT)
            lbl_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

            lbl_f = tk.Label(
                lbl_left, text=f"•  {disp_name}", font=("Segoe UI", 9, "bold"),
                fg=COLOR_TEXT_MAIN, bg=COLOR_BG_LIGHT,
                width=28, anchor="w"
            )
            lbl_f.pack(side=tk.LEFT)

            # Data de geração posicionada mais ao meio
            if dt_ger:
                lbl_d = tk.Label(
                    lbl_left, text=f"({dt_ger})", font=("Segoe UI", 8),
                    fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
                )
                lbl_d.pack(side=tk.LEFT, padx=(18, 0))

            sent_time = send_log.get(f_name)
            if sent_time:
                lbl_sent = tk.Label(
                    lbl_left, text=f"• ✉ Enviado em {sent_time}", font=("Segoe UI", 8, "italic"),
                    fg="#2D6B22", bg=COLOR_BG_LIGHT
                )
                lbl_sent.pack(side=tk.LEFT, padx=(10, 0))

        # Agenda próxima checagem periódica em 30 segundos
        self._history_timer_id = self.root.after(30000, lambda: self._refresh_history(force=False))

    def _open_specific_file(self, path):
        """Abre com segurança um arquivo do histórico recente."""
        if not os.path.exists(path):
            messagebox.showwarning(
                "Arquivo Não Encontrado",
                f"O arquivo não foi localizado:\n{path}\n\nEle pode ter sido movido ou excluído.",
                parent=self.root
            )
            self._refresh_history(force=True)
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo:\n{e}", parent=self.root)

    def _send_email_action(self, xlsx_path, pdf_path=None):
        """Abre o diálogo de envio de e-mail e atualiza o histórico ao fechar."""
        if not os.path.exists(xlsx_path):
            messagebox.showwarning(
                "Arquivo Não Encontrado",
                f"O arquivo não foi localizado:\n{xlsx_path}\n\nEle pode ter sido movido ou excluído.",
                parent=self.root
            )
            self._refresh_history(force=True)
            return
        dlg = SendEmailDialog(self.root, xlsx_path, pdf_path)
        self.root.wait_window(dlg)
        self._refresh_history(force=True)

    def _delete_specific_file(self, xlsx_path, pdf_path=None):
        """Solicita confirmação e exclui o relatório (Excel e PDF associado) com segurança."""
        f_name = os.path.basename(xlsx_path)
        resp = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja realmente excluir este relatório?\n\n'{f_name}'\n\n(A planilha Excel e a cópia em PDF serão removidas do disco)",
            parent=self.root,
            icon="warning"
        )
        if not resp:
            return

        try:
            # Exclui o arquivo Excel
            if os.path.exists(xlsx_path):
                os.remove(xlsx_path)

            # Exclui o PDF correspondente se existir
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            else:
                cand_pdf = os.path.splitext(xlsx_path)[0] + ".pdf"
                if os.path.exists(cand_pdf):
                    os.remove(cand_pdf)

            # Se a subpasta do mês ficou vazia (ex: 2026\10.26), limpa a subpasta
            parent_dir = os.path.dirname(xlsx_path)
            if os.path.exists(parent_dir) and not os.listdir(parent_dir):
                try:
                    os.rmdir(parent_dir)
                except Exception:
                    pass

            messagebox.showinfo("Sucesso", f"O relatório '{f_name}' foi excluído com sucesso.", parent=self.root)
        except PermissionError:
            messagebox.showerror(
                "Arquivo Aberto",
                f"Não foi possível excluir o arquivo porque ele está aberto no Excel.\nFeche o Excel e tente novamente.",
                parent=self.root
            )
        except Exception as e:
            messagebox.showerror("Erro ao Excluir", f"Falha ao excluir o arquivo:\n{e}", parent=self.root)
        finally:
            self._refresh_history()

    def _update_progress(self, percent, message):
        """Atualiza a barra de progresso e o status da interface de forma suave."""
        self.prog_bar["value"] = percent
        self.lbl_status.config(text=message)

    def _test_connection_action(self):
        """Testa a conexão com o banco de dados em segundo plano sem travar a interface."""
        self.lbl_status.config(text="1/2: Testando conexão com o SQL Server (StruxureWare)...")
        self.prog_bar["value"] = 30
        self.root.update_idletasks()

        def _test_worker():
            ok, msg, drv = test_db_connection(self.config)
            def _ui_done():
                self.prog_bar["value"] = 100 if ok else 0
                if ok:
                    self.lbl_status.config(text="Conexão com o banco bem-sucedida.")
                    messagebox.showinfo("Conexão OK", msg, parent=self.root)
                else:
                    self.lbl_status.config(text="Erro ao conectar ao banco de dados.")
                    messagebox.showerror("Falha na Conexão", msg, parent=self.root)
            self.root.after(0, _ui_done)

        threading.Thread(target=_test_worker, daemon=True).start()

    def _open_output_folder(self):
        folder = self.lbl_pasta.get().strip()
        if not os.path.exists(folder):
            try:
                os.makedirs(folder, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar a pasta:\n{e}", parent=self.root)
                return
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Erro ao Abrir Pasta", f"Não foi possível abrir a pasta:\n{e}", parent=self.root)

    def _start_processing(self):
        # Validação das datas pelo DateEntry
        try:
            d_ini_date = self.cal_inicio.get_date()
            d_fim_date = self.cal_fim.get_date()
            if d_ini_date > d_fim_date:
                messagebox.showerror("Erro de Período", "A data inicial não pode ser posterior à data final.", parent=self.root)
                return
            d_ini_sql = f"{d_ini_date.strftime('%Y-%m-%d')} 00:00:00"
            d_fim_sql = f"{d_fim_date.strftime('%Y-%m-%d')} 23:59:59"
        except Exception as e:
            messagebox.showerror("Erro de Data", f"Falha ao ler datas do calendário: {e}", parent=self.root)
            return

        try:
            v_m3 = float(self.ent_valor.get().strip().replace(",", "."))
            if v_m3 <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Erro de Preenchimento", "Por favor, insira um valor válido para o m³.", parent=self.root)
            return

        base_out = self.lbl_pasta.get().strip()
        if not base_out:
            base_out = os.path.join(os.path.expanduser("~"), "Documents", "Relatorios_Hidrometros")

        # Organização automática por subpastas: Ano / Mês
        target_dir = get_report_output_folder(base_out, d_fim_date)
        default_filename = format_report_filename(d_fim_date)
        output_path = os.path.join(target_dir, default_filename)

        # Salvar o último valor do m3 no config
        self.config["default_m3_price"] = v_m3
        save_config(self.config)

        # Bloquear botão e iniciar etapas com barra de progresso
        self.btn_gerar.config(state=tk.DISABLED)
        self.prog_bar["value"] = 5
        self.lbl_status.config(text="Iniciando processamento da medição...")

        sort_choice = self.var_sort_desc.get()
        threading.Thread(
            target=self._worker_thread,
            args=(d_ini_sql, d_fim_sql, v_m3, output_path, sort_choice),
            daemon=True
        ).start()

    def _worker_thread(self, d_ini, d_fim, v_m3, output_path, sort_by_consumption=True):
        try:
            def prog_cb(pct, msg):
                self.root.after(0, lambda: self._update_progress(pct, msg))

            final_file, warnings, anomalies = execute_extraction(
                d_ini, d_fim, v_m3, output_path, self.config,
                sort_by_consumption=sort_by_consumption,
                progress_callback=prog_cb
            )
            self.root.after(0, lambda: self._on_success(final_file, warnings, anomalies))
        except PermissionError as pe:
            self.root.after(0, lambda: self._on_error(str(pe)))
        except Exception as e:
            self.root.after(0, lambda: self._on_error(str(e)))

    def _on_success(self, final_file, warnings=None, anomalies=None):
        self.prog_bar["value"] = 100
        self.btn_gerar.config(state=tk.NORMAL)

        pdf_file = os.path.splitext(final_file)[0] + ".pdf"
        has_pdf = os.path.exists(pdf_file)
        status_txt = f"Relatório concluído com sucesso: {os.path.basename(final_file)}"
        if has_pdf:
            status_txt += " (+ PDF)"
        self.lbl_status.config(text=status_txt)
        self._refresh_history()

        # Exibir avisos não-fatais (ex: gráficos não gerados)
        if warnings:
            warning_text = "\n".join(f"• {w}" for w in warnings)
            messagebox.showwarning(
                "Aviso",
                f"O relatório foi gerado, porém com os seguintes avisos:\n\n{warning_text}",
                parent=self.root
            )

        # Auditoria interna: alertar apenas se houver salas suspeitas
        if anomalies:
            try:
                p_str = f"{self.cal_inicio.get_date().strftime('%d/%m/%Y')} a {self.cal_fim.get_date().strftime('%d/%m/%Y')}"
            except Exception:
                p_str = ""

            resp_anom = messagebox.askyesno(
                "🔍 Auditoria de Consumo — Alertas",
                f"Relatório gerado com sucesso!\n\n"
                f"⚠️ O detector encontrou {len(anomalies)} sala(s) com consumo atípico ou suspeita de vazamento.\n\n"
                f"Deseja conferir a lista de suspeitas agora?",
                parent=self.root
            )
            if resp_anom:
                AnomalyDialog(self.root, anomalies, p_str)

        # Oferecer envio imediato por e-mail para fluxo contínuo
        pdf_line = f"\n• PDF: {os.path.basename(pdf_file)}" if has_pdf else ""
        resp_email = messagebox.askyesno(
            "Relatório Concluído — Enviar por E-mail",
            f"Relatório gerado com sucesso!\n\n"
            f"• Planilha: {os.path.basename(final_file)}{pdf_line}\n\n"
            f"Deseja abrir a tela de e-mail para conferir os destinatários e enviar agora?",
            parent=self.root,
            default=messagebox.YES
        )
        if resp_email:
            # Foco 100% no e-mail: o Excel não é aberto para não disputar a tela nem roubar foco
            self._send_email_action(final_file, pdf_file if has_pdf else None)
        else:
            # Usuário optou por não enviar e-mail agora: abre o Excel se a opção estiver marcada
            if self.var_open_excel.get():
                try:
                    os.startfile(final_file)
                except Exception as e:
                    messagebox.showwarning("Aviso", f"Relatório gerado em:\n{final_file}\n\nNão foi possível abrir o Excel automaticamente: {e}", parent=self.root)
            else:
                folder_dir = os.path.dirname(final_file)
                resp_folder = messagebox.askyesno(
                    "Abrir Pasta",
                    f"Relatório salvo em:\n{folder_dir}\n\nDeseja abrir a pasta agora?",
                    parent=self.root
                )
                if resp_folder:
                    try:
                        os.startfile(folder_dir)
                    except Exception:
                        pass

    def _on_error(self, err_msg):
        self.prog_bar["value"] = 0
        self.btn_gerar.config(state=tk.NORMAL)
        self.lbl_status.config(text="Ocorreu um erro durante a geração do relatório.")
        messagebox.showerror("Erro de Execução", f"Falha ao gerar relatório:\n\n{err_msg}", parent=self.root)

def start_gui():
    root = tk.Tk()
    app = AppHidrometrosWindow(root)
    root.mainloop()
