import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from core.config_manager import get_base_dir
from gui.ui_helpers import (
    apply_window_icon, center_modal, setup_common_styles,
    create_btn_primary, create_btn_secondary, create_btn_danger,
    COLOR_PRIMARY, COLOR_PRIMARY_HOVER, COLOR_ACCENT, COLOR_BG_LIGHT,
    COLOR_TEXT_MAIN, COLOR_TEXT_MUTED
)

TASK_NAME = "ExtracaoHidrometrosPamplona"

class SchedulerDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Agendamento Automático no Windows — CompaSSS")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG_LIGHT)
        self.transient(parent)
        self.grab_set()

        apply_window_icon(self)
        setup_common_styles(ttk.Style(self))

        self._build_ui()
        self._check_existing_task()
        center_modal(self, parent, 530, 400)

    def _build_ui(self):
        pad = ttk.Frame(self, padding=16)
        pad.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            pad, text="⏰ Agendador de Relatório 100% Automático",
            font=("Segoe UI", 11, "bold"), fg=COLOR_PRIMARY, bg=COLOR_BG_LIGHT
        ).pack(anchor=tk.W, pady=(0, 6))

        info_text = (
            "Com este recurso, o Windows executará a extração dos hidrômetros silenciosamente "
            "no dia e horário configurados, gerando a planilha Excel na sua pasta de destino "
            "sem necessidade de abrir o aplicativo manualmente."
        )
        lbl_info = tk.Label(
            pad, text=info_text, wraplength=480, justify=tk.LEFT,
            font=("Segoe UI", 9), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        lbl_info.pack(anchor=tk.W, pady=(0, 15))

        frame_form = ttk.LabelFrame(pad, text="  Parâmetros de Agendamento  ", padding=12)
        frame_form.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(frame_form, text="Dia do Mês para Executar:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.cmb_day = ttk.Combobox(frame_form, values=[str(i) for i in range(1, 29)], width=6, state="readonly", font=("Segoe UI", 9))
        self.cmb_day.set("29")  # Padrão: dia 29 (logo após fechamento do dia 28)
        self.cmb_day.grid(row=0, column=1, sticky=tk.W, pady=5, padx=8)

        ttk.Label(frame_form, text="Horário (HH:MM):", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ent_time = ttk.Entry(frame_form, width=8, font=("Segoe UI", 9))
        self.ent_time.insert(0, "08:00")
        self.ent_time.grid(row=1, column=1, sticky=tk.W, pady=5, padx=8)

        self.lbl_task_status = tk.Label(
            pad, text="Status da Tarefa no Windows: Verificando...",
            font=("Segoe UI", 9, "italic"), fg=COLOR_TEXT_MUTED, bg=COLOR_BG_LIGHT
        )
        self.lbl_task_status.pack(anchor=tk.W, pady=(0, 15))

        frame_btns = tk.Frame(pad, bg=COLOR_BG_LIGHT)
        frame_btns.pack(fill=tk.X, side=tk.BOTTOM)

        self.btn_create = create_btn_primary(
            frame_btns, "Ativar Agendamento",
            self._create_task, pady=6
        )
        self.btn_create.pack(side=tk.LEFT)

        self.btn_remove = create_btn_danger(
            frame_btns, "🗑️ Remover Agendamento",
            self._delete_task, pady=6
        )
        self.btn_remove.pack(side=tk.LEFT, padx=8)

        btn_close = create_btn_secondary(frame_btns, "Fechar", self.destroy, pady=6)
        btn_close.pack(side=tk.RIGHT)

    def _get_target_command(self):
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
            return f'"{exe_path}" --auto'
        else:
            python_exe = sys.executable
            main_py = os.path.join(get_base_dir(), "main.py")
            return f'"{python_exe}" "{main_py}" --auto'

    def _check_existing_task(self):
        try:
            res = subprocess.run(
                ["schtasks", "/Query", "/TN", TASK_NAME],
                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            if res.returncode == 0:
                self.lbl_task_status.config(text=f"Status: ATIVA no Windows ({TASK_NAME})", foreground="green")
                self.btn_create.config(text="Atualizar Agendamento")
            else:
                self.lbl_task_status.config(text="Status: Não configurada no Windows", foreground="#666666")
                self.btn_create.config(text="Ativar Agendamento Automático")
        except Exception as e:
            self.lbl_task_status.config(text=f"Status: Não foi possível consultar ({e})", foreground="red")

    def _create_task(self):
        day = self.cmb_day.get().strip()
        time_str = self.ent_time.get().strip()
        cmd = self._get_target_command()

        # Comando schtasks
        sch_cmd = [
            "schtasks", "/Create", "/F",
            "/TN", TASK_NAME,
            "/TR", cmd,
            "/SC", "MONTHLY",
            "/D", day,
            "/ST", time_str
        ]

        try:
            res = subprocess.run(sch_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if res.returncode == 0:
                messagebox.showinfo("Sucesso", f"Tarefa agendada com sucesso!\nO Windows executará o relatório todo dia {day} do mês às {time_str}.", parent=self)
                self._check_existing_task()
            else:
                err_msg = res.stderr or res.stdout
                messagebox.showerror("Erro ao Criar Agendamento", f"Não foi possível registrar no Windows:\n{err_msg}", parent=self)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao executar schtasks: {str(e)}", parent=self)

    def _delete_task(self):
        try:
            res = subprocess.run(
                ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            if res.returncode == 0:
                messagebox.showinfo("Sucesso", "Agendamento removido com sucesso.", parent=self)
            else:
                messagebox.showinfo("Aviso", "Nenhuma tarefa agendada foi encontrada para remover.", parent=self)
            self._check_existing_task()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao remover tarefa: {str(e)}", parent=self)
