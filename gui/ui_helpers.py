import os
import sys
import tkinter as tk
from tkinter import ttk

# ─── PALETA CORPORATIVA COMPASSS SMARTHYDRO ───
COLOR_PRIMARY = "#3D6B24"          # Verde Corporativo CompaSSS
COLOR_PRIMARY_HOVER = "#2D501A"    # Verde Escuro (Hover)
COLOR_ACCENT = "#78A65A"           # Verde Médio
COLOR_BG_LIGHT = "#F4F7F2"         # Fundo Geral da Aplicação
COLOR_CARD_BG = "#EBF3E6"          # Fundo Suave para Cards e Botões Secundários
COLOR_CARD_BORDER = "#C5DCBA"      # Borda Sutil
COLOR_TEXT_MAIN = "#1E2B18"        # Texto Principal
COLOR_TEXT_MUTED = "#55664C"       # Texto Secundário / Labels
COLOR_DANGER = "#D9534F"           # Vermelho Ação Destrutiva
COLOR_DANGER_BG = "#FDF2F2"        # Fundo Vermelho Suave
COLOR_DANGER_BORDER = "#F5C6CB"    # Borda Vermelha Suave


def setup_common_styles(style: ttk.Style):
    """Configura o tema ttk (clam) globalmente com as cores e formas corporativas da CompaSSS."""
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(".", background=COLOR_BG_LIGHT, font=("Segoe UI", 9))
    style.configure("TLabel", background=COLOR_BG_LIGHT, foreground=COLOR_TEXT_MAIN)
    style.configure("TLabelframe", background=COLOR_BG_LIGHT, bordercolor=COLOR_ACCENT)
    style.configure("TLabelframe.Label", background=COLOR_BG_LIGHT, foreground=COLOR_PRIMARY, font=("Segoe UI", 9, "bold"))
    style.configure("TCheckbutton", background=COLOR_BG_LIGHT, foreground=COLOR_TEXT_MAIN)
    style.configure("TRadiobutton", background=COLOR_BG_LIGHT, foreground=COLOR_TEXT_MAIN)
    style.configure("TProgressbar", troughcolor="#E3EDD8", background=COLOR_ACCENT)

    # Botão padrão ttk (Secondary look)
    style.configure(
        "TButton",
        font=("Segoe UI", 9, "bold"),
        background=COLOR_CARD_BG,
        foreground=COLOR_PRIMARY,
        borderwidth=1,
        bordercolor=COLOR_CARD_BORDER,
        focuscolor="none",
        padding=(10, 4)
    )
    style.map(
        "TButton",
        background=[("active", "#D3E4CB"), ("disabled", "#F0F0F0")],
        foreground=[("active", COLOR_PRIMARY), ("disabled", "#A0A0A0")],
        bordercolor=[("active", COLOR_PRIMARY)]
    )

    # Botão de Ação Primária ttk (CTA)
    style.configure(
        "Primary.TButton",
        font=("Segoe UI", 9, "bold"),
        background=COLOR_PRIMARY,
        foreground="#FFFFFF",
        borderwidth=1,
        bordercolor=COLOR_PRIMARY,
        focuscolor="none",
        padding=(14, 5)
    )
    style.map(
        "Primary.TButton",
        background=[("active", COLOR_PRIMARY_HOVER), ("disabled", "#A0B896")],
        foreground=[("disabled", "#E0E0E0")]
    )

    # Botão de Ação Secundária ttk
    style.configure(
        "Secondary.TButton",
        font=("Segoe UI", 9, "bold"),
        background=COLOR_CARD_BG,
        foreground=COLOR_PRIMARY,
        borderwidth=1,
        bordercolor=COLOR_CARD_BORDER,
        focuscolor="none",
        padding=(10, 4)
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#D3E4CB"), ("disabled", "#F0F0F0")],
        foreground=[("active", COLOR_PRIMARY), ("disabled", "#A0A0A0")]
    )

    # Botão de Histórico e Ações Compactas
    style.configure(
        "History.TButton",
        font=("Segoe UI", 8, "bold"),
        background=COLOR_CARD_BG,
        foreground=COLOR_PRIMARY,
        borderwidth=1,
        bordercolor=COLOR_CARD_BORDER,
        focuscolor="none",
        padding=(6, 2)
    )
    style.map(
        "History.TButton",
        background=[("active", "#D3E4CB"), ("disabled", "#F0F0F0")],
        foreground=[("active", COLOR_PRIMARY), ("disabled", "#A0A0A0")]
    )

    # Botões de Abas do Notebook
    style.configure(
        "TNotebook",
        background=COLOR_BG_LIGHT,
        borderwidth=0
    )
    style.configure(
        "TNotebook.Tab",
        font=("Segoe UI", 9, "bold"),
        background="#E6ECE0",
        foreground=COLOR_TEXT_MUTED,
        padding=[10, 5],
        borderwidth=1,
        bordercolor=COLOR_CARD_BORDER
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLOR_BG_LIGHT), ("active", "#DDE7D6")],
        foreground=[("selected", COLOR_PRIMARY), ("active", COLOR_PRIMARY)]
    )


def create_btn_primary(parent, text, command, **kwargs):
    """Cria um botão primário com identidade corporativa verde CompaSSS."""
    pad_x = kwargs.pop("padx", 16)
    pad_y = kwargs.pop("pady", 6)
    btn = tk.Button(
        parent, text=text, command=command,
        bg=COLOR_PRIMARY, fg="white",
        activebackground=COLOR_PRIMARY_HOVER, activeforeground="white",
        font=("Segoe UI", 9, "bold"),
        relief="flat", bd=1,
        highlightbackground=COLOR_PRIMARY, highlightthickness=1,
        padx=pad_x, pady=pad_y, cursor="hand2", **kwargs
    )
    return btn


def create_btn_secondary(parent, text, command, **kwargs):
    """Cria um botão secundário suave, com fundo claro e borda verde sutil."""
    pad_x = kwargs.pop("padx", 12)
    pad_y = kwargs.pop("pady", 6)
    btn = tk.Button(
        parent, text=text, command=command,
        bg=COLOR_CARD_BG, fg=COLOR_PRIMARY,
        activebackground="#D3E4CB", activeforeground=COLOR_PRIMARY,
        font=("Segoe UI", 9, "bold"),
        relief="flat", bd=1,
        highlightbackground=COLOR_CARD_BORDER, highlightthickness=1,
        padx=pad_x, pady=pad_y, cursor="hand2", **kwargs
    )
    return btn


def create_btn_danger(parent, text, command, **kwargs):
    """Cria um botão de perigo/exclusão com fundo avermelhado suave."""
    pad_x = kwargs.pop("padx", 10)
    pad_y = kwargs.pop("pady", 5)
    btn = tk.Button(
        parent, text=text, command=command,
        bg=COLOR_DANGER_BG, fg="#C9302C",
        activebackground="#FADBD8", activeforeground="#C9302C",
        font=("Segoe UI", 8, "bold"),
        relief="flat", bd=1,
        highlightbackground=COLOR_DANGER_BORDER, highlightthickness=1,
        padx=pad_x, pady=pad_y, cursor="hand2", **kwargs
    )
    return btn


def apply_window_icon(window):
    """
    Aplica o ícone oficial da CompaSSS na janela (eliminando o ícone de pena padrão do Tkinter).
    """
    try:
        # Procurar arquivo .ico
        candidates = []
        if hasattr(sys, '_MEIPASS'):
            candidates.append(os.path.join(sys._MEIPASS, "app_icon.ico"))
        if getattr(sys, 'frozen', False):
            candidates.append(os.path.join(os.path.dirname(sys.executable), "app_icon.ico"))
        candidates.extend([
            os.path.join(os.getcwd(), "app_icon.ico"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "app_icon.ico"),
        ])

        ico_applied = False
        for p in candidates:
            if os.path.exists(p):
                try:
                    window.iconbitmap(p)
                    ico_applied = True
                    break
                except Exception:
                    pass

        # Também aplicar iconphoto com PNG para garantir renderização nítida
        png_candidates = []
        if hasattr(sys, '_MEIPASS'):
            png_candidates.append(os.path.join(sys._MEIPASS, "logo_final.png"))
            png_candidates.append(os.path.join(sys._MEIPASS, "gui_logo.png"))
        if getattr(sys, 'frozen', False):
            png_candidates.append(os.path.join(os.path.dirname(sys.executable), "logo_final.png"))
            png_candidates.append(os.path.join(os.path.dirname(sys.executable), "gui_logo.png"))
        png_candidates.extend([
            os.path.join(os.getcwd(), "logo_final.png"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "logo_final.png"),
            os.path.join(os.getcwd(), "gui_logo.png"),
        ])

        for p in png_candidates:
            if os.path.exists(p):
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(p).resize((32, 32), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    window._dialog_icon_ref = photo
                    window.iconphoto(True, photo)
                    return True
                except Exception:
                    pass

        return ico_applied
    except Exception:
        pass
    return False


def center_modal(dialog, parent, width, height):
    """
    Centraliza o diálogo perfeitamente sobre a janela-mãe (parent),
    garantindo margens simétricas e a sensação visual de estar 'por cima do software'.
    """
    try:
        dialog.update_idletasks()
        parent.update_idletasks()
        p_x = parent.winfo_x()
        p_y = parent.winfo_y()
        p_w = parent.winfo_width()
        p_h = parent.winfo_height()

        x = p_x + (p_w - width) // 2
        y = p_y + (p_h - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
    except Exception:
        dialog.geometry(f"{width}x{height}")


class ToolTip:
    """
    Exibe um tooltip leve e moderno ao passar o mouse sobre um widget Tkinter.
    """
    def __init__(self, widget, text, delay=350):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self._after_id = None

        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, event=None):
        self._schedule()

    def _on_leave(self, event=None):
        self._cancel()
        self._hide()

    def _schedule(self):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self.tip_window or not self.text:
            return
        try:
            import tkinter as tk
            x = self.widget.winfo_rootx() + (self.widget.winfo_width() // 2)
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self.tip_window = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            tw.attributes("-topmost", True)
            label = tk.Label(
                tw, text=self.text, justify=tk.LEFT,
                background="#2B3E22", foreground="#FFFFFF",
                relief=tk.FLAT, borderwidth=0,
                padx=8, pady=4,
                font=("Segoe UI", 8)
            )
            label.pack()
        except Exception:
            pass

    def _hide(self):
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


def create_tooltip(widget, text):
    """Cria e anexa um tooltip moderno ao widget Tkinter."""
    return ToolTip(widget, text)


class SlidingSegmentedSwitch(tk.Canvas):
    """
    Controle deslizante moderno com efeito de pílula (segmented pill slider).
    Renderiza trilho e indicador deslizante anti-aliasing via Pillow,
    proporcionando transições suaves a 60 FPS com estilo corporativo CompaSSS.
    """
    def __init__(self, parent, options, initial_idx=0, command=None, width=470, height=36, bg_color=COLOR_BG_LIGHT):
        super().__init__(parent, width=width, height=height, bg=bg_color, highlightthickness=0, cursor="hand2")
        self.options = options
        self.current_idx = initial_idx
        self.command = command
        self.w = width
        self.h = height
        self.n = len(options)
        self.padding = 3
        self.seg_w = (width - 2 * self.padding) / self.n
        self.pill_w = int(self.seg_w)
        self.pill_h = height - 2 * self.padding
        self.thumb_x = self.padding + initial_idx * self.seg_w
        self.target_x = self.thumb_x
        self._animating = False

        from PIL import Image, ImageDraw, ImageTk

        # 1. Trilha do fundo suave (anti-aliasing)
        im_track = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        d_tr = ImageDraw.Draw(im_track)
        d_tr.rounded_rectangle([1, 1, width - 2, height - 2], radius=height // 2, fill="#E8F1E4", outline="#C5DCBA", width=1)
        self.photo_track = ImageTk.PhotoImage(im_track)

        # 2. Pílula deslizante ativa verde CompaSSS (anti-aliasing)
        im_pill = Image.new("RGBA", (self.pill_w, self.pill_h), (0, 0, 0, 0))
        d_pill = ImageDraw.Draw(im_pill)
        d_pill.rounded_rectangle([0, 0, self.pill_w - 1, self.pill_h - 1], radius=self.pill_h // 2, fill=COLOR_PRIMARY)
        self.photo_pill = ImageTk.PhotoImage(im_pill)

        # Desenhar no canvas
        self.create_image(0, 0, image=self.photo_track, anchor="nw")
        self.pill_item = self.create_image(self.thumb_x, self.padding, image=self.photo_pill, anchor="nw")

        # Rótulos de texto com ícones
        self.text_items = []
        for i, text in enumerate(self.options):
            cx = self.padding + i * self.seg_w + self.seg_w / 2
            cy = height / 2
            color = "#FFFFFF" if (i == initial_idx) else COLOR_PRIMARY
            t_item = self.create_text(cx, cy, text=text, fill=color, font=("Segoe UI", 9, "bold"))
            self.text_items.append(t_item)

        self.bind("<Button-1>", self._on_click)

    def _on_click(self, event):
        idx = int((event.x - self.padding) // self.seg_w)
        idx = max(0, min(self.n - 1, idx))
        if idx != self.current_idx:
            self.set_index(idx)

    def get_index(self):
        return self.current_idx

    def set_index(self, idx, trigger_cmd=True):
        if idx == self.current_idx and abs(self.target_x - self.thumb_x) < 1:
            return
        self.current_idx = idx
        self.target_x = self.padding + idx * self.seg_w
        self._update_text_colors()
        if not self._animating:
            self._animating = True
            self._animate_step()
        if trigger_cmd and self.command:
            self.command(self.current_idx)

    def _update_text_colors(self):
        for i, t_item in enumerate(self.text_items):
            color = "#FFFFFF" if (i == self.current_idx) else COLOR_PRIMARY
            self.itemconfig(t_item, fill=color)

    def _animate_step(self):
        dx = self.target_x - self.thumb_x
        if abs(dx) < 1.5:
            self.thumb_x = self.target_x
            self.coords(self.pill_item, self.thumb_x, self.padding)
            self._animating = False
        else:
            self.thumb_x += dx * 0.35
            self.coords(self.pill_item, self.thumb_x, self.padding)
            self.after(16, self._animate_step)


