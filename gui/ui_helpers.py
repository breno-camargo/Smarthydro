import os
import sys

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


