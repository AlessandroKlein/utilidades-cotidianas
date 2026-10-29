# -*- coding: utf-8 -*-
r"""Lanzador gráfico de utilidades-cotidianas.

Una ventana única que descubre automáticamente las herramientas del repo
(cada subcarpeta con un único script .py) y las ejecuta con una ruta de
destino + argumentos extra, mostrando la salida en vivo.

Sigue el tema claro/oscuro del sistema (Windows) y usa una paleta de colores
moderna, todo con tkinter (sin dependencias externas).

Uso:
    python lanzador.py            # abre la ventana
    python lanzador.py --selftest # lista las herramientas detectadas (sin ventana)
"""
import os
import platform
import shlex
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PYTHON = sys.executable

FONT_UI = ("Segoe UI", 10)
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SUB = ("Segoe UI", 10)
FONT_MONO = ("Consolas", 9)

PALETTES = {
    "light": {
        "window": "#eef1f6",
        "panel": "#ffffff",
        "field": "#ffffff",
        "text": "#1c2430",
        "title": "#111827",
        "muted": "#6b7280",
        "border": "#dbe0e7",
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "accent_fg": "#ffffff",
        "button": "#f2f4f8",
        "button_hover": "#e4e9f2",
        "disabled": "#dfe3ea",
        "disabled_fg": "#98a1ad",
        "trough": "#e9ecf2",
        "scrollbar": "#c6ccd6",
        "scrollbar_hover": "#aab3c0",
        "output_bg": "#f8fafc",
        "output_fg": "#1c2430",
    },
    "dark": {
        "window": "#0e1116",
        "panel": "#161b23",
        "field": "#1c222c",
        "text": "#e7eaef",
        "title": "#ffffff",
        "muted": "#98a2b1",
        "border": "#2a323d",
        "accent": "#3b82f6",
        "accent_hover": "#2563eb",
        "accent_fg": "#ffffff",
        "button": "#222936",
        "button_hover": "#2c3443",
        "disabled": "#252b35",
        "disabled_fg": "#5b6471",
        "trough": "#1b212a",
        "scrollbar": "#343c49",
        "scrollbar_hover": "#434d5c",
        "output_bg": "#12161c",
        "output_fg": "#e7eaef",
    },
}


def detect_system_dark():
    """True si el sistema usa tema oscuro (Windows). Fallback: claro."""
    if platform.system() == "Windows":
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            return value == 0
        except OSError:
            return False
    return False


def discover_tools(root):
    """Devuelve [(carpeta, script)] por cada subcarpeta con un único .py."""
    tools = []
    for entry in sorted(os.listdir(root)):
        folder = os.path.join(root, entry)
        if not os.path.isdir(folder):
            continue
        scripts = sorted(
            f for f in os.listdir(folder)
            if f.endswith(".py") and not f.startswith("_")
        )
        if len(scripts) == 1:
            tools.append((entry, os.path.join(folder, scripts[0])))
    return tools


def describe(script):
    """Lee la primera línea de la docstring del script (su descripción)."""
    try:
        with open(script, encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s.startswith('r"""'):
                    desc = s[4:]
                elif s.startswith('"""'):
                    desc = s[3:]
                else:
                    continue
                if desc.endswith('"""'):
                    desc = desc[:-3]
                return desc.strip()
    except OSError:
        pass
    return ""


def humanize(folder):
    return folder.replace("-", " ").replace("_", " ").title()


class App:
    def __init__(self, root):
        self.root = root
        self.tools = discover_tools(REPO_ROOT)
        if not self.tools:
            messagebox.showerror(
                "Sin herramientas",
                "No encontré herramientas (carpetas con un script .py) "
                f"en:\n{REPO_ROOT}",
            )
            root.destroy()
            return

        self.c = PALETTES["dark" if detect_system_dark() else "light"]
        self.by_display = {
            humanize(name): script for name, script in self.tools
        }

        root.title("Utilidades Cotidianas — Lanzador")
        root.geometry("780x600")
        root.minsize(640, 500)
        root.configure(bg=self.c["window"])

        self._build_style()
        self._build_header()
        self._build_card()
        self._build_output()
        self._on_tool()

    # ------------------------------------------------------------------ tema
    def _build_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        c = self.c

        style.configure(".", background=c["panel"], foreground=c["text"],
                        bordercolor=c["border"], lightcolor=c["border"],
                        darkcolor=c["border"], troughcolor=c["trough"],
                        focuscolor=c["accent"], font=FONT_UI)
        style.configure("TFrame", background=c["panel"])
        style.configure("TLabel", background=c["panel"], foreground=c["text"])
        style.configure("Muted.TLabel", foreground=c["muted"])
        style.configure("TEntry", fieldbackground=c["field"],
                        foreground=c["text"], insertcolor=c["text"],
                        bordercolor=c["border"], lightcolor=c["border"],
                        darkcolor=c["border"])
        style.configure(
            "TCombobox",
            fieldbackground=c["field"], background=c["field"],
            foreground=c["text"], arrowcolor=c["text"],
            bordercolor=c["border"], lightcolor=c["border"],
            darkcolor=c["border"],
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", c["field"])],
            foreground=[("readonly", c["text"])],
            selectbackground=[("readonly", c["field"])],
            selectforeground=[("readonly", c["text"])],
        )
        style.configure("TButton", background=c["button"],
                        foreground=c["text"], borderwidth=0,
                        focusthickness=0, padding=(10, 6))
        style.map("TButton",
                  background=[("active", c["button_hover"]),
                              ("pressed", c["button_hover"])])
        style.configure("Accent.TButton", background=c["accent"],
                        foreground=c["accent_fg"], borderwidth=0,
                        focusthickness=0, padding=(10, 9),
                        font=("Segoe UI", 11, "bold"))
        style.map("Accent.TButton",
                  background=[("active", c["accent_hover"]),
                              ("pressed", c["accent_hover"]),
                              ("disabled", c["disabled"])],
                  foreground=[("disabled", c["disabled_fg"])])
        style.configure("Vertical.TScrollbar", background=c["scrollbar"],
                        troughcolor=c["trough"], bordercolor=c["trough"],
                        arrowcolor=c["muted"])
        style.map("Vertical.TScrollbar",
                  background=[("active", c["scrollbar_hover"])])

    # -------------------------------------------------------------- widgets
    def _build_header(self):
        c = self.c
        header = tk.Frame(self.root, bg=c["window"])
        header.pack(fill="x", padx=24, pady=(18, 2))
        tk.Label(header, text="Utilidades Cotidianas", bg=c["window"],
                 fg=c["title"], font=FONT_TITLE).pack(anchor="w")
        tk.Label(
            header,
            text="Elegí una herramienta, una carpeta o archivo, y ejecutá.",
            bg=c["window"], fg=c["muted"], font=FONT_SUB,
        ).pack(anchor="w", pady=(2, 0))

    def _build_card(self):
        c = self.c
        card = tk.Frame(self.root, bg=c["panel"], highlightthickness=1,
                        highlightbackground=c["border"])
        card.pack(fill="x", padx=24, pady=(10, 8))
        card.columnconfigure(1, weight=1)

        pad = {"padx": 12, "pady": 6}

        ttk.Label(card, text="Herramienta").grid(
            row=0, column=0, sticky="w", **pad)
        self.tool_var = tk.StringVar()
        self.tool_combo = ttk.Combobox(
            card, textvariable=self.tool_var, state="readonly",
            values=list(self.by_display), width=42,
        )
        self.tool_combo.grid(row=0, column=1, columnspan=2,
                             sticky="ew", **pad)
        self.tool_combo.current(0)
        self.tool_combo.bind("<<ComboboxSelected>>", self._on_tool)

        self.desc_label = ttk.Label(card, text="", style="Muted.TLabel",
                                    wraplength=700)
        self.desc_label.grid(row=1, column=0, columnspan=3, sticky="w",
                             padx=12, pady=(0, 4))

        ttk.Label(card, text="Carpeta/archivo").grid(
            row=2, column=0, sticky="w", **pad)
        self.target_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.target_var).grid(
            row=2, column=1, sticky="ew", **pad)
        btns = ttk.Frame(card)
        btns.grid(row=2, column=2, sticky="e", **pad)
        ttk.Button(btns, text="Carpeta…", command=self._pick_folder).pack(
            side="left", padx=(0, 6))
        ttk.Button(btns, text="Archivo…", command=self._pick_file).pack(
            side="left")

        ttk.Label(card, text="Argumentos").grid(
            row=3, column=0, sticky="w", **pad)
        self.args_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.args_var).grid(
            row=3, column=1, sticky="ew", **pad)
        ttk.Label(card, text="opcional", style="Muted.TLabel").grid(
            row=3, column=2, sticky="w", **pad)

        self.run_btn = ttk.Button(card, text="Ejecutar",
                                  style="Accent.TButton",
                                  command=self.run_tool)
        self.run_btn.grid(row=4, column=0, columnspan=3, sticky="ew",
                          padx=12, pady=(10, 12))

    def _build_output(self):
        c = self.c
        ttk.Label(self.root, text="Salida", style="Muted.TLabel").pack(
            anchor="w", padx=24, pady=(0, 4))

        out_frame = ttk.Frame(self.root)
        out_frame.pack(fill="both", expand=True, padx=24, pady=(0, 6))

        self.output = tk.Text(
            out_frame, state="disabled", font=FONT_MONO,
            bg=c["output_bg"], fg=c["output_fg"],
            insertbackground=c["text"], selectbackground=c["accent"],
            selectforeground=c["accent_fg"], relief="flat",
            highlightthickness=1, highlightbackground=c["border"],
            highlightcolor=c["accent"], padx=10, pady=8, wrap="word",
        )
        scroll = ttk.Scrollbar(out_frame, style="Vertical.TScrollbar",
                               command=self.output.yview)
        self.output.configure(yscrollcommand=scroll.set)
        self.output.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        tk.Label(
            self.root,
            text="Consejo: corré primero sin argumentos para ver el plan "
                 "(dry-run); para aplicar agregá --aplicar/--borrar y --si.",
            bg=c["window"], fg=c["muted"], font=FONT_SUB,
        ).pack(anchor="w", padx=24, pady=(2, 12))

    # ------------------------------------------------------------- acciones
    def _on_tool(self, _event=None):
        script = self.by_display.get(self.tool_var.get(), "")
        self.desc_label.configure(text=describe(script) if script else "")

    def _pick_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.target_var.set(path)

    def _pick_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.target_var.set(path)

    def run_tool(self):
        display = self.tool_var.get()
        script = self.by_display.get(display)
        if not script:
            return

        cmd = [PYTHON, script]
        target = self.target_var.get().strip()
        if target:
            cmd.append(target)
        extra = self.args_var.get().strip()
        if extra:
            try:
                cmd += shlex.split(extra)
            except ValueError as e:
                messagebox.showerror(
                    "Argumentos", f"No pude leer los argumentos extra:\n{e}"
                )
                return

        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.insert("end", f"> {subprocess.list2cmdline(cmd)}\n\n")
        self.output.configure(state="disabled")
        self.run_btn.configure(state="disabled")

        threading.Thread(target=self._run, args=(cmd,), daemon=True).start()

    def _run(self, cmd):
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=os.path.dirname(cmd[1]),
        )
        for line in proc.stdout:
            self.root.after(0, self._append, line)
        proc.wait()
        self.root.after(0, self._done, proc.returncode)

    def _append(self, text):
        self.output.configure(state="normal")
        self.output.insert("end", text)
        self.output.see("end")
        self.output.configure(state="disabled")

    def _done(self, code):
        self._append(f"\n[Terminó con código {code}]\n")
        self.run_btn.configure(state="normal")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        tools = discover_tools(REPO_ROOT)
        for name, script in tools:
            print(f"{name}\t{describe(script)}")
        assert tools, "No se encontraron herramientas"
        print(f"OK: {len(tools)} herramientas detectadas")
        sys.exit(0)
    main()
