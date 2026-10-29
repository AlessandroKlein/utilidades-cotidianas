# -*- coding: utf-8 -*-
r"""Lanzador gráfico de utilidades-cotidianas.

Una ventana única que descubre automáticamente las herramientas del repo
(cada subcarpeta con un único script .py) y las ejecuta con una ruta de
destino + argumentos extra, mostrando la salida en vivo.

Uso:
    python lanzador.py            # abre la ventana
    python lanzador.py --selftest # lista las herramientas detectadas (sin ventana)

Sin dependencias externas: usa tkinter, incluido en la instalación estándar
de Python para Windows.
"""
import os
import shlex
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PYTHON = sys.executable


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

        self.by_display = {
            humanize(name): script for name, script in self.tools
        }

        root.title("Utilidades Cotidianas — Lanzador")
        root.geometry("780x560")
        root.minsize(640, 460)

        pad = {"padx": 10, "pady": 4}

        # Herramienta
        ttk.Label(root, text="Herramienta:").grid(
            row=0, column=0, sticky="w", **pad
        )
        self.tool_var = tk.StringVar()
        self.tool_combo = ttk.Combobox(
            root, textvariable=self.tool_var, state="readonly",
            values=list(self.by_display), width=48,
        )
        self.tool_combo.grid(row=0, column=1, columnspan=2, sticky="ew", **pad)
        self.tool_combo.current(0)
        self.tool_combo.bind("<<ComboboxSelected>>", self._on_tool)

        # Descripción
        self.desc_label = ttk.Label(
            root, text="", wraplength=740, foreground="#555"
        )
        self.desc_label.grid(row=1, column=0, columnspan=3, sticky="w", **pad)

        # Ruta de destino
        ttk.Label(root, text="Carpeta/archivo:").grid(
            row=2, column=0, sticky="w", **pad
        )
        self.target_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.target_var).grid(
            row=2, column=1, sticky="ew", **pad
        )
        btns = ttk.Frame(root)
        btns.grid(row=2, column=2, sticky="e", **pad)
        ttk.Button(btns, text="Carpeta…", command=self._pick_folder).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btns, text="Archivo…", command=self._pick_file).pack(
            side="left"
        )

        # Argumentos extra
        ttk.Label(root, text="Argumentos extra:").grid(
            row=3, column=0, sticky="w", **pad
        )
        self.args_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.args_var).grid(
            row=3, column=1, sticky="ew", **pad
        )
        ttk.Label(
            root, text="opcional", foreground="#999"
        ).grid(row=3, column=2, sticky="w", **pad)
        ttk.Label(
            root,
            text="Consejo: corré primero sin argumentos para ver el plan "
                 "(dry-run); para aplicar agregá --aplicar/--borrar y --si.",
            wraplength=740, foreground="#555",
        ).grid(row=4, column=0, columnspan=3, sticky="w", **pad)

        # Botón ejecutar
        self.run_btn = ttk.Button(root, text="Ejecutar", command=self.run_tool)
        self.run_btn.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)

        # Salida
        self.output = scrolledtext.ScrolledText(
            root, height=18, state="disabled", font=("Consolas", 9)
        )
        self.output.grid(row=6, column=0, columnspan=3, sticky="nsew", **pad)

        root.columnconfigure(1, weight=1)
        root.rowconfigure(6, weight=1)

        self._on_tool()

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
