"""
Diálogos e prompts tkinter/pyautogui para AutoMDFText.

Responsabilidades:
- Wrappers de pyautogui que pausam o timer de automação durante interação (focused_*)
- Prompt bloqueante de DT com validação embutida (prompt_dt_blocking)
- Diálogo unificado para CT-e, NF1, NF2 e NCM (prompt_batch_info)

Todos os diálogos pausam o timer de automação durante sua exibição via
mdfe.timing.pause_automation_timer / resume_automation_timer.
"""

import tkinter as tk
from tkinter import messagebox

import pyautogui

from mdfe.timing import pause_automation_timer, resume_automation_timer


def focused_alert(text: str = "", title: str = "", button: str = "OK"):
    """Wrapper para pyautogui.alert."""
    pause_automation_timer()  # Pausar timer durante alert

    try:
        result = pyautogui.alert(text=text, title=title, button=button)
    finally:
        resume_automation_timer()  # Resumir timer após alert

    return result


def focused_confirm(text: str = "", title: str = "", buttons: list[str] | None = None):
    """Wrapper para pyautogui.confirm."""
    pause_automation_timer()

    try:
        result = pyautogui.confirm(text=text, title=title, buttons=buttons)
    finally:
        resume_automation_timer()

    return result


def focused_prompt(text: str = "", title: str = "", default: str = ""):
    """Wrapper para pyautogui.prompt."""
    pause_automation_timer()  # Pausar timer durante prompt

    try:
        result = pyautogui.prompt(text=text, title=title, default=default)
    finally:
        resume_automation_timer()  # Resumir timer após prompt

    return result


def prompt_dt_blocking(text: str, title: str = "DT") -> str | None:
    """Prompt dedicado para DT sem bloqueio global de interacao."""
    pause_automation_timer()
    root = tk.Tk()
    root.withdraw()

    result: str | None = None
    try:
        dialog = tk.Toplevel(root)
        dialog.title(title)
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)
        dialog.geometry("460x200")
        dialog.lift()

        font_label  = ("Segoe UI", 11)
        font_entry  = ("Segoe UI", 11)
        font_button = ("Segoe UI", 11)

        frame = tk.Frame(dialog, padx=16, pady=14)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=text, font=font_label, wraplength=420, justify="left").pack(anchor="w")
        entry_var = tk.StringVar()
        entry = tk.Entry(frame, textvariable=entry_var, font=font_entry)
        entry.pack(fill="x", pady=(8, 12))

        button_frame = tk.Frame(frame)
        button_frame.pack(fill="x")

        def finalize(value: str | None) -> None:
            nonlocal result
            result = value
            dialog.destroy()

        def on_ok() -> None:
            value = entry_var.get().strip()
            if not value:
                messagebox.showwarning(
                    "DT obrigatoria",
                    "A DT precisa ser digitada para continuar.",
                    parent=dialog,
                )
                entry.focus_set()
                return
            finalize(value)

        def on_cancel() -> None:
            finalize(None)

        ok_button     = tk.Button(button_frame, text="OK",       command=on_ok,     width=10, font=font_button)
        ok_button.pack(side="right", padx=(6, 0))
        cancel_button = tk.Button(button_frame, text="Cancelar", command=on_cancel, width=10, font=font_button)
        cancel_button.pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        dialog.bind("<Return>", lambda _e: on_ok())
        dialog.bind("<Escape>", lambda _e: on_cancel())

        entry.focus_set()
        dialog.focus_force()
        dialog.after(50, entry.focus_set)
        dialog.wait_window()
    finally:
        try:
            root.destroy()
        except Exception:
            pass
        resume_automation_timer()

    return result


def prompt_batch_info(ncm_options: list[str]) -> dict[str, str] | None:
    """Prompt unico para CT-e, NF1/NF2 e NCM com validacoes basicas.

    Retorna None se o usuario cancelar.
    """
    pause_automation_timer()
    root = tk.Tk()
    root.withdraw()
    result: dict[str, str] = {}

    try:
        dialog = tk.Toplevel(root)
        dialog.title("Dados para Averbação")
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)
        dialog.geometry("600x560")

        font_label  = ("Segoe UI", 11)
        font_entry  = ("Segoe UI", 11)
        font_radio  = ("Segoe UI", 11)
        font_button = ("Segoe UI", 11)

        cte_var       = tk.StringVar()
        nf1_var       = tk.StringVar()
        nf2_var       = tk.StringVar()
        ncm_var       = tk.StringVar()
        ncm_other_var = tk.StringVar()

        frame = tk.Frame(dialog, padx=16, pady=14)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Número do CT-e (obrigatório):", font=font_label).grid(row=0, column=0, sticky="w")
        cte_entry = tk.Entry(frame, textvariable=cte_var, width=34, font=font_entry)
        cte_entry.grid(row=1, column=0, sticky="ew", pady=(2, 8))

        tk.Label(frame, text="NF1 (opcional):", font=font_label).grid(row=2, column=0, sticky="w")
        tk.Entry(frame, textvariable=nf1_var, width=34, font=font_entry).grid(row=3, column=0, sticky="ew", pady=(2, 8))

        tk.Label(frame, text="NF2 (opcional):", font=font_label).grid(row=4, column=0, sticky="w")
        tk.Entry(frame, textvariable=nf2_var, width=34, font=font_entry).grid(row=5, column=0, sticky="ew", pady=(2, 10))

        tk.Label(frame, text="Selecione o NCM:", font=font_label).grid(row=6, column=0, sticky="w")
        ncm_frame = tk.Frame(frame)
        ncm_frame.grid(row=7, column=0, sticky="w", pady=(2, 6))

        ncm_values: list[str] = []

        def select_radio_value(value: str) -> None:
            ncm_var.set(value)

        def on_radio_key(event, value: str) -> None:
            select_radio_value(value)

        for idx, option in enumerate(ncm_options):
            rb = tk.Radiobutton(ncm_frame, text=option, value=option, variable=ncm_var, takefocus=True, font=font_radio)
            rb.grid(row=idx, column=0, sticky="w")
            rb.bind("<Return>", lambda e, v=option: on_radio_key(e, v))
            rb.bind("<space>",  lambda e, v=option: on_radio_key(e, v))
            ncm_values.append(option)

        rb_other = tk.Radiobutton(ncm_frame, text="Outro:", value="__outro__", variable=ncm_var, takefocus=True, font=font_radio)
        rb_other.grid(row=len(ncm_options), column=0, sticky="w")
        rb_other.bind("<Return>", lambda e, v="__outro__": on_radio_key(e, v))
        rb_other.bind("<space>",  lambda e, v="__outro__": on_radio_key(e, v))
        ncm_values.append("__outro__")
        tk.Entry(ncm_frame, textvariable=ncm_other_var, width=22, takefocus=True, font=font_entry).grid(
            row=len(ncm_options), column=1, sticky="w", padx=(6, 0)
        )

        button_frame = tk.Frame(frame)
        button_frame.grid(row=8, column=0, sticky="e", pady=(8, 0))

        def on_ok() -> None:
            ncm_choice = ncm_var.get().strip()
            if ncm_choice == "__outro__":
                ncm_choice = ncm_other_var.get().strip()

            if not ncm_choice:
                messagebox.showwarning("NCM obrigatório", "Selecione um NCM ou informe um código em \"Outro\".")
                return

            result["cte"] = cte_var.get().strip()
            result["nf1"] = nf1_var.get().strip()
            result["nf2"] = nf2_var.get().strip()
            result["ncm"] = ncm_choice
            dialog.destroy()

        def on_cancel() -> None:
            result.clear()
            dialog.destroy()

        ok_button     = tk.Button(button_frame, text="OK",       command=on_ok,     width=10, font=font_button, takefocus=True)
        ok_button.pack(side="right", padx=(6, 0))
        cancel_button = tk.Button(button_frame, text="Cancelar", command=on_cancel, width=10, font=font_button, takefocus=True)
        cancel_button.pack(side="right")
        ok_button.bind(    "<Return>", lambda e: ok_button.invoke())
        cancel_button.bind("<Return>", lambda e: cancel_button.invoke())

        def select_focused_radio(event=None) -> None:
            widget = dialog.focus_get()
            if isinstance(widget, tk.Radiobutton):
                ncm_var.set(widget.cget("value"))

        def move_radio(delta: int) -> None:
            current = ncm_var.get()
            if current not in ncm_values:
                if ncm_values:
                    select_radio_value(ncm_values[0])
                return
            idx = ncm_values.index(current)
            next_idx = max(0, min(len(ncm_values) - 1, idx + delta))
            select_radio_value(ncm_values[next_idx])

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        dialog.bind("<Return>", select_focused_radio)
        dialog.bind("<space>",  select_focused_radio)
        dialog.bind("<Up>",     lambda e: move_radio(-1))
        dialog.bind("<Down>",   lambda e: move_radio(1))
        dialog.grab_set()
        if ncm_options:
            ncm_var.set(ncm_options[0])
        dialog.after(150, lambda: dialog.lift())
        dialog.after(200, lambda: dialog.attributes("-topmost", True))
        dialog.after(250, lambda: dialog.focus_force())
        dialog.after(300, lambda: cte_entry.focus_set())
        root.wait_window(dialog)
    finally:
        root.destroy()
        resume_automation_timer()

    if not result:
        return None
    return result
