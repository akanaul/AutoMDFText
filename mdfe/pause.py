"""
Controle de pausa segura para AutoMDFText.

Responsabilidades:
- Exibir o diálogo de pausa topmost (show_pause_dialog)
- Verificar e executar a pausa no próximo ponto seguro (check_pause, pause_point)
- Revalidar o último campo digitado antes de bloquear (verify_last_write_before_pause)

Estado de pausa (_pause_requested, _pause_active, _pause_lock) vive em
mdfe.failsafe porque F9 (o gatilho) é parte do mesmo listener de failsafe.
Este módulo importa o estado de lá.

Nota de imports lazy:
    _verify_last_write_before_pause referencia paste_text e _normalize_text
    de mdfe.keyboard (criado na Parte 3). O import é feito dentro da função
    para evitar dependência circular em nível de módulo.
"""

import ctypes
import os
import time
import tkinter as tk


def show_pause_dialog() -> str:
    """Exibe um dialogo topmost de pausa sem roubar foco.

    Retorna "resume" para continuar ou "cancel" para encerrar a automacao.
    """
    root = tk.Tk()
    root.withdraw()

    result = {"value": "resume"}

    dialog = tk.Toplevel(root)
    dialog.title("Automacao pausada")
    dialog.attributes("-topmost", True)
    dialog.resizable(False, False)
    dialog.geometry("460x360")

    hwnd = dialog.winfo_id()

    def keep_visible() -> None:
        if os.name != "nt" or not dialog.winfo_exists():
            return
        try:
            user32 = ctypes.windll.user32
            SW_SHOW      = 5
            HWND_TOPMOST = -1
            SWP_NOMOVE   = 0x0002
            SWP_NOSIZE   = 0x0001
            SWP_NOACTIVATE = 0x0010
            user32.ShowWindow(hwnd, SW_SHOW)
            user32.SetWindowPos(
                hwnd,
                HWND_TOPMOST,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            )
        except Exception:
            pass

    font_label  = ("Segoe UI", 11)
    font_button = ("Segoe UI", 11)

    frame = tk.Frame(dialog, padx=16, pady=18)
    frame.pack(fill="both", expand=True)

    tk.Label(
        frame,
        text=(
            "A automacao esta pausada.\n\n"
            "Clique em Retomar para continuar ou em Cancelar para encerrar."
        ),
        font=font_label,
        justify="left",
        wraplength=380,
    ).pack(anchor="w")

    button_frame = tk.Frame(frame)
    button_frame.pack(fill="x", pady=(16, 0))

    def on_resume() -> None:
        result["value"] = "resume"
        dialog.destroy()

    def on_cancel() -> None:
        result["value"] = "cancel"
        dialog.destroy()

    resume_button = tk.Button(button_frame, text="Retomar",           command=on_resume, width=10, font=font_button)
    resume_button.pack(side="right", padx=(6, 0))
    cancel_button = tk.Button(button_frame, text="Cancelar automacao", command=on_cancel, width=18, font=font_button)
    cancel_button.pack(side="right")

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    dialog.bind("<Return>",  lambda _e: on_resume())
    dialog.bind("<Escape>",  lambda _e: on_cancel())

    def watchdog_visible() -> None:
        if not dialog.winfo_exists():
            return
        keep_visible()
        dialog.lift()
        dialog.grab_set()
        dialog.after(600, watchdog_visible)

    keep_visible()
    dialog.after(200, watchdog_visible)
    dialog.wait_window()
    try:
        root.destroy()
    except Exception:
        pass

    return result["value"]


def _verify_last_write_before_pause() -> None:
    """Revalida o ultimo smart_write para evitar campo vazio antes da pausa."""
    # Import lazy — mdfe.keyboard é criado na Parte 3
    from mdfe.keyboard import paste_text, _normalize_text, _last_write_value, _last_write_verify
    from mdfe.logger import log
    import pyautogui
    import pyperclip

    if not _last_write_verify or not _last_write_value:
        return
    try:
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.05)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.05)
        captured = pyperclip.paste() or ""
        if _normalize_text(captured) != _normalize_text(_last_write_value):
            log("Aviso: campo divergente antes da pausa; reaplicando valor.")
            paste_text(_last_write_value, verify=True, retries=1)
    except Exception as exc:
        log(f"Aviso: falha ao reverificar ultimo campo antes da pausa ({exc})")
    finally:
        # Limpa o estado após verificação
        from mdfe import keyboard as _kb
        _kb._last_write_value  = None
        _kb._last_write_verify = False


def check_pause() -> None:
    """Pausa em ponto seguro e valida o ultimo campo digitado antes de bloquear."""
    from mdfe.failsafe import _pause_lock, request_pause
    from mdfe.logger import log
    from mdfe.timing import pause_automation_timer, resume_automation_timer
    import mdfe.failsafe as _fs

    if not _fs._pause_requested or _fs._pause_active:
        return

    _fs._pause_active = True
    pause_automation_timer()
    log("Automacao pausada pelo usuario.")
    _verify_last_write_before_pause()
    try:
        decision = show_pause_dialog()
    finally:
        _fs._pause_active = False

    if decision == "cancel":
        log("Automacao cancelada pelo usuario durante a pausa.")
        raise SystemExit(1)

    with _pause_lock:
        _fs._pause_requested = False
    resume_automation_timer()
    log("Automacao retomada pelo usuario.")


def pause_point() -> None:
    """Ponto seguro para pausar entre etapas (fora de sequencias tab/enter)."""
    check_pause()
