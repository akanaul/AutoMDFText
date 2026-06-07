"""
Script desacoplado para extensões de automação MDF-e.

Este script é chamado após a finalização do modular_mdfe.py e pode ser executado
independentemente se necessário. Oferece framework genérico para adicionar novos campos.
"""

import ctypes
import os
import sys
import tkinter as tk
from tkinter import messagebox
import time
from pathlib import Path

import pyautogui
import pyperclip

pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Importar constantes centralizadas
from constants import (
    TAB_DELAY, CTRL_F_DELAY, SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG, SLEEP_ONE,
    LOG_TIMESTAMP_FORMAT, LOG_MESSAGE_FORMAT, MIN_PASTE_LENGTH, WRITE_INTERVAL
)

# Log para esta sessão
SESSION_TS = time.strftime(LOG_TIMESTAMP_FORMAT, time.localtime())
LOG_FILE = LOG_DIR / f"extension_filler_{SESSION_TS}.log"


def log(msg: str) -> None:
    """Registra mensagem no arquivo de log."""
    ts = time.strftime(LOG_MESSAGE_FORMAT, time.localtime())
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def ui_print(msg: str, style: str = "info") -> None:
    """Imprime mensagem formatada no console."""
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BLUE = "\033[94m"
    BOLD = "\033[1m"
    RESET = "\033[0m"
    
    if style == "success":
        print(f"{GREEN}✓{RESET} {msg}")
    elif style == "error":
        print(f"{RED}✗{RESET} {msg}")
    elif style == "warning":
        print(f"{YELLOW}⚠{RESET} {msg}")
    elif style == "step":
        print(f"{BLUE}▸{RESET} {msg}")
    elif style == "header":
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{msg}{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
    else:
        print(f"  {msg}")


def focused_alert(text: str = "", title: str = "", button: str = "OK"):
    """Exibe alerta com foco na tela."""
    try:
        result = pyautogui.alert(text=text, title=title, button=button)
    except Exception as e:
        log(f"Erro ao exibir alerta: {e}")
    return result


def prompt_field_value(field_name: str, field_label: str = "") -> str | None:
    """Prompt genérico para o usuário informar valor de um campo.
    
    Args:
        field_name: Nome do campo (ex: 'CIOT')
        field_label: Label descritivo opcional (ex: 'Conhecimento de Transporte Intermodal Operacional')
    
    Returns:
        Valor informado pelo usuário ou None se cancelado
    """
    root = tk.Tk()
    root.withdraw()
    
    result = None
    
    try:
        dialog = tk.Toplevel(root)
        dialog.title(f"Informar {field_name}")
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)
        dialog.geometry("460x200")
        dialog.lift()

        font_label = ("Segoe UI", 11)
        font_entry = ("Segoe UI", 11)
        font_button = ("Segoe UI", 11)

        frame = tk.Frame(dialog, padx=16, pady=14)
        frame.pack(fill="both", expand=True)

        label_text = f"Informe o {field_label or f'valor do {field_name}'}:"
        tk.Label(
            frame,
            text=label_text,
            font=font_label,
            wraplength=420,
            justify="left"
        ).pack(anchor="w")

        entry_var = tk.StringVar()
        entry = tk.Entry(frame, textvariable=entry_var, font=font_entry)
        entry.pack(fill="x", pady=(8, 12))

        button_frame = tk.Frame(frame)
        button_frame.pack(fill="x")

        def on_ok() -> None:
            nonlocal result
            value = entry_var.get().strip()
            if not value:
                messagebox.showwarning(
                    f"{field_name} obrigatório",
                    f"O {field_name} precisa ser digitado para continuar.",
                    parent=dialog,
                )
                entry.focus_set()
                return
            result = value
            dialog.destroy()

        def on_cancel() -> None:
            dialog.destroy()

        ok_button = tk.Button(button_frame, text="OK", command=on_ok, width=10, font=font_button)
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

    return result


def smart_write(
    value: str,
    interval: float = 0.10,
    min_paste_len: int = 4,
) -> None:
    """Escreve valor via clipboard ou digitação, escolhendo o melhor método."""
    if value is None:
        return
    
    text = str(value)
    if not text:
        return

    use_paste = len(text) >= min_paste_len or any(ch in text for ch in " /-_:.\t")
    
    if use_paste:
        try:
            pyperclip.copy(text)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(interval)
        except Exception as e:
            log(f"Erro ao colar texto: {e}")
            pyautogui.write(text, interval=interval)
    else:
        pyautogui.write(text, interval=interval)


def press_tab(count: int = 1, delay: float = TAB_DELAY) -> None:
    """Pressiona Tab N vezes com delay entre cada um."""
    for _ in range(count):
        pyautogui.press("tab")
        time.sleep(delay)


def restore_console_popup() -> None:
    """Restaura o terminal como popup."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        hwnd = kernel32.GetConsoleWindow()
        if not hwnd:
            return

        SW_SHOW = 5
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040

        user32.ShowWindow(hwnd, SW_SHOW)
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
        user32.BringWindowToTop(hwnd)

        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        time.sleep(0.15)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
    except Exception:
        pass


def main() -> None:
    """Fluxo principal do framework de extensões de automação."""
    try:
        # Limpar tela
        os.system('cls' if os.name == 'nt' else 'clear')
        ui_print("EXTENSÃO DE AUTOMAÇÃO", style="header")
        
        # Solicitar informação genérica ao usuário
        log("Exibindo prompt genérico para extensão")
        field_value = prompt_field_value(
            field_name="CIOT",
            field_label="número do CIOT (Conhecimento de Transporte Intermodal Operacional)"
        )
        
        if not field_value:
            log("Usuário cancelou a operação")
            ui_print("Operação cancelada pelo usuário", style="warning")
            return
        
        ui_print(f"Valor informado: {field_value}", style="success")
        log(f"Valor informado pelo usuário: {field_value}")
        
        # Aqui você pode chamar qualquer função de preenchimento específica
        # Por exemplo: fill_ciot(field_value) ou qualquer outra
        # Por enquanto, apenas mostramos o resumo
        
        # Exibir resumo
        GREEN = "\033[92m"
        CYAN = "\033[96m"
        YELLOW = "\033[93m"
        BOLD = "\033[1m"
        RESET = "\033[0m"
        
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{GREEN}  ✓ EXTENSÃO EXECUTADA COM SUCESSO!{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
        print(f"{BOLD}Valor Capturado:{RESET}\n")
        print(f"  {YELLOW}CIOT:{RESET}  {field_value}")
        print(f"\n{CYAN}{'═' * 60}{RESET}\n")
        
        # Restaurar terminal
        restore_console_popup()
        
        log("Extensão finalizada com sucesso")
        
    except Exception as e:
        import traceback
        error_msg = f"ERRO FATAL: {e}\n{traceback.format_exc()}"
        log(error_msg)
        
        try:
            focused_alert(
                f"Erro na extensão:\n\n{str(e)}\n\nVer log para detalhes completos.",
                title="Erro na execução"
            )
        except Exception:
            pass
        
        restore_console_popup()
        raise


if __name__ == "__main__":
    main()
