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
import subprocess

pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Importar constantes centralizadas
from constants import (
    TAB_DELAY, CTRL_F_DELAY, SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG, SLEEP_ONE,
    LOG_TIMESTAMP_FORMAT, LOG_MESSAGE_FORMAT, MIN_PASTE_LENGTH, WRITE_INTERVAL
)
from constants import MUTEX_NAME

# Log para esta sessão
SESSION_TS = time.strftime(LOG_TIMESTAMP_FORMAT, time.localtime())
LOG_FILE = LOG_DIR / f"extension_filler_{SESSION_TS}.log"

# Constantes específicas do CIOT
CIOT_END_DELAY = 0.5  # Delay após pressionar End para página carregar


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


def navigate_to_page_bottom(delay: float = CIOT_END_DELAY) -> None:
    """Navega até o fim da página usando tecla End (resiliente e determinístico).
    
    Args:
        delay: Delay após chegar ao fim para a página carregar (segundos)
    """
    ui_print("Navegando até fim da página...", style="step")
    log("Pressionando End para ir ao fim da página")
    
    # Usar End key para ir diretamente ao fim da página
    pyautogui.press("end")
    time.sleep(delay)
    
    log("Chegado ao fim da página usando End key")
    ui_print("Fim da página atingido", style="success")



def find_and_click_ciot_toggle() -> bool:
    """Localiza e ativa o toggle de CIOT usando navegação por Tab.
    
    Fluxo simples e determinístico:
    1. Ctrl+F para buscar "INFORMAR DADOS DO CIOT"
    2. Esc para fechar Find
    3. Tab (1x) para mover para o toggle
    4. Space para ativar o toggle
    
    Returns:
        True se ativou com sucesso, False caso contrário
    """
    ui_print("Procurando e ativando toggle do CIOT...", style="step")
    log("Iniciando busca e ativação do toggle CIOT")
    
    try:
        # Passo 1: Abrir Find (Ctrl+F)
        ui_print("  • Abrindo busca (Ctrl+F)...", style="step")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(CTRL_F_DELAY)
        log("Ctrl+F pressionado")
        
        # Passo 2: Digitar termo de busca
        ui_print("  • Buscando 'INFORMAR DADOS DO CIOT'...", style="step")
        search_term = "INFORMAR DADOS DO CIOT"
        pyperclip.copy(search_term)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(SLEEP_SHORT)
        log(f"Termo buscado: '{search_term}'")
        
        # Aguardar resultado da busca
        time.sleep(SLEEP_MEDIUM)
        
        # Passo 3: Fechar Find
        ui_print("  • Fechando Find (Esc)...", style="step")
        pyautogui.press("escape")
        time.sleep(SLEEP_SHORT)
        log("Find fechado")
        
        # Passo 4: Um Tab para o toggle
        ui_print("  • Navegando para toggle (Tab)...", style="step")
        press_tab(1, TAB_DELAY)
        log("Tab pressionado (1x)")
        
        # Passo 5: Espaço para ativar
        ui_print("  • Ativando toggle (Space)...", style="step")
        pyautogui.press("space")
        time.sleep(SLEEP_MEDIUM)
        log("Space pressionado - toggle ativado")
        
        ui_print("Toggle do CIOT ativado com sucesso!", style="success")
        return True
        
    except Exception as e:
        log(f"Erro ao ativar toggle CIOT: {e}")
        ui_print(f"Erro: {e}", style="error")
        return False


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


def focus_browser_window(preferred_titles: list[str] | None = None) -> bool:
    """Tenta trazer a janela do navegador (onde o formulário está) para primeiro plano.

    Estratégia:
    1. Usa `pygetwindow` (se disponível) para buscar janelas cujo título contenha
       qualquer substring em `preferred_titles`.
    2. Se não houver `pygetwindow` ou nada for encontrado, faz um único Alt+Tab
       como fallback para minimizar a chance da console ficar em primeiro plano.

    Retorna True se parece ter focado uma janela, False caso contrário.
    """
    titles = preferred_titles or ["Emissor MDF-e", "MDF-e", "InvoiSys", "Emissor"]
    try:
        # 1) Primeiro, tentar pygetwindow se disponível (mais simples/limpo)
        try:
            import pygetwindow as gw
        except Exception:
            gw = None

        if gw:
            for sub in titles:
                try:
                    wins = gw.getWindowsWithTitle(sub)
                except Exception:
                    wins = []
                if wins:
                    win = wins[0]
                    try:
                        if win.isMinimized:
                            win.restore()
                    except Exception:
                        pass
                    try:
                        win.activate()
                        time.sleep(SLEEP_SHORT)
                        log(f"Focado navegador via pygetwindow: {win.title}")
                        return True
                    except Exception:
                        continue

        user32 = ctypes.windll.user32

        # 2) Tentar Win+1 (ativa o primeiro app fixado na barra; costuma ser navegador)
        ui_print("Tentando Win+1 para focar navegador...", style="step")
        pyautogui.hotkey("winleft", "1")
        time.sleep(SLEEP_SHORT)
        # checar se agora o foreground parece ser navegador
        try:
            fg = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(fg)
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(fg, buf, length + 1)
            title = (buf.value or "").lower()
            if any(k.lower() in title for k in titles) or any(k in title for k in ("chrome", "edge", "invoisys")):
                log("Navegador focado via Win+1")
                return True
        except Exception:
            pass

        # 3) Fallback Alt+Tab único para tentar tirar foco da console
        ui_print("Tentando Alt+Tab como fallback...", style="step")
        pyautogui.keyDown("alt")
        pyautogui.press("tab")
        pyautogui.keyUp("alt")
        time.sleep(SLEEP_SHORT)
        try:
            fg = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(fg)
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(fg, buf, length + 1)
            title = (buf.value or "").lower()
            if any(k.lower() in title for k in titles) or any(k in title for k in ("chrome", "edge", "invoisys")):
                log("Navegador focado via Alt+Tab")
                return True
        except Exception:
            pass

        # 4) EnumWindows: procurar por processo do navegador (msedge/chrome) e forçar foreground
        try:
            psapi = ctypes.windll.psapi
            kernel32 = ctypes.windll.kernel32

            found_hwnd = None

            @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
            def enum_proc(hwnd, lparam):
                if not user32.IsWindowVisible(hwnd):
                    return True
                length = user32.GetWindowTextLengthW(hwnd)
                if length == 0:
                    return True
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = (buf.value or "").lower()
                # obter nome do processo
                pid = ctypes.c_uint32()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if not pid.value:
                    return True
                process = kernel32.OpenProcess(0x0410, False, pid.value)
                if not process:
                    return True
                try:
                    name_buf = ctypes.create_unicode_buffer(260)
                    if psapi.GetModuleBaseNameW(process, None, name_buf, 260) == 0:
                        return True
                    pname = name_buf.value.lower()
                finally:
                    kernel32.CloseHandle(process)

                if any(k in pname for k in ("msedge.exe", "chrome.exe")) or any(k.lower() in title for k in titles) or any(k in title for k in ("chrome", "edge", "invoisys")):
                    nonlocal found_hwnd
                    found_hwnd = hwnd
                    return False
                return True

            user32.EnumWindows(enum_proc, 0)

            if found_hwnd:
                # Tentar forçar para foreground usando AttachThreadInput
                try:
                    SW_RESTORE = 9
                    if user32.IsIconic(found_hwnd):
                        user32.ShowWindow(found_hwnd, SW_RESTORE)
                    user32.SetForegroundWindow(found_hwnd)
                    time.sleep(SLEEP_SHORT)
                    log(f"Forçado foreground para janela do navegador (hwnd={found_hwnd})")
                    return True
                except Exception as e:
                    log(f"Falha ao forçar foreground: {e}")

        except Exception as e:
            log(f"EnumWindows fallback falhou: {e}")

        # Se todos os métodos falharem, retornar False
        log("Não foi possível focar o navegador")
        return False
    except Exception as e:
        log(f"Falha ao focar navegador: {e}")
        return False


def main() -> None:
    """Fluxo principal do framework de extensões de automação."""
    try:
        # Limpar tela
        os.system('cls' if os.name == 'nt' else 'clear')
        ui_print("AUTOMAÇÃO DO CIOT", style="header")
        
        log("═" * 60)
        log("Iniciando automação CIOT")
        log("═" * 60)
        
        # ==============================================================
        # FASE 1: NAVEGAR E ATIVAR TOGGLE (FUNDAÇÃO)
        # ==============================================================
        
        ui_print("\n[FASE 1] Navegando e ativando toggle...", style="header")
        log("FASE 1: Navegação até fim e ativação do toggle CIOT")
        
        # Passo 0: Garantir foco no navegador antes de enviar teclas
        focused = focus_browser_window()
        log(f"focus_browser_window returned: {focused}")
        # Pequena espera para garantir que o navegador processe o foco
        time.sleep(SLEEP_SHORT)

        # Passo 1: Navegar até o fim da página
        navigate_to_page_bottom()
        
        # Passo 2: Encontrar e ativar o toggle com Tab+Space
        toggle_success = find_and_click_ciot_toggle()
        
        if not toggle_success:
            log("Falha ao ativar toggle CIOT - interrompendo")
            ui_print("Não foi possível ativar o toggle. Verifique a página.", style="error")
            return
        
        time.sleep(SLEEP_MEDIUM)
        
        # ==============================================================
        # FASE 2: PREENCHER CAMPOS (PRÓXIMA)
        # ==============================================================
        
        ui_print("\n[FASE 2] Coletando informações CIOT...", style="header")
        log("FASE 2: Solicitar e preencher campos CIOT")
        
        # Solicitar informação do usuário
        log("Exibindo prompt para informar CIOT")
        field_value = prompt_field_value(
            field_name="CIOT",
            field_label="número do CIOT (Conhecimento de Transporte Intermodal Operacional)"
        )
        
        if not field_value:
            log("Usuário cancelou a operação na FASE 2")
            ui_print("Operação cancelada pelo usuário", style="warning")
            # Relançar o script principal para retornar ao fluxo original
            try:
                main_script = BASE_DIR / "modular_mdfe.py"
                if main_script.exists():
                    cmd = [str(sys.executable), str(main_script), "--skip-ciot"]
                    # Pequena espera para garantir que a instância anterior finalize
                    time.sleep(1.0)
                    if os.name == "nt":
                        subprocess.Popen(cmd, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NEW_CONSOLE)
                    else:
                        subprocess.Popen(cmd, cwd=str(BASE_DIR))
                    log("Relançado modular_mdfe.py após cancelamento do CIOT")
            except Exception as e:
                log(f"Falha ao relançar modular_mdfe.py: {e}")
            return
        
        ui_print(f"CIOT informado: {field_value}", style="success")
        log(f"CIOT capturado: {field_value}")
        
        # ==============================================================
        # FASE 3: SUBMETER (FINAL)
        # ==============================================================
        
        ui_print("\n[FASE 3] Pronto para submeter...", style="header")
        log("FASE 3: Preparado para submissão (aguardando próxima implementação)")
        
        # Exibir resumo
        GREEN = "\033[92m"
        CYAN = "\033[96m"
        YELLOW = "\033[93m"
        BOLD = "\033[1m"
        RESET = "\033[0m"
        
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{GREEN}  ✓ CIOT PRONTO PARA PREENCHIMENTO!{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
        print(f"{BOLD}Dados Capturados:{RESET}\n")
        print(f"  {YELLOW}CIOT:{RESET}  {field_value}")
        print(f"\n{BOLD}Próximos passos:{RESET}")
        print(f"  1. Preencher campo CIOT na página")
        print(f"  2. Solicitar CPF/CNPJ do Responsável")
        print(f"  3. Clicar botão 'Emitir'")
        print(f"\n{CYAN}{'═' * 60}{RESET}\n")
        
        log("Automação CIOT concluída com sucesso")
        # Ao concluir (ou mesmo se o usuário cancelar), relançar o script
        # principal para que o terminal principal seja exibido novamente.
        try:
            main_script = BASE_DIR / "modular_mdfe.py"
            if main_script.exists():
                cmd = [str(sys.executable), str(main_script), "--skip-ciot"]
                # Esperar a liberação do mutex (se presente) antes de relançar.
                def _is_mutex_present(name: str) -> bool:
                    if os.name != "nt":
                        return False
                    try:
                        kernel32 = ctypes.windll.kernel32
                        OpenMutexW = kernel32.OpenMutexW
                        OpenMutexW.argtypes = (ctypes.c_uint32, ctypes.c_int, ctypes.c_wchar_p)
                        OpenMutexW.restype = ctypes.c_void_p
                        SYNCHRONIZE = 0x00100000 if False else 0x0001
                        handle = OpenMutexW(SYNCHRONIZE, False, name)
                        if not handle:
                            return False
                        kernel32.CloseHandle(handle)
                        return True
                    except Exception:
                        return False

                max_wait = 10.0
                waited = 0.0
                interval = 0.5
                while _is_mutex_present(MUTEX_NAME) and waited < max_wait:
                    log(f"Mutex {MUTEX_NAME} ainda presente; aguardando {interval}s")
                    time.sleep(interval)
                    waited += interval

                if _is_mutex_present(MUTEX_NAME):
                    log(f"Mutex {MUTEX_NAME} ainda presente após {max_wait}s; relançamento pode falhar")

                # Pequena espera final antes de iniciar o novo processo
                time.sleep(0.3)
                if os.name == "nt":
                    subprocess.Popen(cmd, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NEW_CONSOLE)
                else:
                    subprocess.Popen(cmd, cwd=str(BASE_DIR))
                log("Relançado modular_mdfe.py após conclusão do CIOT")
        except Exception as e:
            log(f"Falha ao relançar modular_mdfe.py: {e}")
        
    except Exception as e:
        log(f"Erro geral na automação CIOT: {e}")
        ui_print(f"Erro geral: {e}", style="error")
        import traceback
        log(traceback.format_exc())
        
        # Restaurar terminal
        restore_console_popup()


if __name__ == "__main__":
    main()
