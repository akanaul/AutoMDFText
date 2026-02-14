"""Automacao MDF-e com selecao de perfil, prompts gui e preenchimento via teclado.

Inclui failsafe (F8), pausa (F9) e validacoes de tela para reduzir erros de
preenchimento em formularios do navegador.
"""

import argparse
import ctypes
import os
import re
import tkinter as tk
from tkinter import messagebox
import time
import threading
from pathlib import Path

import pyautogui
import pyperclip

pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "logs"
CONFIG_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
# Log por sessao (timestamp) para facilitar debug
SESSION_TS = time.strftime("%Y%m%d_%H%M%S", time.localtime())
LOG_FILE = LOG_DIR / f"automation_{SESSION_TS}.log"

TAB_DELAY = 0.10
TAB_DELAY_LONG = 0.25
CTRL_F_DELAY = 0.2
DROPDOWN_SETTLE_DELAY = 0.4
SLEEP_SHORT = 0.15
SLEEP_MEDIUM = 0.25
SLEEP_LONG = 0.45
SLEEP_LONGER = 0.7
SLEEP_ONE = 1.0
SLEEP_ONE_HALF = 1.5
# Handle for single-instance mutex to keep it alive during process lifetime
_SINGLETON_MUTEX_HANDLE = None

# Variáveis de tracking de tempo
_automation_start_time = 0.0
_automation_time_paused = 0.0
_pause_start_time = 0.0
_failsafe_listener = None
_pause_requested = False
_pause_active = False
_pause_lock = threading.Lock()
_last_write_value = None
_last_write_verify = False


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def log(msg: str) -> None:
    """Registra mensagem apenas no arquivo de log, sem imprimir no console."""
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Não interromper fluxo por falha de log em disco
        pass


def start_automation_session(selected: str, profile_path: Path) -> float:
    """Cria um novo log de sessao e reinicia os contadores de tempo.

    Retorna o instante monotonic para calculo do tempo total.
    """
    global LOG_FILE, SESSION_TS, _automation_start_time, _automation_time_paused

    SESSION_TS = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    LOG_FILE = LOG_DIR / f"automation_{SESSION_TS}.log"

    real_start_time = time.monotonic()
    _automation_start_time = time.monotonic()
    _automation_time_paused = 0.0

    log("Iniciando automação (main)")
    log(f"[DEBUG] real_start_time={real_start_time}")
    log(f"Perfil selecionado: {selected}")
    log(f"Perfil carregado com sucesso de: {profile_path}")
    log(f"[DEBUG] _automation_start_time iniciado após escolha do perfil: {_automation_start_time}")

    return real_start_time


def start_failsafe_f8() -> None:
    """Inicia listener global para F8 (encerrar) e F9 (pausar).

    Em Windows, ignora eventos injetados para aceitar apenas teclado fisico.
    """
    global _failsafe_listener
    if _failsafe_listener is not None:
        return
    try:
        from pynput import keyboard
    except Exception as exc:
        log(f"Aviso: pynput nao disponivel; failsafe F8 desativado ({exc})")
        return

    injected = {"value": False}

    def win32_event_filter(_msg, data):
        flags = getattr(data, "flags", 0)
        injected["value"] = bool(flags & 0x10)
        return True

    def show_failsafe_alert() -> None:
        if os.name != "nt":
            return
        try:
            ctypes.windll.user32.MessageBoxW(
                0,
                "A automacao foi encerrada pelo botao de seguranca (F8).",
                "Automacao encerrada",
                0x00000040 | 0x00040000 | 0x00010000,
            )
        except Exception:
            pass

    def on_press(key) -> None:
        if injected["value"]:
            return
        if key == keyboard.Key.f8:
            log("Failsafe F8 acionado. Encerrando automacao.")
            show_failsafe_alert()
            os._exit(1)
        if key == keyboard.Key.f9:
            request_pause()

    listener_kwargs = {"on_press": on_press}
    if os.name == "nt":
        listener_kwargs["win32_event_filter"] = win32_event_filter
    _failsafe_listener = keyboard.Listener(**listener_kwargs)
    _failsafe_listener.start()


def stop_failsafe_f8() -> None:
    """Finaliza o listener de failsafe por F8, se ativo."""
    global _failsafe_listener
    if _failsafe_listener is None:
        return
    try:
        _failsafe_listener.stop()
    except Exception:
        pass
    _failsafe_listener = None


def ui_print(msg: str, style: str = "info") -> None:
    """Imprime mensagem formatada no console estilo GUI."""
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


def pause_automation_timer() -> None:
    """Pausa o contador de tempo de automação (usado durante prompts)."""
    global _pause_start_time
    _pause_start_time = time.monotonic()


def resume_automation_timer() -> None:
    """Resume o contador de tempo de automação após prompt."""
    global _automation_time_paused, _pause_start_time
    if _pause_start_time > 0:
        _automation_time_paused += time.monotonic() - _pause_start_time
        _pause_start_time = 0.0

    log(f"[DEBUG PAUSE] Resuming timer. Total paused so far: {_automation_time_paused}s")


def request_pause() -> None:
    """Marca a automacao para pausar na proxima verificacao segura."""
    global _pause_requested
    with _pause_lock:
        if _pause_requested or _pause_active:
            return
        _pause_requested = True
    log("Pausa solicitada (F9). Aguardando ponto seguro para pausar...")


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
            SW_SHOW = 5
            HWND_TOPMOST = -1
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_NOACTIVATE = 0x0010
            user32.ShowWindow(hwnd, SW_SHOW)
            user32.SetWindowPos(
                hwnd,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            )
        except Exception:
            pass

    font_label = ("Segoe UI", 11)
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

    resume_button = tk.Button(button_frame, text="Retomar", command=on_resume, width=10, font=font_button)
    resume_button.pack(side="right", padx=(6, 0))
    cancel_button = tk.Button(button_frame, text="Cancelar automacao", command=on_cancel, width=18, font=font_button)
    cancel_button.pack(side="right")

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    dialog.bind("<Return>", lambda _e: on_resume())
    dialog.bind("<Escape>", lambda _e: on_cancel())

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


def check_pause() -> None:
    """Pausa em ponto seguro e valida o ultimo campo digitado antes de bloquear."""
    global _pause_requested, _pause_active
    if not _pause_requested or _pause_active:
        return

    _pause_active = True
    pause_automation_timer()
    log("Automacao pausada pelo usuario.")
    _verify_last_write_before_pause()
    try:
        decision = show_pause_dialog()
    finally:
        _pause_active = False

    if decision == "cancel":
        log("Automacao cancelada pelo usuario durante a pausa.")
        raise SystemExit(1)

    with _pause_lock:
        _pause_requested = False
    resume_automation_timer()
    log("Automacao retomada pelo usuario.")


def pause_point() -> None:
    """Ponto seguro para pausar entre etapas (fora de sequencias tab/enter)."""
    check_pause()


def _verify_last_write_before_pause() -> None:
    """Revalida o ultimo smart_write para evitar campo vazio antes da pausa."""
    global _last_write_value, _last_write_verify
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
        _last_write_value = None
        _last_write_verify = False


def format_duration(seconds: float) -> str:
    """Formata duracao em segundos para o formato MM:SS ou apenas SS."""
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    if minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"


def skip_tabs(count: int, log_msg: str = "") -> None:
    """Pula N campos (tabs) com log opcional."""
    if log_msg:
        log(log_msg)
    press_tab(count=count, delay=TAB_DELAY)


def press_tab(count: int = 1, delay: float = TAB_DELAY) -> None:
    """Pressiona Tab com delay consistente entre navegacoes."""
    for _ in range(count):
        pyautogui.press("tab")
        time.sleep(delay)


def paste_text(
    text: str,
    verify: bool = True,
    retries: int = 2,
    delay: float = 0.15,
    restore_clipboard: bool = True,
) -> None:
    """Cola texto via clipboard e valida o conteúdo quando possível.

    Quando verify=True, tenta ler o campo com Ctrl+A/C e compara com o valor.
    """
    try:
        previous = pyperclip.paste()
    except Exception:
        previous = None

    def normalize(value: str) -> str:
        return re.sub(r"\s+", " ", value or "").strip()

    def attempt_paste() -> bool:
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(delay)

        if not verify:
            return True

        try:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.05)
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.05)
            captured = pyperclip.paste()
        except Exception:
            return False

        return normalize(captured) == normalize(text)

    try:
        ok = False
        for _ in range(max(1, retries + 1)):
            if attempt_paste():
                ok = True
                break
            time.sleep(delay)

        if verify and not ok:
            log("Aviso: verificação de colagem falhou; seguindo adiante")
    finally:
        if restore_clipboard and previous is not None:
            try:
                pyperclip.copy(previous)
            except Exception:
                pass


def smart_write(
    value: str,
    interval: float = 0.10,
    min_paste_len: int = 4,
    verify: bool = True,
) -> None:
    """Escolhe entre digitar e colar, com verificação opcional.

    Desativa a verificacao para CPF/CNPJ (11/14 digitos) por formatacao automatica.
    """
    global _last_write_value, _last_write_verify
    pause_point()
    if value is None:
        return
    text = str(value)
    if not text:
        return

    use_paste = len(text) >= min_paste_len or any(ch in text for ch in " /-_:.\t")
    verify_effective = verify
    if text.isdigit() and len(text) in (11, 14):
        # CPF/CNPJ normalmente são formatados automaticamente pelo formulário
        verify_effective = False
    _last_write_value = text
    _last_write_verify = verify_effective
    if use_paste:
        paste_text(text, verify=verify_effective)
    else:
        pyautogui.write(text, interval=interval)
    pause_point()


def parse_profile(path: Path) -> dict[str, dict[str, str]]:
    sections: dict[str, dict[str, str]] = {}
    current_section = "GENERAL"
    sections[current_section] = {}

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and "]" in line:
            current_section = line[1:line.index("]")].strip().upper()
            sections.setdefault(current_section, {})
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            sections[current_section][key.strip().lower()] = value.strip()

    return sections


class ConfigProfile:
    def __init__(self, path: Path):
        self.path = path
        self._data: dict[str, dict[str, str]] = {}
        self._mtime = 0.0
        self.reload()

    def reload(self) -> None:
        self._data = parse_profile(self.path)
        try:
            self._mtime = self.path.stat().st_mtime
        except OSError:
            self._mtime = 0.0

    def ensure_current(self) -> None:
        try:
            current_mtime = self.path.stat().st_mtime
        except OSError:
            current_mtime = 0.0
        if current_mtime != self._mtime:
            self.reload()

    def get_value(self, section: str, key: str, default: str = "") -> str:
        self.ensure_current()
        return self._data.get(section.upper(), {}).get(key.lower(), default)


def list_profiles() -> list[str]:
    return sorted(p.name for p in CONFIG_DIR.glob("*.txt"))


def hide_console_window() -> None:
    if os.name != "nt":
        return
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if not hwnd:
        return
    # Apenas oculta o console para evitar encerramento do processo durante o debug
    ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE


def restore_console_popup() -> None:
    """Restaura o terminal, trazendo-o ao topo como um popup curto."""
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

        # Mostrar e trazer para frente
        user32.ShowWindow(hwnd, SW_SHOW)
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
        user32.BringWindowToTop(hwnd)

        # Tornar topmost brevemente para comportamento de popup
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        time.sleep(0.15)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
    except Exception:
        pass


def play_low_beep() -> None:
    """Emite um beep de baixa frequência ao final da automação."""
    try:
        if os.name == "nt":
            # Frequência baixa (~400Hz), duração 180ms
            import winsound
            winsound.Beep(400, 180)
        else:
            # Fallback para bell character em outros sistemas
            print("\a", end="", flush=True)
    except Exception:
        pass


def ensure_single_instance(name: str = "Global\\AutoMDFText_Mutex", on_duplicate: str = "warn") -> None:
    """Impede execução duplicada usando um Mutex nomeado do Windows.
    Se já existir outra instância:
    - on_duplicate == 'warn': exibe alerta e encerra este processo
    - on_duplicate == 'kill': encerra este processo (equivale a matar o duplicado)
    """
    if os.name != "nt":
        return
    kernel32 = ctypes.windll.kernel32
    # CreateMutexW(lpMutexAttributes, bInitialOwner, lpName)
    handle = kernel32.CreateMutexW(None, False, name)
    last_error = kernel32.GetLastError()
    if last_error == 183:  # ERROR_ALREADY_EXISTS
        try:
            if on_duplicate == "warn":
                focused_alert("Já existe uma instância em execução. O processo será encerrado.")
        except Exception:
            pass
        raise SystemExit(0)
    else:
        # Manter handle vivo para não liberar o mutex
        global _SINGLETON_MUTEX_HANDLE
        _SINGLETON_MUTEX_HANDLE = handle


def choose_profile(interactive_list: list[str]) -> str:
    log("Iniciando seleção de perfil")
    if not interactive_list:
        log("Nenhum script disponível; diretório scripts/ está vazio.")
        try:
            focused_alert(
                "Nenhum script encontrado em scripts/.\n\n"
                "Adicione um arquivo .txt de configuração e tente novamente.",
                title="Nenhum script encontrado"
            )
        except Exception:
            pass
        raise SystemExit(1)
    
    # Cores ANSI para destacar o menu
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BOLD = "\033[1m"
    RESET = "\033[0m"
    
    selected_profile = None
    max_attempts = 100  # Proteção contra loops infinitos por erro de lógica
    attempts = 0
    
    # Loop até que uma seleção válida seja feita
    while not selected_profile and attempts < max_attempts:
        attempts += 1
        
        try:
            # Revalidar lista a cada iteração (caso arquivos sejam adicionados/removidos)
            current_list = list_profiles()
            if not current_list:
                log("Lista de scripts ficou vazia durante seleção; sem perfil padrão disponível.")
                raise SystemExit(1)
            
            # Atualizar lista se mudou
            if current_list != interactive_list:
                interactive_list = current_list
                log(f"Lista de scripts atualizada: {len(interactive_list)} scripts disponíveis.")
            
            # Menu destacado com cores e separadores
            print(f"\n{CYAN}{'=' * 60}{RESET}")
            print(f"{BOLD}{GREEN}  SELEÇÃO DE SCRIPT - AUTOMAÇÃO MDF-e{RESET}")
            print(f"{CYAN}{'=' * 60}{RESET}\n")
            
            print(f"{YELLOW}  [0]{RESET} Voltar ao menu anterior\n")
            
            for idx, name in enumerate(interactive_list, start=1):
                print(f"{YELLOW}  [{idx}]{RESET} {name}")
            
            print(f"\n{CYAN}{'─' * 60}{RESET}")
            print(f"{BOLD}Digite o número do script desejado (ou 0 para voltar):{RESET}")
            print(f"{CYAN}{'─' * 60}{RESET}\n")
            
            choice = input(f"{BOLD}Opção: {RESET}").strip()
            
            # Validar entrada
            if not choice:
                print(f"\n{RED}✗ Erro: Você deve selecionar um script!{RESET}")
                log("Entrada vazia; solicitando nova seleção.")
                time.sleep(SLEEP_ONE_HALF)
                continue
            
            # Verificar opção "0" para voltar
            if choice == "0":
                print(f"\n{GREEN}✓ Retornando ao menu anterior...{RESET}\n")
                log("Usuário escolheu retornar ao menu anterior")
                raise SystemExit(99)
            
            # Tentar converter para número
            if choice.isdigit():
                try:
                    index = int(choice) - 1
                    if 0 <= index < len(interactive_list):
                        selected_profile = interactive_list[index]
                        log(f"Script selecionado por índice {choice}: {selected_profile}")
                    else:
                        print(f"\n{RED}✗ Erro: Número inválido! Escolha entre 1 e {len(interactive_list)}.{RESET}")
                        log(f"Opção fora do intervalo: {choice}")
                        time.sleep(SLEEP_ONE_HALF)
                except (ValueError, IndexError) as e:
                    print(f"\n{RED}✗ Erro ao processar número: {str(e)}{RESET}")
                    log(f"Erro ao processar índice {choice}: {e}")
                    time.sleep(SLEEP_ONE_HALF)
            # Aceitar nome exato do arquivo (case-insensitive para maior flexibilidade)
            elif choice.lower() in [s.lower() for s in interactive_list]:
                # Encontrar o nome com case correto
                for script in interactive_list:
                    if script.lower() == choice.lower():
                        selected_profile = script
                        log(f"Script selecionado por nome: {selected_profile}")
                        break
            else:
                print(f"\n{RED}✗ Erro: Opção inválida! Digite um número válido.{RESET}")
                log(f"Opção inválida: {choice}")
                time.sleep(SLEEP_ONE_HALF)
                
        except KeyboardInterrupt:
            print(f"\n\n{RED}✗ Seleção cancelada pelo usuário.{RESET}")
            log("Seleção interrompida por Ctrl+C")
            raise SystemExit(0)
        except Exception as e:
            print(f"\n{RED}✗ Erro inesperado: {str(e)}{RESET}")
            log(f"ERRO durante seleção de perfil: {e}")
            time.sleep(SLEEP_ONE_HALF)
            # Continuar o loop para tentar novamente
            continue
    
    # Verificação de segurança
    if not selected_profile:
        log(f"Número máximo de tentativas atingido ({max_attempts}); nenhum perfil selecionado.")
        raise SystemExit(1)
    
    print(f"\n{GREEN}✓ Script selecionado: {BOLD}{selected_profile}{RESET}\n")
    print(f"{CYAN}{'=' * 60}{RESET}\n")
    
    # Fechar/ocultar o terminal após seleção (com fallback em caso de WM_CLOSE falhar)
    hide_console_window()
    
    return selected_profile


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

        font_label = ("Segoe UI", 11)
        font_entry = ("Segoe UI", 11)
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
        resume_automation_timer()

    return result


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

        font_label = ("Segoe UI", 11)
        font_entry = ("Segoe UI", 11)
        font_radio = ("Segoe UI", 11)
        font_button = ("Segoe UI", 11)

        cte_var = tk.StringVar()
        nf1_var = tk.StringVar()
        nf2_var = tk.StringVar()
        ncm_var = tk.StringVar()
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
            rb.bind("<space>", lambda e, v=option: on_radio_key(e, v))
            ncm_values.append(option)

        rb_other = tk.Radiobutton(ncm_frame, text="Outro:", value="__outro__", variable=ncm_var, takefocus=True, font=font_radio)
        rb_other.grid(row=len(ncm_options), column=0, sticky="w")
        rb_other.bind("<Return>", lambda e, v="__outro__": on_radio_key(e, v))
        rb_other.bind("<space>", lambda e, v="__outro__": on_radio_key(e, v))
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

        ok_button = tk.Button(button_frame, text="OK", command=on_ok, width=10, font=font_button, takefocus=True)
        ok_button.pack(side="right", padx=(6, 0))
        cancel_button = tk.Button(button_frame, text="Cancelar", command=on_cancel, width=10, font=font_button, takefocus=True)
        cancel_button.pack(side="right")
        ok_button.bind("<Return>", lambda e: ok_button.invoke())
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
        dialog.bind("<space>", select_focused_radio)
        dialog.bind("<Up>", lambda e: move_radio(-1))
        dialog.bind("<Down>", lambda e: move_radio(1))
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


def ensure_caps_off() -> None:
    VK_CAPITAL = 0x14
    caps_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)
    if caps_state & 1:
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)


def _get_foreground_title() -> str:
    """Retorna o título da janela em foco (minimiza uso de Win+1 desnecessário)."""
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value or ""


def _get_foreground_class() -> str:
    """Retorna a classe da janela em foco para identificar navegadores com mais confiabilidade."""
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ""
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value or ""


def _get_window_process_name(hwnd: int) -> str:
    """Retorna o nome do processo dono da janela (ex: msedge.exe)."""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    psapi = ctypes.windll.psapi

    pid = ctypes.c_uint32()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ""

    process = kernel32.OpenProcess(0x0410, False, pid.value)
    if not process:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(260)
        if psapi.GetModuleBaseNameW(process, None, buffer, 260) == 0:
            return ""
        return buffer.value.lower()
    finally:
        kernel32.CloseHandle(process)


def _is_cloaked_window(hwnd: int) -> bool:
    """Retorna True se a janela estiver cloaked (UWP em background)."""
    try:
        dwmapi = ctypes.windll.dwmapi
    except Exception:
        return False

    DWMWA_CLOAKED = 14
    cloaked = ctypes.c_int(0)
    if dwmapi.DwmGetWindowAttribute(
        hwnd, DWMWA_CLOAKED, ctypes.byref(cloaked), ctypes.sizeof(cloaked)
    ) != 0:
        return False
    return cloaked.value != 0


def _is_top_level_app_window(hwnd: int) -> bool:
    """Filtra janelas utilitarias/owned que nao representam uma janela real."""
    user32 = ctypes.windll.user32
    GW_OWNER = 4
    GWL_EXSTYLE = -20
    WS_EX_TOOLWINDOW = 0x00000080

    if user32.GetWindow(hwnd, GW_OWNER):
        return False
    ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    if ex_style & WS_EX_TOOLWINDOW:
        return False
    return True


def _is_standard_window(hwnd: int) -> bool:
    """Valida se a janela tem estilo tipico de app (com titulo/borda)."""
    user32 = ctypes.windll.user32
    GWL_STYLE = -16
    WS_OVERLAPPEDWINDOW = 0x00CF0000

    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    return (style & WS_OVERLAPPEDWINDOW) == WS_OVERLAPPEDWINDOW


def _is_browser_window(title: str, cls: str, process_name: str = "") -> bool:
    """Identifica se a janela pertence ao navegador pela combinacao de titulo/classe/processo."""
    title_hits = ("chrome", "edge", "navegador", "invoisys", "google chrome", "microsoft edge")
    process_hits = ("msedge.exe", "chrome.exe")
    # ApplicationFrameWindow aparece em varios apps UWP; evite class-only match.
    class_hits = ("chrome_widgetwin_1",)

    if process_name in process_hits:
        return True
    return any(k in title for k in title_hits) or (cls in class_hits)


def _find_browser_windows() -> list[int]:
    """Encontra janelas de navegador visiveis (Chrome/Edge) pela classe/titulo."""
    if os.name != "nt":
        return []

    user32 = ctypes.windll.user32
    windows: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_proc(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        if _is_cloaked_window(hwnd):
            return True
        if not _is_top_level_app_window(hwnd):
            return True
        if not _is_standard_window(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        title = (buffer.value or "").lower()
        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buffer, 256)
        cls = (class_buffer.value or "").lower()

        process_name = _get_window_process_name(int(hwnd))
        if _is_browser_window(title, cls, process_name):
            windows.append(int(hwnd))
        return True

    user32.EnumWindows(enum_proc, 0)
    return windows


def focus_browser_if_needed() -> None:
    """Só pressiona Win+1 se o navegador não estiver em foco, evitando minimizar."""
    browser_windows = _find_browser_windows()
    if len(browser_windows) > 1:
        focused_alert(
            "Foram detectadas multiplas janelas do navegador abertas. "
            "Isso pode atrapalhar a automacao. A primeira janela sera utilizada.",
            title="Aviso: Multiplas janelas do navegador"
        )
        log("Multiplas janelas detectadas; usando Win+1 para ir para a primeira.")
        pyautogui.hotkey("winleft", "1")
        time.sleep(0.8)

    title = _get_foreground_title().lower()
    cls = _get_foreground_class().lower()
    process_name = _get_window_process_name(ctypes.windll.user32.GetForegroundWindow())
    if _is_browser_window(title, cls, process_name):
        log("Navegador já em foco; Win+1 ignorado para evitar minimizar.")
        return

    log("Navegador fora de foco; tentando Win+1.")
    pyautogui.hotkey("winleft", "1")
    time.sleep(SLEEP_ONE)

    # Pós-checagem: se ainda não estiver em foco, tentar fallback suave
    title2 = _get_foreground_title().lower()
    cls2 = _get_foreground_class().lower()
    process_name2 = _get_window_process_name(ctypes.windll.user32.GetForegroundWindow())
    if _is_browser_window(title2, cls2, process_name2):
        log("Navegador em foco após Win+1.")
        return
    log("Win+1 não focou o navegador; evitando minimizar e mantendo estado.")


def upload_latest_xml() -> None:
    """Seleciona o arquivo mais recente em Downloads e confirma o upload."""
    time.sleep(SLEEP_MEDIUM)
    downloads_path = Path.home() / "Downloads"
    list_of_files = list(downloads_path.glob("*"))
    if not list_of_files:
        focused_alert("A pasta Downloads está vazia!")
        return
    latest_file = max(list_of_files, key=os.path.getctime)
    smart_write(str(latest_file), interval=0.12)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")




def wait_for_form(target_text: str, tempo_maximo: float = 15.0, intervalo: float = 1.0, copy_attempts: int = 2) -> str:
    """Aguarda o formulario abrir detectando texto via clipboard.

    Retorna o conteudo copiado quando o alvo e encontrado.
    """
    short_sleep = 0.12

    inicio = time.monotonic()
    ultimo_conteudo = ""
    target_norm = re.sub(r"\s+", " ", target_text).strip().lower()

    log(f"Aguardando o formulário abrir... alvo='{target_text}', timeout={tempo_maximo}s")
    attempt = 0

    while time.monotonic() - inicio < tempo_maximo:
        attempt += 1
        try:
            log(f"Tentativa {attempt}: limpando clipboard e copiando")
            try:
                pyperclip.copy("")
            except Exception as e:
                log(f"Aviso: não foi possível limpar clipboard: {e}")

            for _ in range(copy_attempts):
                pyautogui.hotkey("ctrl", "a")
                time.sleep(short_sleep)
                pyautogui.hotkey("ctrl", "c")
                time.sleep(short_sleep)

            try:
                conteudo = pyperclip.paste() or ""
            except Exception as e:
                conteudo = ""
                log(f"Aviso: erro ao ler clipboard: {e}")

            ultimo_conteudo = conteudo
            conteudo_norm = re.sub(r"\s+", " ", conteudo).strip().lower()

            log(f"Tentativa {attempt}: comprimento clipboard = {len(conteudo)}")

            if target_norm in conteudo_norm:
                log("Formulário detectado!")
                log(f"Conteúdo capturado (preview): {conteudo[:200]}")
                log("Continuando com o restante da automação...")
                time.sleep(0.8)
                return conteudo

            log(f"Não encontrado. Aguardando {intervalo}s antes da próxima tentativa.")
            time.sleep(intervalo)

        except Exception as e:
            log(f"Erro interno durante a tentativa: {e}")
            time.sleep(intervalo)

    log(f"Formulário não foi detectado dentro de {tempo_maximo} segundos. Encerrando o processo.")
    if ultimo_conteudo:
        log("Último conteúdo capturado (preview):")
        print(ultimo_conteudo[:400])
    raise SystemExit(1)


def _normalize_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


EDGE_SEARCHBAR_HEIGHT = 70
EDGE_CLICK_OFFSET = 200


def _click_below_edge_searchbar(offset: int = EDGE_CLICK_OFFSET) -> None:
    """Clica ~50px abaixo da barra de pesquisa do Edge para focar o conteúdo."""
    if os.name != "nt":
        width, height = pyautogui.size()
        pyautogui.click(width // 2, height // 2)
        return

    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        width, height = pyautogui.size()
        pyautogui.click(width // 2, height // 2)
        return

    rect = ctypes.wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        width, height = pyautogui.size()
        pyautogui.click(width // 2, height // 2)
        return

    left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
    x = (left + right) // 2
    y = top + EDGE_SEARCHBAR_HEIGHT + offset
    if y >= bottom - 10:
        y = top + max(10, (bottom - top) // 3)

    pyautogui.click(x, y)


def _focus_page_for_copy() -> None:
    """Garante foco no corpo da página antes de copiar (evita pegar barra de endereço)."""
    try:
        pyautogui.press("esc")
        time.sleep(0.05)
        _click_below_edge_searchbar()
        time.sleep(0.10)
    except Exception:
        pass


def verify_cte_on_page(numero_cte: str, tempo_maximo: float = 6.0, intervalo: float = 1.0) -> None:
    """Copia o conteúdo da página e confirma a presença do CT-e informado."""
    if not numero_cte:
        return

    raw_cte = str(numero_cte).strip()
    digits_cte = _normalize_digits(raw_cte)
    pattern = None
    if digits_cte:
        pattern = re.compile(rf"(?<!\d){re.escape(digits_cte)}(?!\d)")

    inicio = time.monotonic()
    while time.monotonic() - inicio < tempo_maximo:
        try:
            _focus_page_for_copy()
            for _ in range(2):
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.12)
                pyautogui.hotkey("ctrl", "c")
                time.sleep(SLEEP_SHORT)

            conteudo = pyperclip.paste() or ""
            if raw_cte and raw_cte in conteudo:
                log(f"CT-e {numero_cte} encontrado na página (match direto).")
                return
            if digits_cte:
                if pattern and pattern.search(conteudo):
                    log(f"CT-e {numero_cte} encontrado na página (match numérico).")
                    return
                conteudo_digits = _normalize_digits(conteudo)
                if digits_cte in conteudo_digits:
                    log(f"CT-e {numero_cte} encontrado na página (match normalizado).")
                    return
        except Exception as exc:
            log(f"Aviso: falha ao verificar CT-e na página ({exc})")

        time.sleep(intervalo)

    focused_alert(
        "O número do CT-e informado não foi encontrado na página.\n\n"
        "Verifique se o número digitado é o mesmo que aparece na tela e tente novamente.",
        title="CT-e não encontrado"
    )
    raise SystemExit(1)


def navigate_to_mdfe() -> None:
    """Navega para o formulario MDF-e seguindo o fluxo legado ajustado."""
    # IR PARA 3ª PAGINA - MDFE
    pyautogui.hotkey("ctrl", "3")
    time.sleep(SLEEP_ONE)

    # ABRIR DADOS DO MDF-E
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_MEDIUM)
    log("Procurando por 'EMITIR NOTA'")
    smart_write("EMITIR NOTA", interval=0.10)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONGER)
    
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_MEDIUM)
    log("Procurando por 'MDF-E'")
    smart_write("MDF-E", interval=0.10)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONGER)


def fill_mdfe(profile: ConfigProfile, codigo_ncm: str) -> None:
    """Preenche o formulario MDF-e usando valores do perfil e NCM selecionado.

    Espera que o formulario MDF-e esteja aberto e com foco.
    """
    time.sleep(SLEEP_ONE)
    log("Iniciando preenchimento MDF-e: PRESTADOR DE SERVIÇO, EMITENTE, UF, MUNICÍPIO")
    
    # PRESTADOR DE SERVIÇO
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_MEDIUM)
    smart_write("SELECIONE...", interval=0.20)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    prestador = profile.get_value('mdfe', 'prestador_tipo')
    smart_write(prestador, interval=0.1)
    time.sleep(SLEEP_LONG)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)

    # EMITENTE
    emitente = profile.get_value("mdfe", "emitente_codigo")
    smart_write(emitente, interval=0.10)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONGER)
    skip_tabs(7)
    pyautogui.press("space")
    time.sleep(SLEEP_MEDIUM)

    # UF CARREGAMENTO E DESCARREGAMENTO
    uf_car = profile.get_value("mdfe", "uf_carregamento")
    smart_write(uf_car, interval=0.20)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("space")
    time.sleep(SLEEP_MEDIUM)
    uf_desc = profile.get_value("mdfe", "uf_descarga")
    smart_write(uf_desc, interval=0.20)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONGER)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    municipio = profile.get_value("mdfe", "municipio_carregamento").upper()
    smart_write(municipio, interval=0.15)
    time.sleep(SLEEP_MEDIUM)

    for _ in range(4):
        pyautogui.press("down")
        time.sleep(0.1)
    for _ in range(3):
        pyautogui.press("up")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    log(f"MDF-e: Prestador={prestador}, Emitente={emitente}, UF_Car={uf_car}, UF_Desc={uf_desc}, Municipio={municipio}")
    
    # UPLOAD DO ARQUIVO XML
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_ONE_HALF)
    log("Carregando arquivo XML...")
    upload_latest_xml()
    time.sleep(SLEEP_LONG)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_ONE_HALF)

    # UNIDADE DE MEDIDA, TIPO CARGA E DESCRIÇÃO
    skip_tabs(5)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    unidade = profile.get_value("mdfe", "unidade_medida")
    smart_write(unidade, interval=0.1)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    carga = profile.get_value("mdfe", "carga_tipo")
    smart_write(carga, interval=0.1)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    log(f"MDF-e: Unidade={unidade}, Tipo_Carga={carga}")

    # DESCRIÇÃO DO PRODUTO
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    descricao = profile.get_value("mdfe", "codigo_produto_descricao")
    log(f"Preenchendo DESCRIÇÃO PRODUTO: {descricao}")
    smart_write(descricao, interval=0.1)
    time.sleep(SLEEP_SHORT)

    # CÓDIGO NCM (já selecionado e passado como parâmetro)
    skip_tabs(2)
    
    codigo_ncm_upper = codigo_ncm.upper()
    smart_write(codigo_ncm_upper, interval=0.1)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    # CEP ORIGEM E DESTINO
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    cep_orig = profile.get_value("mdfe", "cep_origem")
    smart_write(cep_orig, interval=0.12)
    time.sleep(SLEEP_MEDIUM)

    skip_tabs(3)
    cep_dest = profile.get_value("mdfe", "cep_destino")
    smart_write(cep_dest, interval=0.12)
    time.sleep(SLEEP_MEDIUM)
    
    log(f"MDF-e concluído: NCM={codigo_ncm_upper}, CEP_Orig={cep_orig}, CEP_Dest={cep_dest}")


def fill_modal_rodo(profile: ConfigProfile) -> None:
    """Preenche os dados do Modal Rodoviario conforme o perfil ativo."""
    log("Iniciando preenchimento Modal Rodoviário")
    time.sleep(SLEEP_MEDIUM)
    
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_SHORT)
    smart_write("modal rodo", interval=0.10)
    time.sleep(SLEEP_SHORT)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONG)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONG)

    # RNTRC, CONTRATANTE, CNPJ
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_MEDIUM)
    smart_write("RNTRC", interval=0.10)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    rntrc = profile.get_value("modal_rodoviario", "rntrc")
    smart_write(rntrc, interval=0.10)
    time.sleep(SLEEP_SHORT)
    
    skip_tabs(6)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    contratante = profile.get_value("modal_rodoviario", "contratante_nome")
    smart_write(contratante, interval=0.20)
    time.sleep(SLEEP_SHORT)

    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    cnpj_cont = profile.get_value("modal_rodoviario", "contratante_cnpj")
    smart_write(cnpj_cont, interval=0.12)
    time.sleep(SLEEP_SHORT)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(1.25)
    
    log(f"Modal Rodoviário: RNTRC={rntrc}, Contratante={contratante}, CNPJ={cnpj_cont}")


def fill_additional_info(profile: ConfigProfile) -> None:
    """Preenche Informacoes Adicionais (seguradora, frete, banco e parcelas)."""
    log("Iniciando preenchimento Informações Adicionais")

    def write_additional(value: str, interval: float = 0.10, **_kwargs) -> None:
        smart_write(value, interval=interval, verify=False)
    
    # ABRIR SEÇÃO OPCIONAIS
    pyautogui.hotkey("ctrl", "f")
    write_additional("OPCIONAIS", interval=0.10)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    # ADICIONAIS - CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_LONG)
    write_additional("ADICIONAIS", interval=0.10)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_LONG)

    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_LONG)
    write_additional("CONTRIBUINTE", interval=0.10)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    skip_tabs(3)
    cnpj_contrib = profile.get_value("informacoes_adicionais", "contribuinte_cnpj")
    write_additional(cnpj_contrib, interval=0.10)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    write_additional("CONTRA", interval=0.10)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    
    # Seguradora e dados relacionados
    pyautogui.press("tab")
    cnpj_cont2 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    write_additional(cnpj_cont2, interval=0.12)
    pyautogui.press("tab")
    seguradora = profile.get_value("informacoes_adicionais", "seguradora_nome")
    write_additional(seguradora, interval=0.10)
    pyautogui.press("tab")
    cnpj_seg = profile.get_value("informacoes_adicionais", "seguradora_cnpj")
    write_additional(cnpj_seg, interval=0.12)
    pyautogui.press("tab")
    apolice = profile.get_value("informacoes_adicionais", "numero_apolice")
    write_additional(apolice, interval=0.10)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    skip_tabs(4)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)

    # Contratante 2 e Frete
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    contratante2 = profile.get_value("modal_rodoviario", "contratante_nome")
    write_additional(contratante2, interval=0.12)
    time.sleep(SLEEP_SHORT)

    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    cnpj_cont3 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    write_additional(cnpj_cont3, interval=0.12)
    time.sleep(SLEEP_SHORT)

    skip_tabs(2)
    frete_val = profile.get_value("informacoes_adicionais", "frete_valor")
    write_additional(frete_val, interval=0.12)
    time.sleep(SLEEP_SHORT)

    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)
    forma_pag = profile.get_value("informacoes_adicionais", "forma_pagamento")
    write_additional(forma_pag, interval=0.12)
    time.sleep(SLEEP_SHORT)

    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    # Banco e Agência
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    numero_banco = profile.get_value("informacoes_adicionais", "numero_banco")
    write_additional(numero_banco, interval=0.12)
    time.sleep(SLEEP_SHORT)

    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    agencia = profile.get_value("informacoes_adicionais", "agencia")
    write_additional(agencia, interval=0.12)
    time.sleep(SLEEP_SHORT)

    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    # Seção FRETE
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")

    pyautogui.hotkey("ctrl", "f")
    time.sleep(CTRL_F_DELAY)
    write_additional("SELECIONE...", interval=0.10)
    time.sleep(CTRL_F_DELAY)
    pyautogui.press("esc")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(CTRL_F_DELAY)
    write_additional("FRETE", interval=0.10)
    time.sleep(CTRL_F_DELAY)
    pyautogui.press("enter")
    time.sleep(DROPDOWN_SETTLE_DELAY)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    frete_val2 = profile.get_value("informacoes_adicionais", "frete_valor")
    write_additional(frete_val2, interval=0.12)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    frete_tipo = profile.get_value("informacoes_adicionais", "frete_tipo")
    write_additional(frete_tipo, interval=0.12)
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_SHORT)

    # Seção de Parcelas
    pyautogui.hotkey("ctrl", "f")
    write_additional("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)

    skip_tabs(5)
    numero_parcelas = profile.get_value("informacoes_adicionais", "numero_parcelas")
    if not numero_parcelas:
        log("Perfil sem 'informacoes_adicionais.numero_parcelas' obrigatório")
        focused_alert(
            "O perfil está faltando a chave obrigatória: informacoes_adicionais.numero_parcelas",
            title="Perfil inválido"
        )
        raise SystemExit(1)
    write_additional(numero_parcelas, interval=0.10)
    time.sleep(0.15)

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(SLEEP_SHORT)

    skip_tabs(7)
    for _ in range(2):
        pyautogui.press("space")
        time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    frete_val3 = profile.get_value("informacoes_adicionais", "frete_valor")
    write_additional(frete_val3, interval=0.10, verify=False)
    time.sleep(0.15)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_ONE)
    
    log(f"Informações Adicionais: Seguradora={seguradora}, Frete={frete_val}, Banco={numero_banco}, Parcelas={numero_parcelas}")


def perform_averbacao(numero_cte: str = "", numero_dt: str = "", nf_concat: str = "") -> None:
    """Executa a averbacao, extrai o numero e preenche a area de contribuinte.

    Assume que a aba de averbacao e a aba do sistema estao abertas.
    """
    def write_averbacao(value: str, interval: float = 0.10) -> None:
        smart_write(value, interval=interval, verify=False)

    # Abrir site/aba de averbação e enviar XML
    pyautogui.hotkey("ctrl", "4")
    time.sleep(SLEEP_LONG)

    ##Pequeno hotfix para GAP relacionado a envio de XMLs, verificar alternativas
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)

    for search in ("OK", "XML", "ENVIAR"):
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.4)
        write_averbacao(search, interval=0.1)
        time.sleep(0.4)
        pyautogui.press("esc")
        time.sleep(SLEEP_SHORT)
        pyautogui.press("enter")
        time.sleep(SLEEP_MEDIUM)

    upload_latest_xml()
    time.sleep(2.5)

    # Extrair número de averbação e copiar apenas os dígitos
    pyautogui.hotkey("ctrl", "a")
    time.sleep(SLEEP_SHORT)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(SLEEP_SHORT)
    texto = pyperclip.paste()
    numero_averbacao = ""
    match = re.search(r"Número de Averbação:\s*([\d]+)", texto)
    if match:
        numero_averbacao = match.group(1)
        print("Número de Averbação copiado:", numero_averbacao)
    else:
        print("Número de Averbação não encontrado")

    time.sleep(SLEEP_LONG)
    pyautogui.hotkey("alt", "tab")
    time.sleep(SLEEP_LONGER)

    # Preencher detalhes na outra aba
    pyautogui.hotkey("ctrl", "home")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.4)
    pyautogui.write("DETALHES", interval=0.1)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(SLEEP_MEDIUM)
    if numero_averbacao:
        pyperclip.copy(numero_averbacao)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)

    # Preencher DT/CT-e/NF na área de CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(SLEEP_MEDIUM)
    write_averbacao("CONTRIBUINTE", interval=0.15)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("esc")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)

    # DT sempre vem do prompt (numero_dt) capturado no início
    write_averbacao("DT: ", interval=0.1)
    time.sleep(SLEEP_MEDIUM)
    if numero_dt:
        pyperclip.copy(numero_dt)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(SLEEP_LONG)

    # CT-e: usar o valor informado pelo usuário
    write_averbacao(" CTE: ", interval=0.1)
    time.sleep(SLEEP_LONG)
    if numero_cte:
        log(f"Usando CT-e informado pelo usuário: {numero_cte}")
        pyperclip.copy(numero_cte)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(SLEEP_LONG)
    else:
        log("CT-e não informado; campo ficará vazio")
        time.sleep(SLEEP_LONG)

    # Finalizar com rótulo NF sempre; deixa vazio se não informado
    if nf_concat:
        write_averbacao(f" NF: {nf_concat}", interval=0.1)
    else:
        write_averbacao(" NF: ", interval=0.1)
    time.sleep(SLEEP_LONG)


def main() -> None:
    """Fluxo principal da automacao MDF-e."""
    global _automation_start_time, _automation_time_paused

    prev_failsafe = pyautogui.FAILSAFE
    pyautogui.FAILSAFE = False
    try:
        # Iniciar contadores de tempo após seleção do script
        real_start_time = 0.0
        _automation_time_paused = 0.0
        
        # Sleep inicial para aguardar inicialização
        time.sleep(SLEEP_LONGER)
        
        # Limpar tela e mostrar header
        os.system('cls' if os.name == 'nt' else 'clear')
        ui_print("AUTOMAÇÃO MDF-e", style="header")
        
        # Impedir execução duplicada
        ensure_single_instance()
        
        # Selecionar perfil de configuracao
        parser = argparse.ArgumentParser()
        parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
        args = parser.parse_args()

        selected = args.profile
        if not selected:
            log("Exibindo menu de seleção de perfil")
            selected = choose_profile(list_profiles())
        ui_print(f"Perfil: {selected}", style="success")
        profile_path = CONFIG_DIR / selected
        if not profile_path.exists():
            log(f"Perfil {profile_path} não encontrado.")
            try:
                focused_alert(
                    f"O arquivo de perfil não foi encontrado:\n{profile_path}\n\nCorrija o nome do script e tente novamente.",
                    title="Perfil ausente"
                )
            except Exception:
                pass
            raise SystemExit(1)

        profile = ConfigProfile(profile_path)
        real_start_time = start_automation_session(selected, profile_path)
        start_failsafe_f8()
        pause_point()

        ui_print("Iniciando preenchimento...", style="step")
        
        # Preparar navegador e validar pagina inicial
        # Abrir/focar navegador sem minimizar (usa Win+1 só se não estiver em foco)
        log("Focando navegador (evitando minimizar)")
        focus_browser_if_needed()
        pause_point()
        
        # GAP - Pressionar ESC 2x
        log("Enviando ESC x2")
        for _ in range(2):
            pyautogui.press("esc")
            time.sleep(SLEEP_MEDIUM)

        # Recarregar aba 3
        log("Recarregando aba 3 (Ctrl+3, F5)")
        pyautogui.hotkey("ctrl", "3")
        time.sleep(SLEEP_LONG)
        pyautogui.press("f5")
        time.sleep(SLEEP_ONE)
        pause_point()

        # Voltar para aba 1 (uma vez, como no legado)
        log("Voltando para aba 1")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)
        # Tab para garantir foco correto
        pyautogui.press("esc")
        time.sleep(SLEEP_LONG)
        pause_point()

        pause_point()

        # Coleta de informacoes obrigatorias
        # Prompt para DT ANTES de buscar o campo
        prompt_text = profile.get_value("general", "dt_prompt_text")
        if not prompt_text:
            log("Perfil sem 'general.dt_prompt_text' obrigatório")
            focused_alert(
                "O perfil está faltando a chave obrigatória: general.dt_prompt_text",
                title="Perfil inválido"
            )
            raise SystemExit(1)
        log("Exibindo prompt de DT")
        numero_dt = prompt_dt_blocking(text=prompt_text, title="DT")
        if not numero_dt:
            focused_alert("Nenhum código DT informado. O script foi pausado.")
            return
        log(f"DT informado: {numero_dt}")
        log("Focando a primeira aba do navegador (Ctrl+1) apos o DT")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)
        pause_point()

        # Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        pyautogui.press("backspace")
        time.sleep(0.15)
        smart_write("DO DT", interval=0.12)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("esc")
        time.sleep(SLEEP_LONG)
        pyautogui.press("tab")
        time.sleep(SLEEP_LONG)
        
        # Usar o DT armazenado previamente
        log(f"Preenchendo campo DT com valor armazenado: {numero_dt}")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        paste_text(numero_dt.upper(), verify=True)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("enter")
        time.sleep(SLEEP_LONG)

        # Aviso inicial para baixar o CT-e, antes do prompt unificado
        log("Exibindo aviso para baixar o CT-e antes de seguir")
        focused_alert(
            text=(
                "Antes de prosseguir, faça o download do XML do CT-e correspondente \n"
                "à DT e mantenha-o salvo. Em seguida clique em OK para continuar."
            ),
            title="Aviso: Baixe o CT-e"
        )
        pause_point()

        ncm_primary = profile.get_value("mdfe", "ncm_primary")
        ncm_secondary = profile.get_value("mdfe", "ncm_secondary")
        ncm_tertiary = profile.get_value("mdfe", "ncm_tertiary")
        if not (ncm_primary and ncm_secondary and ncm_tertiary):
            log("Perfil sem códigos NCM obrigatórios (mdfe.ncm_primary/secondary/tertiary)")
            focused_alert(
                "O perfil está faltando códigos NCM obrigatórios:\n"
                "mdfe.ncm_primary, mdfe.ncm_secondary, mdfe.ncm_tertiary",
                title="Perfil inválido"
            )
            raise SystemExit(1)

        log("Exibindo prompt unificado para CT-e, NFs e NCM")
        batch = prompt_batch_info([ncm_primary, ncm_secondary, ncm_tertiary])
        if not batch:
            focused_alert("Nenhuma informação foi informada. O script foi pausado.")
            raise SystemExit(1)

        log("Focando a primeira aba do navegador (Ctrl+1) apos o prompt de dados")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)

        numero_cte = batch.get("cte", "")
        nf1 = batch.get("nf1", "")
        nf2 = batch.get("nf2", "")
        codigo_ncm = batch.get("ncm", "")

        log(f"CT-e informado: '{numero_cte}'")
        if not numero_cte:
            log("Nenhum número de CT-e informado; encerrando automação conforme solicitado")
            focused_alert(
                "ERRO: Nenhum número de CT-e foi informado.\n\n"
                "A automação será encerrada.",
                title="CT-e obrigatório"
            )
            raise SystemExit(1)

        # Verificar se o CT-e informado aparece na página após inserir a DT
        pyautogui.hotkey("ctrl", "1")
        verify_cte_on_page(numero_cte)
        pause_point()

        # Concatenar apenas se ambas informadas; caso contrário usar a que foi preenchida
        if nf1 and nf2:
            nf_concat = f"{nf1}/{nf2}"
        elif nf1:
            nf_concat = nf1
        elif nf2:
            nf_concat = nf2
        else:
            nf_concat = ""
        log(f"NF coletadas: '{nf1}' e '{nf2}' => '{nf_concat}'")
        log(f"NCM selecionado e armazenado: {codigo_ncm}")
        ui_print(f"NCM selecionado: {codigo_ncm}", style="success")
        time.sleep(SLEEP_LONG)
        pyautogui.press("esc")
        time.sleep(0.15)
        ensure_caps_off()
        pause_point()

        ui_print("Preenchendo formulário MDF-e...", style="step")
        
        # Preencher formularios principais
        # Navegar para MDF-e e detectar formulário (lógica e tempos do legado)
        navigate_to_mdfe()
        wait_for_form("Emissor MDF-e", tempo_maximo=15.0, intervalo=3.0, copy_attempts=3)
        pause_point()
        
        # Preencher formulário (passando código NCM já selecionado)
        log("Iniciando preenchimento dos dados MDF-e")
        fill_mdfe(profile, codigo_ncm)
        log("Dados MDF-e preenchidos com sucesso")
        ui_print("Dados MDF-e preenchidos", style="success")
        pause_point()
        
        log("Iniciando preenchimento do modal rodoviário")
        fill_modal_rodo(profile)
        log("Modal rodoviário preenchido com sucesso")
        ui_print("Modal rodoviário preenchido", style="success")
        pause_point()
        
        log("Iniciando preenchimento de informações adicionais")
        fill_additional_info(profile)
        log("Informações adicionais preenchidas com sucesso")
        ui_print("Informações adicionais preenchidas", style="success")
        pause_point()
        
        log("Iniciando processamento de averbação")
        ui_print("Processando averbação...", style="step")
        perform_averbacao(numero_cte, numero_dt, nf_concat)
        log("Averbação processada com sucesso")
        pause_point()
        
        # Encerramento com resumo
        # Ao finalizar com sucesso, calcular tempos
        real_end_time = time.monotonic()
        automation_end_time = time.monotonic()
        
        real_duration = real_end_time - real_start_time
        automation_duration = (automation_end_time - _automation_start_time) - _automation_time_paused
        
        log(f"[DEBUG TIMING] real_end_time={real_end_time}, automation_end_time={automation_end_time}")
        log(f"[DEBUG TIMING] _automation_start_time={_automation_start_time}, _automation_time_paused={_automation_time_paused}")
        log(f"[DEBUG TIMING] real_duration calc: {real_end_time} - {real_start_time} = {real_duration}")
        log(f"[DEBUG TIMING] automation_duration calc: ({automation_end_time} - {_automation_start_time}) - {_automation_time_paused} = {automation_duration}")
        
        # Restaurar o terminal como popup e emitir beep baixo
        restore_console_popup()
        play_low_beep()
        
        # Exibir resumo no terminal estilo GUI
        GREEN = "\033[92m"
        CYAN = "\033[96m"
        YELLOW = "\033[93m"
        BOLD = "\033[1m"
        RESET = "\033[0m"
        
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{GREEN}  ✓ AUTOMAÇÃO CONCLUÍDA COM SUCESSO!{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
        print(f"{BOLD}Resumo das Informações:{RESET}\n")
        print(f"  {YELLOW}DT:{RESET}      {numero_dt}")
        print(f"  {YELLOW}CT-e:{RESET}    {numero_cte if numero_cte else 'Não capturado'}")
        print(f"  {YELLOW}NCM:{RESET}     {codigo_ncm}")
        print(f"  {YELLOW}NF:{RESET}      {nf_concat if nf_concat else 'Não informado'}")
        print(f"\n{BOLD}Tempo de Execução:{RESET}\n")
        print(f"  {YELLOW}Tempo de Automação:{RESET}  {format_duration(automation_duration)} (apenas automação)")
        print(f"  {YELLOW}Tempo Real:{RESET}          {format_duration(real_duration)} (incluindo prompts)")
        print(f"\n{CYAN}{'─' * 60}{RESET}")
        print(f"{BOLD}Próximos passos:{RESET}")
        print(f"  • Preencha os dados do motorista")
        print(f"\n{CYAN}{'═' * 60}{RESET}\n")
        
        # Pausar 3 segundos para permitir leitura do resumo
        time.sleep(3)
        
        log(f"Automação finalizada com sucesso - Tempo automação: {format_duration(automation_duration)}, Tempo real: {format_duration(real_duration)}")
        log("═" * 60)
        log(f"Resumo final: DT={numero_dt}, CT-e={numero_cte if numero_cte else 'Não capturado'}, NCM={codigo_ncm}, NF={nf_concat if nf_concat else 'Não informado'}")
        log("═" * 60)
    except SystemExit as e:
        # Capturar saídas como exit code 99 (menu), 1 (erro), etc
        if e.code == 99:
            log("Programa finalizado - usuário retornou ao menu (código 99)")
        else:
            log(f"Programa finalizado com código de saída: {e.code}")
        raise
    except Exception as e:
        import traceback
        error_msg = f"ERRO FATAL: {e}\n{traceback.format_exc()}"
        log(error_msg)
        # Exibir alerta do erro
        try:
            focused_alert(
                f"Erro durante a automação:\n\n{str(e)}\n\nVer log para detalhes completos.",
                title="Erro na automação"
            )
        except Exception:
            pass
        # Restaurar terminal para que o usuário possa ver o erro
        restore_console_popup()
        raise
    finally:
        stop_failsafe_f8()
        pyautogui.FAILSAFE = prev_failsafe


if __name__ == "__main__":
    main()
