import argparse
import ctypes
import os
import re
import threading
import time
from pathlib import Path

import pyautogui
import pyperclip


pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "logs"
CONFIG_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
DEFAULT_PROFILE = "template_config.txt"
# Log por sessão (timestamp) para facilitar debug
SESSION_TS = time.strftime("%Y%m%d_%H%M%S", time.localtime())
LOG_FILE = LOG_DIR / f"automation_{SESSION_TS}.log"

# Handle for single-instance mutex to keep it alive during process lifetime
_SINGLETON_MUTEX_HANDLE = None

# Variáveis de tracking de tempo
_automation_start_time = 0.0
_automation_time_paused = 0.0
_pause_start_time = 0.0


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

def format_duration(seconds: float) -> str:
    """Formata duracao em segundos para o formato MM:SS ou apenas SS."""
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    if minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"


# ============================================================================
# FUNÇÕES HELPER PARA CONSOLIDAR AÇÕES COM LOGS
# ============================================================================

def find_and_fill(search_text: str, fill_value: str, log_msg: str = "", use_enter: bool = True) -> None:
    """Localiza campo, preenche com valor e registra. Consolida log + ação."""
    msg = log_msg or f"Preenchendo {search_text}: {fill_value}"
    log(msg)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write(search_text, interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(fill_value, interval=0.10)
    if use_enter:
        pyautogui.press("enter")
        time.sleep(0.3)


def find_text(search_text: str, log_msg: str = "") -> None:
    """Localiza texto na página com Ctrl+F."""
    msg = log_msg or f"Procurando por '{search_text}'"
    log(msg)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write(search_text, interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)


def fill_field(value: str, log_msg: str = "", interval: float = 0.10) -> None:
    """Preenche campo atual com valor."""
    if log_msg:
        log(log_msg)
    pyautogui.write(value, interval=interval)
    time.sleep(0.2)


def skip_tabs(count: int, log_msg: str = "") -> None:
    """Pula N campos (tabs) com log opcional."""
    if log_msg:
        log(log_msg)
    for _ in range(count):
        pyautogui.press("tab")
        time.sleep(0.25)


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
                time.sleep(1.5)
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
                        time.sleep(1.5)
                except (ValueError, IndexError) as e:
                    print(f"\n{RED}✗ Erro ao processar número: {str(e)}{RESET}")
                    log(f"Erro ao processar índice {choice}: {e}")
                    time.sleep(1.5)
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
                time.sleep(1.5)
                
        except KeyboardInterrupt:
            print(f"\n\n{RED}✗ Seleção cancelada pelo usuário.{RESET}")
            log("Seleção interrompida por Ctrl+C")
            raise SystemExit(0)
        except Exception as e:
            print(f"\n{RED}✗ Erro inesperado: {str(e)}{RESET}")
            log(f"ERRO durante seleção de perfil: {e}")
            time.sleep(1.5)
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


def ensure_prompt_focus() -> None:
    """Garante que os prompts do PyAutoGUI ganhem foco ao aparecer."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        
        # Permite que a próxima janela possa ganhar foco
        user32.AllowSetForegroundWindow(-1)  # ASFW_ANY
        
        # Obtém o ID do thread e processo atual
        current_thread = kernel32.GetCurrentThreadId()
        foreground_thread = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
        
        # Anexa o input do thread atual ao thread em foreground
        if foreground_thread != current_thread:
            user32.AttachThreadInput(foreground_thread, current_thread, True)
            
        # Força mudança de foco
        user32.BringWindowToTop(user32.GetForegroundWindow())
        user32.SetFocus(user32.GetForegroundWindow())
        
        # Desanexa os threads
        if foreground_thread != current_thread:
            user32.AttachThreadInput(foreground_thread, current_thread, False)
            
        time.sleep(0.15)
    except Exception:
        pass


def make_window_topmost(hwnd) -> None:
    """Define uma janela como topmost (sempre no topo)."""
    if os.name != "nt" or not hwnd:
        return
    try:
        user32 = ctypes.windll.user32
        # Constantes do Windows
        HWND_TOPMOST = -1
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040
        
        # Define a janela como topmost
        user32.SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        )
        # Garante que está em primeiro plano
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
        user32.BringWindowToTop(hwnd)
    except Exception:
        pass


def find_and_focus_pymsgbox() -> None:
    """Encontra e foca a janela do PyMsgBox (usado por pyautogui) e o campo de entrada."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        
        # Callback para enumerar janelas
        def enum_callback(hwnd, lParam):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buffer, length + 1)
                    # PyAutoGUI usa PyMsgBox que cria janelas com classe específica
                    class_buffer = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, class_buffer, 256)
                    # Procura por janelas do tipo dialog ou PyMsgBox
                    if "#32770" in class_buffer.value or "tk" in class_buffer.value.lower():
                        make_window_topmost(hwnd)
                        # Focar no campo de entrada dentro da janela
                        try:
                            # Procurar controle Edit (campo de texto) dentro da janela
                            edit_hwnd = user32.FindWindowExW(hwnd, None, "Edit", None)
                            if edit_hwnd:
                                user32.SetFocus(edit_hwnd)
                                time.sleep(0.05)
                                # Selecionar todo o texto para facilitar digitação
                                user32.SendMessageW(edit_hwnd, 0x00B1, 0, -1)  # EM_SETSEL
                                return False  # Parar de enumerar
                        except Exception:
                            pass
            return True
        
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    except Exception:
        pass


def focused_prompt(text: str = "", title: str = "", default: str = ""):
    """Wrapper para pyautogui.prompt com foco garantido."""
    pause_automation_timer()  # Pausar timer durante prompt
    ensure_prompt_focus()
    time.sleep(0.2)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.3)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.1, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.prompt(text=text, title=title, default=default)
    finally:
        timer.cancel()
        resume_automation_timer()  # Resumir timer após prompt
    
    return result


def focused_alert(text: str = "", title: str = "", button: str = "OK"):
    """Wrapper para pyautogui.alert com foco garantido."""
    pause_automation_timer()  # Pausar timer durante alert
    ensure_prompt_focus()
    time.sleep(0.2)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.3)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.1, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.alert(text=text, title=title, button=button)
    finally:
        timer.cancel()
        resume_automation_timer()  # Resumir timer após alert
    
    return result


def focused_confirm(text: str = "", title: str = "", buttons=None):
    """Wrapper para pyautogui.confirm com foco garantido."""
    pause_automation_timer()  # Pausar timer durante confirm
    ensure_prompt_focus()
    time.sleep(0.2)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.3)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.1, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.confirm(text=text, title=title, buttons=buttons)
    finally:
        timer.cancel()
        resume_automation_timer()  # Resumir timer após confirm
    
    return result
    return result_container[0]


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


def focus_browser_if_needed() -> None:
    """Só pressiona Win+1 se o navegador não estiver em foco, evitando minimizar."""
    title = _get_foreground_title().lower()
    cls = _get_foreground_class().lower()
    title_hits = ("chrome", "edge", "navegador", "invoisys", "google chrome", "microsoft edge")
    class_hits = ("chrome_widgetwin_1", "applicationframewindow", "windows.ui.core.corewindow")

    if any(k in title for k in title_hits) or (cls in class_hits):
        log("Navegador já em foco; Win+1 ignorado para evitar minimizar.")
        return

    log("Navegador fora de foco; tentando Win+1.")
    pyautogui.hotkey("winleft", "1")
    time.sleep(1)

    # Pós-checagem: se ainda não estiver em foco, tentar fallback suave
    title2 = _get_foreground_title().lower()
    cls2 = _get_foreground_class().lower()
    if any(k in title2 for k in title_hits) or (cls2 in class_hits):
        log("Navegador em foco após Win+1.")
        return
    log("Win+1 não focou o navegador; evitando minimizar e mantendo estado.")


def upload_latest_xml() -> None:
    time.sleep(1)
    downloads_path = Path.home() / "Downloads"
    list_of_files = list(downloads_path.glob("*"))
    if not list_of_files:
        focused_alert("A pasta Downloads está vazia!")
        return
    latest_file = max(list_of_files, key=os.path.getctime)
    pyautogui.write(str(latest_file), interval=0.12)
    time.sleep(0.3)
    pyautogui.press("enter")





def wait_for_form(target_text: str, tempo_maximo: float = 15.0, intervalo: float = 1.0, copy_attempts: int = 2) -> str:
    """Aguarda o formulário abrir detectando texto específico (lógica do legado).
    Retorna o conteúdo do clipboard capturado quando o formulário é encontrado."""
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


def navigate_to_mdfe() -> None:
    """Navega para o formulário MDF-e - cópia exata do script legado"""
    # IR PARA 3ª PAGINA - MDF
    pyautogui.hotkey("ctrl", "3")
    time.sleep(1)

    # ABRIR DADOS DO MDF-E
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'EMITIR NOTA'")
    pyautogui.write("EMITIR NOTA", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.7)
    
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'MDF-E'")
    pyautogui.write("MDF-E", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.7)


def select_ncm(profile: ConfigProfile) -> str:
    """Exibe prompt para seleção de NCM no início da automação.
    Retorna o código NCM selecionado."""
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
    
    opcao = focused_confirm(
        text='Selecione o código NCM ou escolha "Outro código" para digitar manualmente:',
        title='Escolha de NCM',
        buttons=[ncm_primary, ncm_secondary, ncm_tertiary, 'Outro código', 'Cancelar']
    )

    if opcao == 'Cancelar':
        focused_alert('Nenhum código NCM selecionado. O script foi pausado.')
        pyautogui.FAILSAFE = True
        raise SystemExit(1)
    elif opcao == 'Outro código':
        codigo = focused_prompt('Digite o código NCM:')
        if not codigo:
            focused_alert('Nenhum código NCM digitado. O script foi pausado.')
            pyautogui.FAILSAFE = True
            raise SystemExit(1)
    else:
        codigo = opcao
    
    log(f"Código NCM selecionado: {codigo}")
    return codigo


def fill_mdfe(profile: ConfigProfile, codigo_ncm: str) -> None:
    """Preenche dados do MDF-e - cópia exata do script legado
    Recebe o código NCM já selecionado como parâmetro."""
    time.sleep(1)
    log("Iniciando preenchimento MDF-e: PRESTADOR DE SERVIÇO, EMITENTE, UF, MUNICÍPIO")
    
    # PRESTADOR DE SERVIÇO
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("SELECIONE...", interval=0.20)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    prestador = profile.get_value('mdfe', 'prestador_tipo')
    pyautogui.write(prestador, interval=0.1)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    # EMITENTE
    emitente = profile.get_value("mdfe", "emitente_codigo")
    pyautogui.write(emitente, interval=0.10)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)
    skip_tabs(7)
    pyautogui.press("space")
    time.sleep(0.3)

    # UF CARREGAMENTO E DESCARREGAMENTO
    uf_car = profile.get_value("mdfe", "uf_carregamento")
    pyautogui.write(uf_car, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("space")
    time.sleep(0.3)
    uf_desc = profile.get_value("mdfe", "uf_descarga")
    pyautogui.write(uf_desc, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(0.2)
    municipio = profile.get_value("mdfe", "municipio_carregamento").upper()
    pyautogui.write(municipio, interval=0.15)
    time.sleep(0.3)

    for _ in range(4):
        pyautogui.press("down")
        time.sleep(0.1)
    for _ in range(3):
        pyautogui.press("up")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(0.3)

    log(f"MDF-e: Prestador={prestador}, Emitente={emitente}, UF_Car={uf_car}, UF_Desc={uf_desc}, Municipio={municipio}")
    
    # UPLOAD DO ARQUIVO XML
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(1.5)
    log("Carregando arquivo XML...")
    upload_latest_xml()
    time.sleep(0.5)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.5)

    # UNIDADE DE MEDIDA, TIPO CARGA E DESCRIÇÃO
    skip_tabs(5)
    pyautogui.press("space")
    time.sleep(0.2)
    unidade = profile.get_value("mdfe", "unidade_medida")
    pyautogui.write(unidade, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    carga = profile.get_value("mdfe", "carga_tipo")
    pyautogui.write(carga, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    log(f"MDF-e: Unidade={unidade}, Tipo_Carga={carga}")

    # DESCRIÇÃO DO PRODUTO
    pyautogui.press("tab")
    time.sleep(0.2)
    descricao = profile.get_value("mdfe", "codigo_produto_descricao")
    log(f"Preenchendo DESCRIÇÃO PRODUTO: {descricao}")
    pyautogui.write(descricao, interval=0.1)
    time.sleep(0.2)

    # CÓDIGO NCM (já selecionado e passado como parâmetro)
    skip_tabs(2)
    
    codigo_ncm_upper = codigo_ncm.upper()
    pyautogui.write(codigo_ncm_upper, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # CEP ORIGEM E DESTINO
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    cep_orig = profile.get_value("mdfe", "cep_origem")
    pyautogui.write(cep_orig, interval=0.1)
    time.sleep(0.3)

    skip_tabs(3)
    cep_dest = profile.get_value("mdfe", "cep_destino")
    pyautogui.write(cep_dest, interval=0.1)
    time.sleep(1.25)
    
    log(f"MDF-e concluído: NCM={codigo_ncm_upper}, CEP_Orig={cep_orig}, CEP_Dest={cep_dest}")


def fill_modal_rodo(profile: ConfigProfile) -> None:
    """Preenche dados do Modal Rodoviário"""
    log("Iniciando preenchimento Modal Rodoviário")
    time.sleep(1.5)
    
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("modal rodo", interval=0.10)
    time.sleep(0.2)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(1.25)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.25)

    # RNTRC, CONTRATANTE, CNPJ
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("RNTRC", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    rntrc = profile.get_value("modal_rodoviario", "rntrc")
    pyautogui.write(rntrc, interval=0.10)
    time.sleep(0.2)
    
    skip_tabs(6)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    contratante = profile.get_value("modal_rodoviario", "contratante_nome")
    pyautogui.write(contratante, interval=0.20)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont = profile.get_value("modal_rodoviario", "contratante_cnpj")
    pyautogui.write(cnpj_cont, interval=0.12)
    time.sleep(0.2)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(1.25)
    
    log(f"Modal Rodoviário: RNTRC={rntrc}, Contratante={contratante}, CNPJ={cnpj_cont}")


def fill_additional_info(profile: ConfigProfile) -> None:
    """Preenche Informações Adicionais (Seguradora, Frete, Banco, etc)"""
    log("Iniciando preenchimento Informações Adicionais")
    
    # ABRIR SEÇÃO OPCIONAIS
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("OPCIONAIS", interval=0.10)
    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("esc")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(1)

    # ADICIONAIS - CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    pyautogui.write("ADICIONAIS", interval=0.10)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)

    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    pyautogui.write("CONTRIBUINTE", interval=0.10)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    skip_tabs(3)
    cnpj_contrib = profile.get_value("informacoes_adicionais", "contribuinte_cnpj")
    pyautogui.write(cnpj_contrib, interval=0.10)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("enter")

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write("CONTRA", interval=0.10)
    pyautogui.press("enter")
    time.sleep(0.3)
    
    # Seguradora e dados relacionados
    pyautogui.press("tab")
    cnpj_cont2 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    pyautogui.write(cnpj_cont2, interval=0.12)
    pyautogui.press("tab")
    seguradora = profile.get_value("informacoes_adicionais", "seguradora_nome")
    pyautogui.write(seguradora, interval=0.10)
    pyautogui.press("tab")
    cnpj_seg = profile.get_value("informacoes_adicionais", "seguradora_cnpj")
    pyautogui.write(cnpj_seg, interval=0.12)
    pyautogui.press("tab")
    apolice = profile.get_value("informacoes_adicionais", "numero_apolice")
    pyautogui.write(apolice, interval=0.10)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(0.3)

    skip_tabs(4)
    pyautogui.press("space")
    time.sleep(0.2)

    # Contratante 2 e Frete
    pyautogui.press("tab")
    time.sleep(0.2)
    contratante2 = profile.get_value("modal_rodoviario", "contratante_nome")
    pyautogui.write(contratante2, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont3 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    pyautogui.write(cnpj_cont3, interval=0.12)
    time.sleep(0.2)

    skip_tabs(2)
    frete_val = profile.get_value("informacoes_adicionais", "frete_valor")
    pyautogui.write(frete_val, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    forma_pag = profile.get_value("informacoes_adicionais", "forma_pagamento")
    pyautogui.write(forma_pag, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("enter")
    time.sleep(0.3)

    # Banco e Agência
    pyautogui.press("tab")
    time.sleep(0.2)
    numero_banco = profile.get_value("informacoes_adicionais", "numero_banco")
    pyautogui.write(numero_banco, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    agencia = profile.get_value("informacoes_adicionais", "agencia")
    pyautogui.write(agencia, interval=0.12)
    time.sleep(0.2)

    skip_tabs(2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # Seção FRETE
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.write("FRETE", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_val2 = profile.get_value("informacoes_adicionais", "frete_valor")
    pyautogui.write(frete_val2, interval=0.12)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_tipo = profile.get_value("informacoes_adicionais", "frete_tipo")
    pyautogui.write(frete_tipo, interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)

    # Seção de Parcelas
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.3)

    skip_tabs(5)
    numero_parcelas = profile.get_value("informacoes_adicionais", "numero_parcelas")
    if not numero_parcelas:
        log("Perfil sem 'informacoes_adicionais.numero_parcelas' obrigatório")
        focused_alert(
            "O perfil está faltando a chave obrigatória: informacoes_adicionais.numero_parcelas",
            title="Perfil inválido"
        )
        raise SystemExit(1)
    pyautogui.write(numero_parcelas, interval=0.10)
    time.sleep(0.15)

    skip_tabs(2)
    pyautogui.press("space")
    time.sleep(0.2)

    skip_tabs(7)
    for _ in range(2):
        pyautogui.press("space")
        time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    frete_val3 = profile.get_value("informacoes_adicionais", "frete_valor")
    pyautogui.write(frete_val3, interval=0.10)
    time.sleep(0.15)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.0)
    
    log(f"Info Adicionais: Seguradora={seguradora}, Frete={frete_val}, Banco={numero_banco}, Parcelas={numero_parcelas}")


def perform_averbacao(numero_cte: str = "", numero_dt: str = "", nf_concat: str = "") -> None:
    """Executa a averbação e preenche DT/CT-e no final.
    - Usa CT-e capturado no início quando disponível, sem voltar à primeira página
    - Em fallback, captura CT-e na INVOISYS sem atrelar ou sobrescrever o DT
    """
    # Abrir site/aba de averbação e enviar XML
    pyautogui.hotkey("ctrl", "4")
    time.sleep(0.5)

    ##Pequeno hotfix para GAP relacionado a envio de XMLs, verificar alternativas
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)

    for search in ("OK", "XML", "ENVIAR"):
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.4)
        pyautogui.write(search, interval=0.1)
        time.sleep(0.4)
        pyautogui.press("esc")
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.3)

    upload_latest_xml()
    time.sleep(2.5)

    # Extrair número de averbação e copiar apenas os dígitos
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.2)
    texto = pyperclip.paste()
    match = re.search(r"Número de Averbação:\s*([\d]+)", texto)
    if match:
        numero_averbacao = match.group(1)
        pyperclip.copy(numero_averbacao)
        print("Número de Averbação copiado:", numero_averbacao)
    else:
        print("Número de Averbação não encontrado")

    time.sleep(0.5)
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.7)

    # Preencher detalhes na outra aba
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.4)
    pyautogui.write("DETALHES", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.7)

    # Preencher DT/CT-e/NF na área de CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("CONTRIBUINTE", interval=0.15)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)

    # DT sempre vem do prompt (numero_dt) capturado no início
    pyautogui.write("DT: ", interval=0.1)
    time.sleep(0.3)
    if numero_dt:
        pyperclip.copy(numero_dt)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    # CT-e: usar o valor informado pelo usuário
    pyautogui.write(" CTE: ", interval=0.1)
    time.sleep(0.5)
    if numero_cte:
        log(f"Usando CT-e informado pelo usuário: {numero_cte}")
        pyperclip.copy(numero_cte)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
    else:
        log("CT-e não informado; campo ficará vazio")
        time.sleep(0.5)

    # Finalizar com rótulo NF sempre; deixa vazio se não informado
    if nf_concat:
        pyautogui.write(f" NF: {nf_concat}", interval=0.1)
    else:
        pyautogui.write(" NF: ", interval=0.1)
    time.sleep(0.5)


def main() -> None:
    global _automation_start_time, _automation_time_paused
    
    try:
        # Iniciar contadores de tempo
        real_start_time = time.monotonic()
        _automation_time_paused = 0.0
        
        # Sleep inicial para aguardar inicialização
        time.sleep(0.7)
        log("Iniciando automação (main)")
        log(f"[DEBUG] real_start_time={real_start_time}")
        
        # Limpar tela e mostrar header
        os.system('cls' if os.name == 'nt' else 'clear')
        ui_print("AUTOMAÇÃO MDF-e", style="header")
        
        # Impedir execução duplicada
        ensure_single_instance()
        
        parser = argparse.ArgumentParser()
        parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
        args = parser.parse_args()

        selected = args.profile
        if not selected:
            log("Exibindo menu de seleção de perfil")
            selected = choose_profile(list_profiles())
        log(f"Perfil selecionado: {selected}")
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
        log(f"Perfil carregado com sucesso de: {profile_path}")

        # Iniciar cronômetro de automação apenas após a escolha do perfil
        _automation_start_time = time.monotonic()
        _automation_time_paused = 0.0
        log(f"[DEBUG] _automation_start_time iniciado após escolha do perfil: {_automation_start_time}")

        ui_print("Iniciando preenchimento...", style="step")
        
        # Abrir/focar navegador sem minimizar (usa Win+1 só se não estiver em foco)
        log("Focando navegador (evitando minimizar)")
        focus_browser_if_needed()
        
        # GAP - Pressionar ESC 2x
        log("Enviando ESC x2")
        for _ in range(2):
            pyautogui.press("esc")
            time.sleep(0.3)

        # Recarregar aba 3
        log("Recarregando aba 3 (Ctrl+3, F5)")
        pyautogui.hotkey("ctrl", "3")
        time.sleep(0.5)
        pyautogui.press("f5")
        time.sleep(1)

        # Voltar para aba 1 (uma vez, como no legado)
        log("Voltando para aba 1 e focando em 'empresa'")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(0.5)
        # Tab para garantir foco correto
        pyautogui.press("tab")
        time.sleep(0.2)

        # VALIDAÇÃO DA PÁGINA CT-E (antes do prompt de DT)
        log("Validando página CT-e antes de solicitar DT...")
        pyautogui.press("tab")
        time.sleep(0.2)
        time.sleep(3)
        try:
            conteudo_validacao = wait_for_form("notas emitidas: ct-e", tempo_maximo=4, intervalo=1, copy_attempts=2)
            log("Página CT-e detectada. Verificando presença de 'NÚMERO CT-E'...")
            
            # Verificar se a página contém "NÚMERO CT-E"
            conteudo_upper = conteudo_validacao.upper()
            if "NÚMERO CT-E" not in conteudo_upper and "NUMERO CT-E" not in conteudo_upper:
                log("ERRO: 'NÚMERO CT-E' não encontrado na página. Página incorreta detectada!")
                focused_alert(
                    "ERRO: Página de CT-e não foi reconhecida!\n\n"
                    "A automação foi interrompida porque não foi possível identificar\n"
                    "a página correta de notas emitidas (CT-e).\n\n"
                    "Verifique:\n"
                    "• Se você está na página correta do sistema\n"
                    "• Se a primeira aba do navegador está aberta no Invoisys em NOTAS EMITIDAS > CT-e\n"

                    "A automação será encerrada.",
                    title="ERRO: Página CT-e não Reconhecida"
                )
                raise SystemExit(1)
            
            log("'NÚMERO CT-E' confirmado. Página válida para continuar.")
        except SystemExit as e:
            # Re-lançar SystemExit para interromper automação
            if e.code == 1:
                raise
            log("Aviso: Página de CT-e não foi encontrada durante validação.")
            focused_alert(
                "ERRO: Página de CT-e não foi encontrada!\n\n"
                "A automação foi interrompida porque a página esperada\n"
                "de notas emitidas (CT-e) não foi detectada.\n\n"
                "Verifique:\n"
                "• Se você está logado no sistema\n"
                "• Se a primeira aba do navegador está aberta no Invoisys em NOTAS EMITIDAS > CT-e\n"
                
                "A automação será encerrada.",
                title="ERRO: Página de CT-e não Encontrada"
            )
            raise SystemExit(1)

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
        numero_dt = focused_prompt(text=prompt_text, title="DT")
        if not numero_dt:
            focused_alert("Nenhum código DT informado. O script foi pausado.")
            pyautogui.FAILSAFE = True
            return
        log(f"DT informado: {numero_dt}")

        # Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(2)
        pyautogui.write("DO DT", interval=0.12)
        pyautogui.press("esc")
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.7)
        
        # Usar o DT armazenado previamente
        log(f"Preenchendo campo DT com valor armazenado: {numero_dt}")
        pyautogui.write(numero_dt.upper(), interval=0.1)
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(0.5)

        # SELEÇÃO DE NCM (APÓS DT E ANTES DE CTE) - NOVO FLUXO
        log("Exibindo prompt para seleção de NCM")
        codigo_ncm = select_ncm(profile)
        log(f"NCM selecionado e armazenado: {codigo_ncm}")
        ui_print(f"NCM selecionado: {codigo_ncm}", style="success")
        time.sleep(0.5)

        # PROMPTS PARA NFs (NF1/NF2) E CONCATENAÇÃO (opcionais)
        log("Exibindo prompt para NF1")
        nf1 = focused_prompt(text="Informe a NF1 (opcional):", title="NF1") or ""
        log(f"NF1 informada: '{nf1}'")
        
        log("Exibindo prompt para NF2")
        nf2 = focused_prompt(text="Informe a NF2 (opcional):", title="NF2") or ""
        log(f"NF2 informada: '{nf2}'")
        
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
        time.sleep(0.3)

        # Aviso inicial para baixar o CT-e, seguido do prompt do número
        log("Exibindo aviso para baixar o CT-e antes de seguir")
        focused_alert(
            text=(
                "Antes de prosseguir, faça o download do XML do CT-e correspondente \n"
                "à DT e mantenha-o salvo. Em seguida clique em OK para continuar."
            ),
            title="Aviso: Baixe o CT-e"
        )
        numero_cte = focused_prompt(
            text="Informe o número do CT-e (XML já baixado):",
            title="Número CT-e"
        ) or ""
        log(f"CT-e informado: '{numero_cte}'")
        if not numero_cte:
            log("Nenhum número de CT-e informado; encerrando automação conforme solicitado")
            focused_alert(
                "ERRO: Nenhum número de CT-e foi informado.\n\n"
                "A automação será encerrada.",
                title="CT-e obrigatório"
            )
            raise SystemExit(1)
        pyautogui.press("esc")
        time.sleep(0.15)
        ensure_caps_off()
        
        ui_print("Preenchendo formulário MDF-e...", style="step")
        
        # Navegar para MDF-e e detectar formulário (lógica e tempos do legado)
        navigate_to_mdfe()
        wait_for_form("Emissor MDF-e", tempo_maximo=15.0, intervalo=3.0, copy_attempts=3)
        
        # Preencher formulário (passando codigo_ncm já selecionado)
        log("Iniciando preenchimento dos dados MDF-e")
        fill_mdfe(profile, codigo_ncm)
        log("Dados MDF-e preenchidos com sucesso")
        ui_print("Dados MDF-e preenchidos", style="success")
        
        log("Iniciando preenchimento do modal rodoviário")
        fill_modal_rodo(profile)
        log("Modal rodoviário preenchido com sucesso")
        ui_print("Modal rodoviário preenchido", style="success")
        
        log("Iniciando preenchimento de informações adicionais")
        fill_additional_info(profile)
        log("Informações adicionais preenchidas com sucesso")
        ui_print("Informações adicionais preenchidas", style="success")
        
        log("Iniciando processamento de averbação")
        ui_print("Processando averbação...", style="step")
        perform_averbacao(numero_cte, numero_dt, nf_concat)
        log("Averbação processada com sucesso")
        
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


if __name__ == "__main__":
    main()