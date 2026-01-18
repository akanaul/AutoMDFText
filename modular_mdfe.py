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


def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Não interromper fluxo por falha de log em disco
        pass


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
                log("Lista de scripts ficou vazia durante seleção; usando template padrão.")
                return default
            
            # Atualizar lista se mudou
            if current_list != interactive_list:
                interactive_list = current_list
                log(f"Lista de scripts atualizada: {len(interactive_list)} scripts disponíveis.")
            
            # Menu destacado com cores e separadores
            print(f"\n{CYAN}{'=' * 60}{RESET}")
            print(f"{BOLD}{GREEN}  SELEÇÃO DE SCRIPT - AUTOMAÇÃO MDF-e{RESET}")
            print(f"{CYAN}{'=' * 60}{RESET}\n")
            
            for idx, name in enumerate(interactive_list, start=1):
                print(f"{YELLOW}  [{idx}]{RESET} {name}")
            
            print(f"\n{CYAN}{'─' * 60}{RESET}")
            print(f"{BOLD}Digite o número do script desejado:{RESET}")
            print(f"{CYAN}{'─' * 60}{RESET}\n")
            
            choice = input(f"{BOLD}Opção: {RESET}").strip()
            
            # Validar entrada
            if not choice:
                print(f"\n{RED}✗ Erro: Você deve selecionar um script!{RESET}")
                log("Entrada vazia; solicitando nova seleção.")
                time.sleep(1.5)
                continue
            
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
        log(f"Número máximo de tentativas atingido ({max_attempts}); usando template padrão.")
        return default
    
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
    """Encontra e foca a janela do PyMsgBox (usado por pyautogui)."""
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
            return True
        
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    except Exception:
        pass


def focused_prompt(text: str = "", title: str = "", default: str = ""):
    """Wrapper para pyautogui.prompt com foco garantido."""
    ensure_prompt_focus()
    time.sleep(0.1)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.15)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.05, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.prompt(text=text, title=title, default=default)
    finally:
        timer.cancel()
    
    return result


def focused_alert(text: str = "", title: str = "", button: str = "OK"):
    """Wrapper para pyautogui.alert com foco garantido."""
    ensure_prompt_focus()
    time.sleep(0.1)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.15)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.05, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.alert(text=text, title=title, button=button)
    finally:
        timer.cancel()
    
    return result


def focused_confirm(text: str = "", title: str = "", buttons=None):
    """Wrapper para pyautogui.confirm com foco garantido."""
    ensure_prompt_focus()
    time.sleep(0.1)
    
    # Criar um timer para forçar topmost logo após a janela ser criada
    def force_topmost():
        time.sleep(0.15)
        find_and_focus_pymsgbox()
    
    timer = threading.Timer(0.05, force_topmost)
    timer.daemon = True
    timer.start()
    
    try:
        result = pyautogui.confirm(text=text, title=title, buttons=buttons)
    finally:
        timer.cancel()
    
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


def extract_cte_from_content(conteudo: str) -> str:
    """Extrai o número do CT-e do conteúdo capturado."""
    linhas = conteudo.splitlines()
    for linha in linhas:
        # Padrão: "100 - Autorizado o uso do CT-e.N" seguido de número com 6 dígitos
        if "100 - Autorizado o uso do CT-e" in linha:
            # Extrai 6 dígitos consecutivos após "CT-e"
            match = re.search(r'CT-e\.?[^\d]*(\d{6})', linha)
            if match:
                return match.group(1)
    return ""


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


def fill_mdfe(profile: ConfigProfile) -> str:
    """Preenche dados do MDF-e - cópia exata do script legado
    Retorna o código NCM selecionado."""
    time.sleep(1)
    # PRESTADOR DE SERVIÇO
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log(f"Preenchendo PRESTADOR DE SERVIÇO: {profile.get_value('mdfe', 'prestador_tipo')}")
    pyautogui.write("SELECIONE...", interval=0.20)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.write(profile.get_value("mdfe", "prestador_tipo"), interval=0.1)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    # EMITENTE
    emitente = profile.get_value("mdfe", "emitente_codigo")
    log(f"Preenchendo EMITENTE: {emitente}")
    pyautogui.write(emitente, interval=0.10)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.3)

    # UF CARREGAMENTO E DESCARREGAMENTO
    uf_car = profile.get_value("mdfe", "uf_carregamento")
    log(f"Preenchendo UF CARREGAMENTO: {uf_car}")
    pyautogui.write(uf_car, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("space")
    time.sleep(0.3)
    uf_desc = profile.get_value("mdfe", "uf_descarga")
    log(f"Preenchendo UF DESCARGA: {uf_desc}")
    pyautogui.write(uf_desc, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(0.2)
    municipio = profile.get_value("mdfe", "municipio_carregamento").upper()
    log(f"Preenchendo MUNICIPIO CARREGAMENTO: {municipio}")
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

    # UPLOAD DO ARQUIVO XML
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(1.5)
    upload_latest_xml()
    time.sleep(0.5)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.5)

    # UNIDADE DE MEDIDA
    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    unidade = profile.get_value("mdfe", "unidade_medida")
    log(f"Preenchendo UNIDADE MEDIDA: {unidade}")
    pyautogui.write(unidade, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # TIPO DE CARGA
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    carga = profile.get_value("mdfe", "carga_tipo")
    log(f"Preenchendo TIPO CARGA: {carga}")
    pyautogui.write(carga, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # DESCRIÇÃO DO PRODUTO
    pyautogui.press("tab")
    time.sleep(0.2)
    descricao = profile.get_value("mdfe", "codigo_produto_descricao")
    log(f"Preenchendo DESCRIÇÃO PRODUTO: {descricao}")
    pyautogui.write(descricao, interval=0.1)
    time.sleep(0.2)

    # CÓDIGO NCM
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    
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
    pyautogui.write(codigo.upper(), interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # CEP ORIGEM
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    cep_orig = profile.get_value("mdfe", "cep_origem")
    log(f"Preenchendo CEP ORIGEM: {cep_orig}")
    pyautogui.write(cep_orig, interval=0.1)
    time.sleep(0.3)

    # CEP DESTINO
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.15)
    cep_dest = profile.get_value("mdfe", "cep_destino")
    log(f"Preenchendo CEP DESTINO: {cep_dest}")
    pyautogui.write(cep_dest, interval=0.1)
    time.sleep(1.25)
    
    return codigo


def fill_modal_rodo(profile: ConfigProfile) -> None:
    # ABRIR DADOS DO MODAL RODOVIÁRIO
    time.sleep(1.5)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'modal rodo'")
    pyautogui.write("modal rodo", interval=0.10)
    time.sleep(0.2)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.25)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.25)

    # RNTRC
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("RNTRC", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    rntrc = profile.get_value("modal_rodoviario", "rntrc")
    log(f"Preenchendo RNTRC: {rntrc}")
    pyautogui.write(rntrc, interval=0.10)
    time.sleep(0.2)
    
    # NOME DO CONTRATANTE
    for _ in range(6):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    contratante = profile.get_value("modal_rodoviario", "contratante_nome")
    log(f"Preenchendo CONTRATANTE: {contratante}")
    pyautogui.write(contratante, interval=0.20)
    time.sleep(0.2)

    # CNPJ DO CONTRATATANTE
    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE: {cnpj_cont}")
    pyautogui.write(cnpj_cont, interval=0.12)
    time.sleep(0.2)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("enter")
    time.sleep(1.25)


def fill_additional_info(profile: ConfigProfile) -> None:
    # ABRIR DADOS DE INFORMAÇÕES OPCIONAIS
    pyautogui.hotkey("ctrl", "f")
    log("Procurando por 'OPCIONAIS'")
    pyautogui.write("OPCIONAIS", interval=0.10)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("esc")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(1)

    # INFORMAÇÕES ADICIONAIS
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    log("Procurando por 'ADICIONAIS'")
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
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.3)
    cnpj_contrib = profile.get_value("informacoes_adicionais", "contribuinte_cnpj")
    log(f"Preenchendo CONTRIBUINTE CNPJ: {cnpj_contrib}")
    pyautogui.write(cnpj_contrib, interval=0.10)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("enter")

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write("CONTRA", interval=0.10)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    cnpj_cont2 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE (CONTRA): {cnpj_cont2}")
    pyautogui.write(cnpj_cont2, interval=0.12)
    pyautogui.press("tab")
    seguradora = profile.get_value("informacoes_adicionais", "seguradora_nome")
    log(f"Preenchendo SEGURADORA: {seguradora}")
    pyautogui.write(seguradora, interval=0.10)
    pyautogui.press("tab")
    cnpj_seg = profile.get_value("informacoes_adicionais", "seguradora_cnpj")
    log(f"Preenchendo CNPJ SEGURADORA: {cnpj_seg}")
    pyautogui.write(cnpj_seg, interval=0.12)
    pyautogui.press("tab")
    apolice = profile.get_value("informacoes_adicionais", "numero_apolice")
    log(f"Preenchendo NUMERO APOLICE: {apolice}")
    pyautogui.write(apolice, interval=0.10)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(0.3)

    for _ in range(4):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    contratante2 = profile.get_value("modal_rodoviario", "contratante_nome")
    log(f"Preenchendo CONTRATANTE (2): {contratante2}")
    pyautogui.write(contratante2, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont3 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE (2): {cnpj_cont3}")
    pyautogui.write(cnpj_cont3, interval=0.12)
    time.sleep(0.2)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)

    frete_val = profile.get_value("informacoes_adicionais", "frete_valor")
    log(f"Preenchendo FRETE VALOR: {frete_val}")
    pyautogui.write(frete_val, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    forma_pag = profile.get_value("informacoes_adicionais", "forma_pagamento")
    log(f"Preenchendo FORMA PAGAMENTO: {forma_pag}")
    pyautogui.write(forma_pag, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("enter")
    time.sleep(0.3)

    # Garantir que o foco chegue no campo "Número do banco" (antes de agência)
    pyautogui.press("tab")
    time.sleep(0.2)
    numero_banco = profile.get_value("informacoes_adicionais", "numero_banco")
    log(f"Preenchendo NUMERO BANCO: {numero_banco}")
    pyautogui.write(numero_banco, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    agencia = profile.get_value("informacoes_adicionais", "agencia")
    log(f"Preenchendo AGENCIA: {agencia}")
    pyautogui.write(agencia, interval=0.12)
    time.sleep(0.2)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)

    pyautogui.press("enter")
    time.sleep(0.3)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    log("Preenchendo seção FRETE")
    pyautogui.write("FRETE", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_val2 = profile.get_value("informacoes_adicionais", "frete_valor")
    log(f"Preenchendo FRETE VALOR (seção): {frete_val2}")
    pyautogui.write(frete_val2, interval=0.12)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_tipo = profile.get_value("informacoes_adicionais", "frete_tipo")
    log(f"Preenchendo FRETE TIPO: {frete_tipo}")
    pyautogui.write(frete_tipo, interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.3)

    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.2)
    numero_parcelas = profile.get_value("informacoes_adicionais", "numero_parcelas")
    if not numero_parcelas:
        log("Perfil sem 'informacoes_adicionais.numero_parcelas' obrigatório")
        focused_alert(
            "O perfil está faltando a chave obrigatória: informacoes_adicionais.numero_parcelas",
            title="Perfil inválido"
        )
        raise SystemExit(1)
    log(f"Preenchendo NUMERO PARCELAS: {numero_parcelas}")
    pyautogui.write(numero_parcelas, interval=0.10)
    time.sleep(0.15)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.3)
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
    log(f"Preenchendo FRETE VALOR (final): {frete_val3}")
    pyautogui.write(frete_val3, interval=0.10)
    time.sleep(0.15)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.0)


def perform_averbacao(numero_cte: str = "", numero_dt: str = "") -> None:
    """Executa a averbação e preenche DT/CT-e no final.
    - Usa CT-e capturado no início quando disponível, sem voltar à primeira página
    - Em fallback, captura CT-e na INVOISYS sem atrelar ou sobrescrever o DT
    """
    # Abrir site/aba de averbação e enviar XML
    pyautogui.hotkey("ctrl", "shift", "a")
    time.sleep(0.5)
    pyautogui.write("ATM", interval=0.1)
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1)

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
    pyautogui.press("enter")
    time.sleep(0.4)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    time.sleep(0.4)
    pyautogui.press("enter")
    time.sleep(0.7)

    # Preencher DT/CT-e/NF na área de CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.4)
    pyautogui.write("CONTRIBUINTE", interval=0.15)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.4)
    pyautogui.press("tab")
    time.sleep(0.4)

    # DT sempre vem do prompt (numero_dt) capturado no início
    pyautogui.write("DT: ", interval=0.1)
    time.sleep(0.3)
    if numero_dt:
        pyperclip.copy(numero_dt)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    # CT-e: se já temos do início, usar direto; caso contrário, fallback
    pyautogui.write(" CTE: ", interval=0.1)
    time.sleep(0.5)
    if numero_cte:
        log(f"Usando CT-e capturado no início: {numero_cte}")
        pyperclip.copy(numero_cte)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
    else:
        log("CT-e ausente; iniciando fallback na INVOISYS para capturar.")
        pyautogui.hotkey("ctrl", "shift", "a")
        time.sleep(0.5)
        pyautogui.write("INVOISYS", interval=0.1)
        for _ in range(2):
            pyautogui.press("tab")
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1)

        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        pyautogui.write("e final", interval=0.2)
        pyautogui.press("esc")
        time.sleep(0.2)
        for _ in range(2):
            pyautogui.press("tab")
            time.sleep(0.1)
        time.sleep(0.3)

        # Copiar página e extrair CT-e
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.7)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.8)
        conteudo = pyperclip.paste()
        linhas = conteudo.splitlines()
        numero_cte_local = ""
        for linha in linhas:
            if "100 - Autorizado o uso do CT-e.N" in linha:
                m = re.search(r"100\s*-\s*Autorizado o uso do CT-e\.N[^\d]*(\d{6})", linha)
                if m:
                    numero_cte_local = m.group(1)
                    break

        # Voltar para a aba de preenchimento e colar somente CT-e se encontrado
        pyautogui.hotkey("alt", "tab")
        time.sleep(0.5)
        if numero_cte_local:
            log(f"Número do CT-e encontrado (fallback): {numero_cte_local}")
            pyperclip.copy(numero_cte_local)
            pyautogui.hotkey("ctrl", "v")
        else:
            log("Não foi possível localizar o número do CT-e (fallback); deixar em branco.")
        time.sleep(0.5)

    # Finalizar com rótulo NF
    pyautogui.write(" NF: ", interval=0.1)
    time.sleep(0.5)


def main() -> None:
    try:
        # Sleep inicial para aguardar inicialização
        time.sleep(0.7)
        log("Iniciando automação (main)")
        # Impedir execução duplicada
        ensure_single_instance()
        
        parser = argparse.ArgumentParser()
        parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
        args = parser.parse_args()

        selected = args.profile
        if not selected:
            selected = choose_profile(list_profiles())
        log(f"Perfil selecionado: {selected}")
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
        log(f"Perfil carregado de: {profile_path}")

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
        # Buscar "empresa" para garantir foco correto
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.write("empresa", interval=0.10)
        pyautogui.press("esc")
        time.sleep(0.5)

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

        # Detectar primeira tela (CT-e) APÓS inserir o DT
        log("Detectando tela CT-e (notas emitidas: ct-e)")
        log("Aguardando 4 segundos para o site carregar antes de copiar...")
        pyautogui.press("tab")
        time.sleep(0.2)
        time.sleep(3)
        numero_cte = ""
        try:
            conteudo_cte_pagina = wait_for_form("notas emitidas: ct-e", tempo_maximo=4, intervalo=1, copy_attempts=2)
            log("Página CT-e detectada. Extraindo CT-e do conteúdo capturado...")
            numero_cte = extract_cte_from_content(conteudo_cte_pagina)
            if numero_cte:
                log(f"CT-e extraído com sucesso: {numero_cte}")
                pyperclip.copy(numero_cte)
            else:
                log("Aviso: Não foi possível extrair CT-e do conteúdo.")
        except SystemExit:
            log("Aviso: Texto 'notas emitidas: ct-e' não encontrado. Continuando sem CT-e capturado.")
            numero_cte = ""

        # Alerta solicitando download do XML
        log("Exibindo alerta para download do XML")
        xml_alert = profile.get_value("general", "xml_download_alert")
        if not xml_alert:
            # Compatibilidade com padrão anterior: usar alert_intro se existir
            xml_alert = profile.get_value("general", "alert_intro")
        if not xml_alert:
            # Hardcode apenas se não houver chave no perfil (não padrão antigo)
            xml_alert = (
                'BAIXE O ARQUIVO XML:\n\n'
                '1. Faça o download do arquivo XML da DT buscada;\n'
                '2. Aguarde o download ser concluído;\n'
                '3. Clique em OK para continuar a automação.\n\n'
                'OBS: Para interromper o processo, mova o mouse repetidamente para o canto superior direito da tela.'
            )
        focused_alert(xml_alert)
        # Desativar Caps Lock
        ensure_caps_off()
        
        # Navegar para MDF-e e detectar formulário (lógica e tempos do legado)
        navigate_to_mdfe()
        wait_for_form("Emissor MDF-e", tempo_maximo=15.0, intervalo=3.0, copy_attempts=3)
        
        # Preencher formulário
        codigo_ncm = fill_mdfe(profile)
        fill_modal_rodo(profile)
        fill_additional_info(profile)
        perform_averbacao(numero_cte, numero_dt)
        
        # Ao finalizar com sucesso, restaurar o terminal como popup e emitir beep baixo
        restore_console_popup()
        play_low_beep()
        
        # Exibir resumo no terminal
        GREEN = "\033[92m"
        CYAN = "\033[96m"
        YELLOW = "\033[93m"
        BOLD = "\033[1m"
        RESET = "\033[0m"
        
        print(f"\n{CYAN}{'=' * 60}{RESET}")
        print(f"{BOLD}{GREEN}  AUTOMAÇÃO CONCLUÍDA COM SUCESSO!{RESET}")
        print(f"{CYAN}{'=' * 60}{RESET}\n")
        print(f"{BOLD}Resumo das Informações:{RESET}\n")
        print(f"  {YELLOW}DT:{RESET}      {numero_dt}")
        print(f"  {YELLOW}CT-e:{RESET}    {numero_cte if numero_cte else 'Não capturado'}")
        print(f"  {YELLOW}NCM:{RESET}     {codigo_ncm}")
        print(f"\n{CYAN}{'─' * 60}{RESET}")
        print(f"{BOLD}Próximos passos:{RESET}")
        print(f"  • Inclua a NF")
        print(f"  • Preencha os dados do motorista")
        print(f"\n{CYAN}{'=' * 60}{RESET}\n")
        log("Automação finalizada com sucesso")
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