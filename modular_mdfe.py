#GAPS CORRIDOS RAFAEL 24/03

"""Automacao MDF-e com selecao de perfil, prompts gui e preenchimento via teclado.

Inclui failsafe (F8) e validacoes de tela para reduzir erros de preenchimento em
formularios do navegador.
"""
#IMPOTAR BIBLIOTECAS
import argparse
import ctypes
import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox
import time
import threading
from pathlib import Path

import pyautogui
import pyperclip

# Importar constantes centralizadas
from constants import (
    TAB_DELAY, TAB_DELAY_LONG, CTRL_F_DELAY, DROPDOWN_SETTLE_DELAY,
    SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG, SLEEP_LONGER, SLEEP_ONE, SLEEP_ONE_HALF,
    FORM_WAIT_TIMEOUT, FORM_WAIT_INTERVAL, FORM_COPY_ATTEMPTS, VERIFY_FIELD_TIMEOUT,
    EDGE_SEARCHBAR_HEIGHT, EDGE_CLICK_OFFSET, MIN_PASTE_LENGTH, WRITE_INTERVAL,
    LOG_TIMESTAMP_FORMAT, LOG_MESSAGE_FORMAT, MUTEX_NAME
)

# ── Módulos extraídos (Parte 3) ───────────────────────────────────────────────
from mdfe.logger import log, ui_print, start_automation_session
from mdfe.timing import (
    pause_automation_timer, resume_automation_timer, format_duration,
    _automation_start_time, _automation_time_paused, _pause_start_time,
)
import mdfe.timing as _timing
from mdfe.failsafe import start_failsafe_f8, stop_failsafe_f8, request_pause
import mdfe.failsafe as _failsafe
from mdfe.console import hide_console_window, restore_console_popup, play_low_beep
from mdfe.instance import ensure_single_instance
from mdfe.pause import show_pause_dialog, check_pause, pause_point
from mdfe.profile import parse_profile, ConfigProfile, list_profiles
from mdfe.dialogs import (
    focused_alert, focused_confirm, focused_prompt,
    prompt_dt_blocking, prompt_batch_info,
)
from mdfe.keyboard import (
    _normalize_text, _normalize_digits, paste_text, smart_write,
    press_tab, skip_tabs, ensure_caps_off, upload_latest_xml,
    _last_write_value, _last_write_verify,
)

pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "logs"
CONFIG_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
# Log por sessao (timestamp) para facilitar debug
SESSION_TS = time.strftime(LOG_TIMESTAMP_FORMAT, time.localtime())
LOG_FILE = LOG_DIR / f"automation_{SESSION_TS}.log"
# Handle for single-instance mutex to keep it alive during process lifetime
# (gerenciado por mdfe.instance._SINGLETON_MUTEX_HANDLE)

# Estado de tempo — gerenciado por mdfe.timing
# Estado de failsafe/pausa — gerenciado por mdfe.failsafe
# Estado do último smart_write — gerenciado por mdfe.keyboard
_pause_lock = _failsafe._pause_lock


# _normalize_text → mdfe.keyboard._normalize_text (importado acima)


# log, start_automation_session → mdfe.logger (importado acima)


# start_failsafe_f8, stop_failsafe_f8, request_pause → mdfe.failsafe (importado acima)
# ui_print → mdfe.logger (importado acima)
# pause_automation_timer, resume_automation_timer → mdfe.timing (importado acima)


# show_pause_dialog, check_pause, pause_point → mdfe.pause (importado acima)
# format_duration → mdfe.timing (importado acima)
# _verify_last_write_before_pause → mdfe.pause (importado acima)

# skip_tabs, press_tab, paste_text, smart_write, check_pause, pause_point,
# _verify_last_write_before_pause, format_duration → mdfe.* (importados acima)


# parse_profile, ConfigProfile, list_profiles → mdfe.profile (importado acima)
# hide_console_window, restore_console_popup, play_low_beep → mdfe.console (importado acima)
# ensure_single_instance → mdfe.instance (importado acima)


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


# focused_prompt, prompt_dt_blocking, focused_alert, focused_confirm,
# prompt_batch_info, ensure_caps_off → mdfe.dialogs / mdfe.keyboard (importados acima)



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




def wait_for_form(target_text: str, tempo_maximo: float = FORM_WAIT_TIMEOUT, intervalo: float = FORM_WAIT_INTERVAL, copy_attempts: int = FORM_COPY_ATTEMPTS) -> str:
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


def verify_cte_on_page(numero_cte: str, tempo_maximo: float = VERIFY_FIELD_TIMEOUT, intervalo: float = FORM_WAIT_INTERVAL) -> None:
    """Copia o conteúdo da página e confirma a presença do CT-e informado."""
    pyautogui.press("esc")
    time.sleep(SLEEP_SHORT)
    pyautogui.hotkey("ctrl", "1")
    time.sleep(SLEEP_LONG)
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
    time.sleep(SLEEP_LONG)
    pyautogui.press("enter")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("space")
    time.sleep(SLEEP_MEDIUM)
    uf_desc = profile.get_value("mdfe", "uf_descarga")
    smart_write(uf_desc, interval=0.20)
    time.sleep(SLEEP_LONG)
    pyautogui.press("enter")
    time.sleep(SLEEP_SHORT)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(SLEEP_MEDIUM)
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
    time.sleep(SLEEP_LONG)
    
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

    skip_tabs(3)
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
        time.sleep(2.5)

    upload_latest_xml()
    time.sleep(2.5)

    # Extrair número de averbação e copiar apenas os dígitos
    pyautogui.hotkey("ctrl", "a")
    time.sleep(SLEEP_LONG)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(SLEEP_LONG)
    texto = pyperclip.paste()
    numero_averbacao = ""
    match = re.search(r"Número de Averbação:\s*([\d]+)", texto)
    if match:
        numero_averbacao = match.group(1)
        print("Número de Averbação copiado:", numero_averbacao)
    else:
        print("Número de Averbação não encontrado")

    time.sleep(SLEEP_LONG)
    pyautogui.hotkey("CTRL", "3")
    time.sleep(0.7)

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
        time.sleep(0.12)
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

    # Preparar texto com DT, CTE e NF
    combined_text = f"DT: {numero_dt if numero_dt else ''} CTE: {numero_cte if numero_cte else ''} NF: {nf_concat if nf_concat else ''}"
    log(f"Colando todas as informações de uma vez: {combined_text}")
    
    # Colar tudo de uma vez
    pyperclip.copy(combined_text)
    time.sleep(SLEEP_SHORT)
    pyautogui.hotkey("ctrl", "v")
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
        parser.add_argument("--skip-ciot", action="store_true", help="Do not prompt to run CIOT extension (used when relaunched after CIOT)")
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

	#GAP_CTRL+F
	# Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        pyautogui.press("backspace")
        time.sleep(0.15)
        smart_write("mero do DT", interval=0.12)
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
        wait_for_form("Emissor MDF-e", tempo_maximo=FORM_WAIT_TIMEOUT, intervalo=3.0, copy_attempts=FORM_COPY_ATTEMPTS)
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
        
        # Perguntar ao usuário se deseja preencher o CIOT (pular se --skip-ciot)
        if not getattr(args, "skip_ciot", False):
            log("Perguntando ao usuário sobre preenchimento de CIOT")
            ciot_buttons = ["Sim, preencher CIOT", "Não, encerrar"]
            ciot_choice = focused_confirm(
                text=(
                    "Deseja preencher o campo CIOT (Conhecimento de Transporte Intermodal Operacional) agora?\n\n"
                    "Clique em 'Sim' para abrir a extensão de preenchimento ou 'Não' para encerrar."
                ),
                title="Preencher CIOT?",
                buttons=ciot_buttons
            )

            if ciot_choice == ciot_buttons[0]:  # Sim, preencher CIOT
                log("Usuário escolheu preencher CIOT - chamando script ciot_filler.py")
                ui_print("Abrindo extensão CIOT...", style="step")
                time.sleep(SLEEP_MEDIUM)

                try:
                    import subprocess
                    ciot_script_path = BASE_DIR / "ciot_filler.py"
                    if ciot_script_path.exists():
                        log(f"Executando: {ciot_script_path}")
                        try:
                            log("Tentando focar navegador antes de iniciar CIOT")
                            focus_browser_if_needed()
                            time.sleep(SLEEP_SHORT)
                        except Exception:
                            pass
                        try:
                            hide_console_window()
                        except Exception:
                            pass

                        cmd = [str(sys.executable), str(ciot_script_path)]
                        try:
                            if os.name == "nt":
                                subprocess.Popen(cmd, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NEW_CONSOLE)
                            else:
                                subprocess.Popen(cmd, cwd=str(BASE_DIR))

                            log("CIOT iniciado em nova janela; encerrando automação principal.")
                            ui_print("CIOT iniciado em nova janela. Encerrando automação principal.", style="success")
                            # Usar os._exit para garantir encerramento imediato do processo
                            try:
                                os._exit(0)
                            except Exception:
                                raise SystemExit(0)
                        except Exception as e:
                            log(f"Erro ao iniciar script CIOT: {e}")
                            focused_alert(
                                f"Erro ao iniciar a extensão CIOT:\n{str(e)}",
                                title="Erro na extensão"
                            )
                    else:
                        log(f"Script CIOT não encontrado: {ciot_script_path}")
                        focused_alert(
                            f"O script de preenchimento de CIOT não foi encontrado:\n{ciot_script_path}",
                            title="Arquivo não encontrado"
                        )
                except Exception as e:
                    log(f"Erro ao executar script CIOT: {e}")
                    focused_alert(
                        f"Erro ao executar a extensão CIOT:\n{str(e)}",
                        title="Erro na extensão"
                    )
            else:
                log("Usuário escolheu não preencher CIOT - encerrando automação principal")
        else:
            log("--skip-ciot presente: pulando prompt de CIOT")
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
