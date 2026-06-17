"""Módulo de detecção e interação com o navegador.

Contém funções para identificar janelas do navegador (Chrome/Edge), focar a página
de forma confiável, aguardar o formulário carregar e verificar se o CT-e está na tela.
"""

import os
import re
import time
import ctypes
import ctypes.wintypes
from pathlib import Path

import pyautogui
import pyperclip

from constants import (
    SLEEP_ONE, SLEEP_MEDIUM, SLEEP_SHORT, SLEEP_LONG,
    FORM_WAIT_TIMEOUT, FORM_WAIT_INTERVAL, FORM_COPY_ATTEMPTS,
    VERIFY_FIELD_TIMEOUT, EDGE_SEARCHBAR_HEIGHT, EDGE_CLICK_OFFSET
)
from mdfe.logger import log
from mdfe.dialogs import focused_alert
from mdfe.keyboard import smart_write, _normalize_digits


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
