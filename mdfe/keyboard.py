"""
Utilitários de teclado e clipboard para AutoMDFText.

Responsabilidades:
- Escolha inteligente entre colar (clipboard) e digitar (smart_write)
- Cola com verificação e retry (paste_text)
- Navegação por Tab com delay consistente (press_tab, skip_tabs)
- Garantir Caps Lock desligado antes do preenchimento (ensure_caps_off)
- Upload do XML mais recente da pasta Downloads (upload_latest_xml)

Estado compartilhado:
    _last_write_value e _last_write_verify são lidos por mdfe.pause
    (_verify_last_write_before_pause) para revalidar o último campo antes da pausa.
"""

import ctypes
import os
import re
import time
from pathlib import Path

import pyautogui
import pyperclip

from constants import TAB_DELAY, MIN_PASTE_LENGTH
from mdfe.logger import log
from mdfe.pause import pause_point

# ── Estado do último smart_write (lido por mdfe.pause) ───────────────────────
_last_write_value: str | None = None
_last_write_verify: bool      = False


# ── Helpers de texto ─────────────────────────────────────────────────────────

def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _normalize_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


# ── Clipboard e escrita ───────────────────────────────────────────────────────

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
    min_paste_len: int = MIN_PASTE_LENGTH,
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
    _last_write_value  = text
    _last_write_verify = verify_effective
    if use_paste:
        paste_text(text, verify=verify_effective)
    else:
        pyautogui.write(text, interval=interval)
    pause_point()


# ── Navegação por Tab ─────────────────────────────────────────────────────────

def press_tab(count: int = 1, delay: float = TAB_DELAY) -> None:
    """Pressiona Tab com delay consistente entre navegacoes."""
    for _ in range(count):
        pyautogui.press("tab")
        time.sleep(delay)


def skip_tabs(count: int, log_msg: str = "") -> None:
    """Pula N campos (tabs) com log opcional."""
    if log_msg:
        log(log_msg)
    press_tab(count=count, delay=TAB_DELAY)


# ── Utilitários de sistema ────────────────────────────────────────────────────

def ensure_caps_off() -> None:
    """Desativa Caps Lock se estiver ligado."""
    VK_CAPITAL = 0x14
    caps_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)
    if caps_state & 1:
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)


def upload_latest_xml() -> None:
    """Seleciona o arquivo mais recente em Downloads e confirma o upload."""
    from constants import SLEEP_MEDIUM
    time.sleep(SLEEP_MEDIUM)
    downloads_path = Path.home() / "Downloads"
    list_of_files  = list(downloads_path.glob("*"))
    if not list_of_files:
        from mdfe.dialogs import focused_alert
        focused_alert("A pasta Downloads está vazia!")
        return
    latest_file = max(list_of_files, key=os.path.getctime)
    smart_write(str(latest_file), interval=0.12)
    time.sleep(SLEEP_MEDIUM)
    pyautogui.press("enter")
