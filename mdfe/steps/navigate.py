"""Etapa de navegação até o formulário do MDF-e."""

import time
import pyautogui

from constants import SLEEP_ONE, SLEEP_MEDIUM, SLEEP_LONGER
from mdfe.logger import log
from mdfe.keyboard import smart_write


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
