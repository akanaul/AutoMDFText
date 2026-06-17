"""Etapa de preenchimento dos dados do Modal Rodoviário."""

import time
import pyautogui

from constants import SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG
from mdfe.logger import log
from mdfe.keyboard import smart_write, skip_tabs
from mdfe.profile import ConfigProfile


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
