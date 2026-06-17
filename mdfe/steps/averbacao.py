"""Etapa de averbação e preenchimento de informações do contribuinte."""

import time
import re
import pyautogui
import pyperclip

from constants import SLEEP_LONG, SLEEP_SHORT, SLEEP_MEDIUM
from mdfe.logger import log
from mdfe.keyboard import smart_write, upload_latest_xml


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
