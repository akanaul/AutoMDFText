"""Etapa de preenchimento de Informações Adicionais/Opcionais do MDF-e."""

import time
import pyautogui

from constants import (
    SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG, SLEEP_ONE,
    CTRL_F_DELAY, DROPDOWN_SETTLE_DELAY
)
from mdfe.logger import log
from mdfe.keyboard import smart_write, skip_tabs
from mdfe.profile import ConfigProfile
from mdfe.dialogs import focused_alert


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
    write_additional(frete_val3, interval=0.10)
    time.sleep(0.15)
    pyautogui.press("tab")
    time.sleep(SLEEP_SHORT)
    pyautogui.press("enter")
    time.sleep(SLEEP_ONE)
    
    log(f"Informações Adicionais: Seguradora={seguradora}, Frete={frete_val}, Banco={numero_banco}, Parcelas={numero_parcelas}")
