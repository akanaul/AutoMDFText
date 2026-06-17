"""Etapa de preenchimento do formulário principal do MDF-e."""

import time
import pyautogui

from constants import (
    SLEEP_ONE, SLEEP_SHORT, SLEEP_MEDIUM, SLEEP_LONG,
    SLEEP_LONGER, SLEEP_ONE_HALF
)
from mdfe.logger import log
from mdfe.keyboard import smart_write, skip_tabs, upload_latest_xml
from mdfe.profile import ConfigProfile


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
