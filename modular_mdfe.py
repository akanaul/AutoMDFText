import argparse
import ctypes
import os
import re
import time
from pathlib import Path

import pyautogui
import pyperclip


pyautogui.FAILSAFE = True

CONFIG_DIR = Path(__file__).parent / "scripts"
CONFIG_DIR.mkdir(exist_ok=True)
DEFAULT_PROFILE = "template_config.txt"


def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    print(f"[{ts}] {msg}")


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

    def get_value(self, section: str, key: str, default: str) -> str:
        self.ensure_current()
        return self._data.get(section.upper(), {}).get(key.lower(), default)


def list_profiles() -> list[str]:
    return sorted(p.name for p in CONFIG_DIR.glob("*.txt"))


def choose_profile(interactive_list: list[str], default: str) -> str:
    if not interactive_list:
        log("Nenhum script disponível; usando o template padrão.")
        return default
    print("Selecione o script a utilizar:")
    for idx, name in enumerate(interactive_list, start=1):
        print(f"  {idx}. {name}")
    print("Enter para usar o template padrão (ou digite o número)")
    choice = input("Opção: ").strip()
    if not choice:
        return default
    if choice.isdigit():
        index = int(choice) - 1
        if 0 <= index < len(interactive_list):
            return interactive_list[index]
    if choice in interactive_list:
        return choice
    log("Opção inválida; usando o template padrão.")
    return default


def ensure_caps_off() -> None:
    VK_CAPITAL = 0x14
    caps_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)
    if caps_state & 1:
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)


def upload_latest_xml() -> None:
    downloads_path = Path.home() / "Downloads"
    list_of_files = list(downloads_path.glob("*"))
    if not list_of_files:
        pyautogui.alert("A pasta Downloads está vazia!")
        return
    latest_file = max(list_of_files, key=os.path.getctime)
    pyautogui.write(str(latest_file), interval=0.05)
    time.sleep(0.3)
    pyautogui.press("enter")


def wait_for_form(target_text: str, tempo_maximo: float = 15.0) -> None:
    """Aguarda o formulário abrir detectando texto específico - cópia do script legado"""
    intervalo = 3.0           # segundos entre tentativas
    copy_attempts = 3         # pequenas tentativas rápidas de Ctrl+C por iteração
    short_sleep = 0.12        # pausa curta entre ações de cópia
    
    inicio = time.monotonic()
    ultimo_conteudo = ""
    target_norm = re.sub(r"\s+", " ", target_text).strip().lower()

    log("Aguardando o formulário abrir...")
    attempt = 0

    while time.monotonic() - inicio < tempo_maximo:
        attempt += 1
        try:
            # limpa a área de transferência para evitar ler conteúdo antigo
            try:
                pyperclip.copy("")
            except Exception as e:
                log(f"Aviso: não foi possível limpar clipboard: {e}")

            # realiza várias tentativas rápidas de selecionar/copiar
            for i in range(copy_attempts):
                pyautogui.hotkey("ctrl", "a")   # selecionar tudo
                time.sleep(short_sleep)
                pyautogui.hotkey("ctrl", "c")   # copiar
                time.sleep(short_sleep)

            # lê o conteúdo
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
                log("Continuando com o restante da automação...")
                time.sleep(0.8)
                return

            log(f"Não encontrado. Aguardando {intervalo}s antes da próxima tentativa.")
            time.sleep(intervalo)

        except Exception as e:
            log(f"Erro interno durante a tentativa: {e}")
            time.sleep(intervalo)

    # tempo esgotado
    log(f"Formulário não foi detectado dentro de {tempo_maximo} segundos. Encerrando o processo.")
    if ultimo_conteudo:
        log("Último conteúdo capturado (preview):")
        print(ultimo_conteudo[:400])
    raise SystemExit(1)


def type_value(value: str, interval: float = 0.3) -> None:
    pyautogui.write(value, interval=interval)


def navigate_to_mdfe() -> None:
    """Navega para o formulário MDF-e - cópia exata do script legado"""
    # IR PARA 3ª PAGINA - MDF
    pyautogui.hotkey("ctrl", "3")
    time.sleep(0.5)

    # ABRIR DADOS DO MDF-E
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("EMITIR NOTA", interval=0.10)
    pyautogui.press("esc")
    pyautogui.press("enter")
    time.sleep(0.5)
    
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("MDF-E", interval=0.10)
    pyautogui.press("esc")
    pyautogui.press("enter")
    time.sleep(0.8)


def fill_mdfe(profile: ConfigProfile) -> None:
    """Preenche dados do MDF-e - cópia exata do script legado"""
    # PRESTADOR DE SERVIÇO
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.20)
    pyautogui.press("esc")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("mdfe", "prestador_tipo", "PRESTADOR"), interval=0.1)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    # EMITENTE
    pyautogui.write(profile.get_value("mdfe", "emitente_codigo", "0315-60"), interval=0.10)
    pyautogui.press("enter")
    time.sleep(0.5)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(0.2)

    # UF CARREGAMENTO E DESCARREGAMENTO
    pyautogui.write(profile.get_value("mdfe", "uf_carregamento", "SP"), interval=0.20)
    pyautogui.press("enter")
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("mdfe", "uf_descarga", "SP"), interval=0.20)
    pyautogui.press("enter")
    time.sleep(0.5)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "municipio_carregamento", "ITU").upper(), interval=0.15)
    time.sleep(0.3)

    for _ in range(4):
        pyautogui.press("down")
        time.sleep(0.1)
    for _ in range(3):
        pyautogui.press("up")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(0.2)

    # UPLOAD DO ARQUIVO XML
    for _ in range(2):
        pyautogui.press("tab")
    pyautogui.press("space")
    time.sleep(2)
    upload_latest_xml()
    time.sleep(0.3)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(2)

    # UNIDADE DE MEDIDA
    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "unidade_medida", "1"), interval=0.1)
    pyautogui.press("enter")

    # TIPO DE CARGA
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    pyautogui.press("tab")
    pyautogui.press("space")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "carga_tipo", "05"), interval=0.1)
    pyautogui.press("enter")

    # DESCRIÇÃO DO PRODUTO
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("mdfe", "produto_descricao", "PA/PALLET"), interval=0.1)

    # CÓDIGO NCM
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    
    opcao = pyautogui.confirm(
        text='Selecione o código NCM ou escolha "Outro código" para digitar manualmente:',
        title='Escolha de NCM',
        buttons=['19041000', '19059090', '20052000', 'Outro código', 'Cancelar']
    )

    if opcao == 'Cancelar':
        pyautogui.alert('Nenhum código NCM selecionado. O script foi pausado.')
        pyautogui.FAILSAFE = True
        raise SystemExit(1)
    elif opcao == 'Outro código':
        codigo = pyautogui.prompt('Digite o código NCM:')
        if not codigo:
            pyautogui.alert('Nenhum código NCM digitado. O script foi pausado.')
            pyautogui.FAILSAFE = True
            raise SystemExit(1)
    else:
        codigo = opcao
    
    pyautogui.write(codigo.upper(), interval=0.1)
    pyautogui.press("enter")

    # CEP ORIGEM
    pyautogui.press("tab")
    pyautogui.press("space")
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("mdfe", "cep_origem", "13300340"), interval=0.1)

    # CEP DESTINO
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "cep_destino", "13315000"), interval=0.1)
    time.sleep(1)


def fill_modal_rodo(profile: ConfigProfile) -> None:
    # ABRIR DADOS DO MODAL RODOVIÁRIO
    time.sleep(1)
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("modal rodo", interval=0.10)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("esc")
    pyautogui.press("enter")
    time.sleep(1)

    # RNTRC
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("RNTRC", interval=0.10)
    pyautogui.press("esc")
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "rntrc", "45501846"), interval=0.10)
    
    # NOME DO CONTRATANTE
    for _ in range(6):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "contratante_nome", "PEPSICO ITU"), interval=0.20)
    time.sleep(0.1)

    # CNPJ DO CONTRATATANTE
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "contratante_cnpj", "02957518000224"), interval=0.12)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(1)


def fill_additional_info(profile: ConfigProfile) -> None:
    # ABRIR DADOS DE INFORMAÇÕES OPCIONAIS
    pyautogui.hotkey("ctrl", "f")
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
    pyautogui.write(profile.get_value("informacoes_adicionais", "contribuinte_cnpj", "04898488000177"), interval=0.10)
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
    pyautogui.write(profile.get_value("informacoes_adicionais", "prestador_adicional", "02957518000224"), interval=0.12)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_nome", "SEGUROS SURA SA"), interval=0.10)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_cnpj", "33065699000127"), interval=0.12)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_apolice", "5400035882"), interval=0.10)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(0.3)

    for _ in range(4):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.press("space")
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "retentor_nome", "PEPSICO DO BRASIL"), interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "retentor_cnpj", "02957518000224"), interval=0.12)
    time.sleep(0.2)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)

    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_valor", "1314.27"), interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "quantidade_unidade", "1"), interval=0.12)
    time.sleep(0.2)

    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "serie", "237"), interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "numero_pedido", "2372/8"), interval=0.12)
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
    pyautogui.press("enter")
    pyautogui.write("FRETE", interval=0.10)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_valor", "1314.27"), interval=0.12)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_tipo", "FRETE"), interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.10)

    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.05)
    pyautogui.write("1", interval=0.10)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.05)
    pyautogui.press("space")
    time.sleep(0.10)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.05)
    for _ in range(2):
        pyautogui.press("space")
        time.sleep(0.05)
    pyautogui.press("tab")
    time.sleep(0.05)
    pyautogui.press("enter")
    time.sleep(0.05)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_valor", "1314.27"), interval=0.10)
    time.sleep(0.05)
    pyautogui.press("tab")
    time.sleep(0.05)
    pyautogui.press("enter")
    time.sleep(1)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.1)
    pyautogui.press("esc")
    pyautogui.press("enter")
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_identificador", "FRETE"), interval=0.1)
    pyautogui.press("enter")
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_valor", "1314.27"), interval=0.12)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("informacoes_adicionais", "frete_identificador", "FRETE"), interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)

    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.05)
    pyautogui.write(profile.get_value("informacoes_adicionais", "quantidade_unidade", "1"), interval=0.1)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.05)
    pyautogui.press("space")
    time.sleep(0.1)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.05)
    for _ in range(2):
        pyautogui.press("space")
        time.sleep(0.05)
    pyautogui.press("tab")
    time.sleep(0.05)
    pyautogui.press("enter")
    time.sleep(0.05)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "adicional_valor", "1321.02"), interval=0.1)
    time.sleep(0.05)
    pyautogui.press("tab")
    time.sleep(0.05)
    pyautogui.press("enter")
    time.sleep(1)


def perform_averbacao() -> None:
    pyautogui.hotkey("ctrl", "shift", "a")
    time.sleep(0.5)
    pyautogui.write("ATM", interval=0.1)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1)

    # Buscar por OK, XML e ENVIAR
    for search in ("OK", "XML", "ENVIAR"):
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        pyautogui.write(search, interval=0.1)
        time.sleep(0.5)
        pyautogui.press("esc")
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.5)

    # Upload do arquivo XML
    upload_latest_xml()
    time.sleep(0.5)

    # Extrair e copiar número de averbação
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.5)

    conteudo = pyperclip.paste()
    linhas = conteudo.splitlines()
    numero_averbacao = None
    for linha in linhas:
        if "Número de Averbação:" in linha:
            match = re.search(r"Número de Averbação:\s*([\d]+)", linha)
            if match:
                numero_averbacao = match.group(1)
                break

    if numero_averbacao:
        log(f"Número de Averbação encontrado: {numero_averbacao}")
        pyperclip.copy(numero_averbacao)
    else:
        log("Não foi possível localizar o número de Averbação")

    time.sleep(0.5)
    pyautogui.hotkey("alt", "tab")
    time.sleep(1)

    # Preencher detalhes na outra aba
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    pyautogui.write("DETALHES", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    pyautogui.press("enter")
    time.sleep(0.5)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.5)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(1)

    # Buscar INVOISYS para coletar DT e CTE
    pyautogui.hotkey("ctrl", "shift", "a")
    time.sleep(0.5)
    pyautogui.write("INVOISYS", interval=0.1)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(1)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("e final", interval=0.2)
    pyautogui.press("esc")
    time.sleep(0.2)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)

    time.sleep(0.5)

    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.5)
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    pyautogui.write("CONTRIBUINTE", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.write("DT: ", interval=0.1)
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)
    pyautogui.write(" CTE: ", interval=0.1)
    time.sleep(0.5)
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.5)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)

    conteudo = pyperclip.paste()
    linhas = conteudo.splitlines()
    numero_cte = None
    for linha in linhas:
        if "100 - Autorizado o uso do CT-e.N" in linha:
            match = re.search(r"100\s*-\s*Autorizado o uso do CT-e\.N[^\d]*(\d{6})", linha)
            if match:
                numero_cte = match.group(1)
                break

    if numero_cte:
        log(f"Número do CT-e encontrado: {numero_cte}")
        pyperclip.copy(numero_cte)
    else:
        log("Não foi possível localizar o número do CT-e")

    pyautogui.hotkey("alt", "tab")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)
    pyautogui.write(" NF: ", interval=0.1)
    time.sleep(0.5)


def main() -> None:
    # Sleep inicial para aguardar inicialização
    time.sleep(0.7)
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
    args = parser.parse_args()

    selected = args.profile
    if not selected:
        selected = choose_profile(list_profiles(), DEFAULT_PROFILE)
    profile_path = CONFIG_DIR / selected
    if not profile_path.exists():
        log(f"Perfil {profile_path} não encontrado, usando template")
        profile_path = CONFIG_DIR / DEFAULT_PROFILE

    profile = ConfigProfile(profile_path)
    
    # Abrir navegador
    pyautogui.hotkey("winleft", "1")
    time.sleep(1)
    
    # GAP - Pressionar ESC 2x
    for _ in range(2):
        pyautogui.press("esc")
        time.sleep(0.3)

    # Recarregar aba 3
    pyautogui.hotkey("ctrl", "3")
    time.sleep(0.5)
    pyautogui.press("f5")
    time.sleep(1)

    # Voltar para aba 1 (2x como no legado)
    pyautogui.hotkey("ctrl", "1")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "1")
    time.sleep(0.5)

    # Pesquisar e posicionar em "e final"
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.2)
    type_value("e final", 0.1)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.2)

    # Tab 2x como no legado
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)

    time.sleep(0.5)

    # Prompt para DT
    prompt_text = profile.get_value("general", "dt_prompt_text", "Digite o número do DT:")
    codigo = pyautogui.prompt(text=prompt_text, title="DT")
    if not codigo:
        pyautogui.alert("Nenhum código informado. O script foi pausado.")
        pyautogui.FAILSAFE = True
        return
    
    # Escrever código DT e dar enter 2x como no legado
    pyautogui.write(codigo.upper(), interval=0.1)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)

    # Alerta com instruções
    pyautogui.alert(profile.get_value("general", "alert_intro", "Antes de prosseguir:\n\n1. Baixe o arquivo XML;\n2. Mantenha 3 abas do Invoisys abertas no começo do navegador;\n2. Mantenha o site de averbação logado.\n\nOBS: Para interromper o processo, deslize o mouse repetidamente em direção ao canto superior direito da tela"))
    
    time.sleep(1)

    # Desativar Caps Lock
    ensure_caps_off()
    
    # Aguardar formulário - navegação e detecção do MDF-e
    navigate_to_mdfe()
    wait_for_form("Emissor MDF-e")
    
    # Preencher formulário
    fill_mdfe(profile)
    fill_modal_rodo(profile)
    fill_additional_info(profile)
    perform_averbacao()
    
    pyautogui.alert("Sucesso! Inclua a NF e os dados do motorista")


if __name__ == "__main__":
    main()