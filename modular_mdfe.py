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
    pyautogui.write(str(latest_file), interval=0.07)
    time.sleep(0.3)
    pyautogui.press("enter")


def wait_for_form(target_text: str, tempo_maximo: float = 40.0) -> None:
    intervalo_inicial = 2.0
    intervalo_maximo = 5.0
    short_sleep = 0.2
    inicio = time.monotonic()
    target_norm = re.sub(r"\s+", " ", target_text).strip().lower()
    
    # Textos alternativos para buscar
    targets = [target_text, "Emissor MDF-e", "emitente", "Dados do MDF-e"]

    log(f"Aguardando o formulário abrir (buscando '{target_text}' ou variantes)...")
    ultimo_conteudo = ""
    conteudo_anterior = ""
    confirmacoes_necessarias = 2
    confirmacoes = 0
    tentativa = 0

    while time.monotonic() - inicio < tempo_maximo:
        tentativa += 1
        intervalo_atual = min(intervalo_inicial + (tentativa * 0.5), intervalo_maximo)
        
        # Garantir que não está em nenhum campo de entrada
        pyautogui.press("esc")
        time.sleep(0.15)
        
        conteudo = ""
        
        # Limpar clipboard
        try:
            pyperclip.copy("")
            time.sleep(0.1)
        except Exception:
            pass

        # Tentar capturar conteúdo com múltiplas estratégias
        for attempt in range(4):
            pyautogui.hotkey("ctrl", "a")
            time.sleep(short_sleep)
            pyautogui.hotkey("ctrl", "c")
            time.sleep(short_sleep * 2)

            try:
                conteudo = pyperclip.paste() or ""
                if len(conteudo) > 100:
                    break
            except Exception:
                conteudo = ""
            
            time.sleep(0.15)

        ultimo_conteudo = conteudo
        conteudo_norm = re.sub(r"\s+", " ", conteudo).strip().lower()
        
        # Mostrar preview do conteúdo capturado
        preview = conteudo_norm[:150].replace("\n", " ") if conteudo_norm else "(vazio)"
        log(f"Tentativa {tentativa}: {len(conteudo)} chars | Preview: {preview}")

        # Verificar se algum dos textos alvo foi encontrado
        encontrado = False
        texto_encontrado = ""
        for target in targets:
            if re.sub(r"\s+", " ", target).strip().lower() in conteudo_norm:
                encontrado = True
                texto_encontrado = target
                break

        if encontrado:
            # Verificar se o conteúdo está estável
            if conteudo == conteudo_anterior and len(conteudo) > 200:
                confirmacoes += 1
                log(f"✓ '{texto_encontrado}' encontrado! Estável ({confirmacoes}/{confirmacoes_necessarias})")
                if confirmacoes >= confirmacoes_necessarias:
                    log("✓ Formulário detectado e estável! Continuando...")
                    time.sleep(0.5)
                    return
            else:
                confirmacoes = 0
                log(f"'{texto_encontrado}' detectado, mas conteúdo ainda mudando...")
            
            conteudo_anterior = conteudo
        else:
            confirmacoes = 0
            conteudo_anterior = ""
            log(f"Nenhum marcador encontrado. Esperando {intervalo_atual:.1f}s...")

        time.sleep(intervalo_atual)

    log(f"✗ Tempo esgotado ao buscar formulário")
    if ultimo_conteudo:
        log("Último conteúdo capturado:")
        print(ultimo_conteudo[:600])
    raise SystemExit(1)


def type_value(value: str, interval: float = 0.3) -> None:
    pyautogui.write(value, interval=interval)


def navigate_to_mdfe() -> None:
    pyautogui.hotkey("ctrl", "3")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("EMITIR NOTA", interval=0.1)
    pyautogui.press("esc")
    time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.2)
    pyautogui.write("MDF-E", interval=0.1)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.8)
    for _ in range(2):
        pyautogui.press("enter")
        time.sleep(0.5)


def fill_mdfe(profile: ConfigProfile) -> None:
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.2)
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

    pyautogui.write(profile.get_value("mdfe", "emitente_codigo", "0315-60"), interval=0.1)
    pyautogui.press("enter")
    time.sleep(0.5)
    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(0.2)

    pyautogui.write(profile.get_value("mdfe", "uf_carregamento", "SP"), interval=0.2)
    pyautogui.press("enter")
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write(profile.get_value("mdfe", "uf_descarga", "SP"), interval=0.2)
    pyautogui.press("enter")
    time.sleep(0.5)

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

    pyautogui.press("tab")
    pyautogui.press("space")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "cep_origem", "13300340"), interval=0.1)

    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "cep_destino", "13315000"), interval=0.1)
    time.sleep(1)

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

    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "unidade_medida", "1"), interval=0.1)
    pyautogui.press("enter")

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    pyautogui.press("tab")
    pyautogui.press("space")
    time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "carga_tipo", "05"), interval=0.1)
    pyautogui.press("enter")

    pyautogui.press("tab")
    time.sleep(0.05)
    pyautogui.write(profile.get_value("mdfe", "codigo_produto_descricao", "PA/PALLET"), interval=0.1)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.write(profile.get_value("mdfe", "ncm_primary", "19041000"), interval=0.1)
    pyautogui.press("enter")

    pyautogui.press("tab")
    pyautogui.press("space")
    pyautogui.press("tab")
    type_value(profile.get_value("mdfe", "cep_destino", "13315000"), 0.1)

    pyautogui.press("tab")
    pyautogui.write(profile.get_value("mdfe", "municipio_descarga", "SOROCABA"), interval=0.15)


def fill_modal_rodo(profile: ConfigProfile) -> None:
    time.sleep(1)
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("modal rodo", interval=0.1)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("esc")
    pyautogui.press("enter")
    time.sleep(1)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("RNTRC", interval=0.1)
    pyautogui.press("esc")
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "rntrc", "45501846"), interval=0.1)
    pyautogui.press("tab")
    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(0.2)

    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "contratante_nome", "PEPSICO ITU"), interval=0.2)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("modal_rodoviario", "contratante_cnpj", "02957518000224"), interval=0.12)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(1)


def fill_additional_info(profile: ConfigProfile) -> None:
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("ADICIONAIS", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("CONTRIBUINTE", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.3)
    pyautogui.write(profile.get_value("informacoes_adicionais", "contribuinte_cnpj", "04898488000177"), interval=0.1)
    pyautogui.press("tab")
    pyautogui.press("enter")

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.write("CONTRA", interval=0.1)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "prestador_adicional", "02957518000224"), interval=0.12)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_nome", "SEGUROS SURA SA"), interval=0.1)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_cnpj", "33065699000127"), interval=0.12)
    pyautogui.press("tab")
    pyautogui.write(profile.get_value("informacoes_adicionais", "terceiro_apolice", "5400035882"), interval=0.1)
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
    pyautogui.write("PEPSICO DO BRASIL", interval=0.12)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.write("02957518000224", interval=0.12)
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
    pyautogui.press("space")
    time.sleep(0.05)
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
    
    # Ir para página 3 - MDF
    pyautogui.hotkey("ctrl", "3")
    time.sleep(0.5)
    
    # Aguardar formulário
    navigate_to_mdfe()
    wait_for_form("Dados do MDF-e")
    
    # Preencher formulário
    fill_mdfe(profile)
    fill_modal_rodo(profile)
    fill_additional_info(profile)
    perform_averbacao()
    
    pyautogui.alert("Sucesso! Inclua a NF e os dados do motorista")


if __name__ == "__main__":
    main()