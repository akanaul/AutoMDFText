import argparse
import ctypes
import os
import re
import time
from pathlib import Path

import pyautogui
import pyperclip


pyautogui.FAILSAFE = True

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "logs"
CONFIG_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
DEFAULT_PROFILE = "template_config.txt"
# Log por sessão (timestamp) para facilitar debug
SESSION_TS = time.strftime("%Y%m%d_%H%M%S", time.localtime())
LOG_FILE = LOG_DIR / f"automation_{SESSION_TS}.log"

# Handle for single-instance mutex to keep it alive during process lifetime
_SINGLETON_MUTEX_HANDLE = None


def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Não interromper fluxo por falha de log em disco
        pass


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

    def get_value(self, section: str, key: str, default: str = "") -> str:
        self.ensure_current()
        return self._data.get(section.upper(), {}).get(key.lower(), default)


def list_profiles() -> list[str]:
    return sorted(p.name for p in CONFIG_DIR.glob("*.txt"))


def hide_console_window() -> None:
    if os.name != "nt":
        return
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if not hwnd:
        return
    # Apenas oculta o console para evitar encerramento do processo durante o debug
    ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE


def ensure_single_instance(name: str = "Global\\AutoMDFText_Mutex", on_duplicate: str = "warn") -> None:
    """Impede execução duplicada usando um Mutex nomeado do Windows.
    Se já existir outra instância:
    - on_duplicate == 'warn': exibe alerta e encerra este processo
    - on_duplicate == 'kill': encerra este processo (equivale a matar o duplicado)
    """
    if os.name != "nt":
        return
    kernel32 = ctypes.windll.kernel32
    # CreateMutexW(lpMutexAttributes, bInitialOwner, lpName)
    handle = kernel32.CreateMutexW(None, False, name)
    last_error = kernel32.GetLastError()
    if last_error == 183:  # ERROR_ALREADY_EXISTS
        try:
            if on_duplicate == "warn":
                pyautogui.alert("Já existe uma instância em execução. O processo será encerrado.")
        except Exception:
            pass
        raise SystemExit(0)
    else:
        # Manter handle vivo para não liberar o mutex
        global _SINGLETON_MUTEX_HANDLE
        _SINGLETON_MUTEX_HANDLE = handle


def choose_profile(interactive_list: list[str], default: str) -> str:
    log("Iniciando seleção de perfil")
    if not interactive_list:
        log("Nenhum script disponível; usando o template padrão.")
        return default
    print("Selecione o script a utilizar:")
    for idx, name in enumerate(interactive_list, start=1):
        print(f"  {idx}. {name}")
    print("Enter para usar o template padrão (ou digite o número)")
    choice = input("Opção: ").strip()
    
    selected_profile = default
    if choice:
        if choice.isdigit():
            index = int(choice) - 1
            if 0 <= index < len(interactive_list):
                selected_profile = interactive_list[index]
        elif choice in interactive_list:
            selected_profile = choice
        else:
            log("Opção inválida; usando o template padrão.")
    
    # Fechar/ocultar o terminal após seleção (com fallback em caso de WM_CLOSE falhar)
    hide_console_window()
    
    return selected_profile


def ensure_caps_off() -> None:
    VK_CAPITAL = 0x14
    caps_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)
    if caps_state & 1:
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)


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


def focus_browser_if_needed() -> None:
    """Só pressiona Win+1 se o navegador não estiver em foco, evitando minimizar."""
    title = _get_foreground_title().lower()
    cls = _get_foreground_class().lower()
    title_hits = ("chrome", "edge", "navegador", "invoisys", "google chrome", "microsoft edge")
    class_hits = ("chrome_widgetwin_1", "applicationframewindow", "windows.ui.core.corewindow")

    if any(k in title for k in title_hits) or (cls in class_hits):
        log("Navegador já em foco; Win+1 ignorado para evitar minimizar.")
        return

    log("Navegador fora de foco; tentando Win+1.")
    pyautogui.hotkey("winleft", "1")
    time.sleep(1)

    # Pós-checagem: se ainda não estiver em foco, tentar fallback suave
    title2 = _get_foreground_title().lower()
    cls2 = _get_foreground_class().lower()
    if any(k in title2 for k in title_hits) or (cls2 in class_hits):
        log("Navegador em foco após Win+1.")
        return
    log("Win+1 não focou o navegador; evitando minimizar e mantendo estado.")


def extract_cte_from_content(conteudo: str) -> str:
    """Extrai o número do CT-e do conteúdo capturado."""
    linhas = conteudo.splitlines()
    for linha in linhas:
        # Padrão: "100 - Autorizado o uso do CT-e.N" seguido de número com 6 dígitos
        if "100 - Autorizado o uso do CT-e" in linha:
            # Extrai 6 dígitos consecutivos após "CT-e"
            match = re.search(r'CT-e\.?[^\d]*(\d{6})', linha)
            if match:
                return match.group(1)
    return ""


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





def wait_for_form(target_text: str, tempo_maximo: float = 15.0, intervalo: float = 1.0, copy_attempts: int = 2) -> str:
    """Aguarda o formulário abrir detectando texto específico (lógica do legado).
    Retorna o conteúdo do clipboard capturado quando o formulário é encontrado."""
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


def navigate_to_mdfe() -> None:
    """Navega para o formulário MDF-e - cópia exata do script legado"""
    # IR PARA 3ª PAGINA - MDF
    pyautogui.hotkey("ctrl", "3")
    time.sleep(1)

    # ABRIR DADOS DO MDF-E
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'EMITIR NOTA'")
    pyautogui.write("EMITIR NOTA", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)
    
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'MDF-E'")
    pyautogui.write("MDF-E", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.5)


def fill_mdfe(profile: ConfigProfile) -> None:
    """Preenche dados do MDF-e - cópia exata do script legado"""
    time.sleep(1)
    # PRESTADOR DE SERVIÇO
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log(f"Preenchendo PRESTADOR DE SERVIÇO: {profile.get_value('mdfe', 'prestador_tipo')}")
    pyautogui.write("SELECIONE...", interval=0.20)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.write(profile.get_value("mdfe", "prestador_tipo"), interval=0.1)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.3)
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    # EMITENTE
    emitente = profile.get_value("mdfe", "emitente_codigo")
    log(f"Preenchendo EMITENTE: {emitente}")
    pyautogui.write(emitente, interval=0.10)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.3)

    # UF CARREGAMENTO E DESCARREGAMENTO
    uf_car = profile.get_value("mdfe", "uf_carregamento")
    log(f"Preenchendo UF CARREGAMENTO: {uf_car}")
    pyautogui.write(uf_car, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("space")
    time.sleep(0.3)
    uf_desc = profile.get_value("mdfe", "uf_descarga")
    log(f"Preenchendo UF DESCARGA: {uf_desc}")
    pyautogui.write(uf_desc, interval=0.20)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)

    # MUNICIPIO DE CARREGAMENTO
    pyautogui.press("tab")
    time.sleep(0.2)
    municipio = profile.get_value("mdfe", "municipio_carregamento").upper()
    log(f"Preenchendo MUNICIPIO CARREGAMENTO: {municipio}")
    pyautogui.write(municipio, interval=0.15)
    time.sleep(0.3)

    for _ in range(4):
        pyautogui.press("down")
        time.sleep(0.1)
    for _ in range(3):
        pyautogui.press("up")
        time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(0.3)

    # UPLOAD DO ARQUIVO XML
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.press("space")
    time.sleep(2)
    upload_latest_xml()
    time.sleep(0.5)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(2.5)

    # UNIDADE DE MEDIDA
    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    unidade = profile.get_value("mdfe", "unidade_medida")
    log(f"Preenchendo UNIDADE MEDIDA: {unidade}")
    pyautogui.write(unidade, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # TIPO DE CARGA
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    carga = profile.get_value("mdfe", "carga_tipo")
    log(f"Preenchendo TIPO CARGA: {carga}")
    pyautogui.write(carga, interval=0.1)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # DESCRIÇÃO DO PRODUTO
    pyautogui.press("tab")
    time.sleep(0.2)
    descricao = profile.get_value("mdfe", "codigo_produto_descricao")
    log(f"Preenchendo DESCRIÇÃO PRODUTO: {descricao}")
    pyautogui.write(descricao, interval=0.1)
    time.sleep(0.2)

    # CÓDIGO NCM
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    
    ncm_primary = profile.get_value("mdfe", "ncm_primary", "19041000")
    ncm_secondary = profile.get_value("mdfe", "ncm_secondary", "19059090")
    ncm_tertiary = profile.get_value("mdfe", "ncm_tertiary", "20052000")
    
    opcao = pyautogui.confirm(
        text='Selecione o código NCM ou escolha "Outro código" para digitar manualmente:',
        title='Escolha de NCM',
        buttons=[ncm_primary, ncm_secondary, ncm_tertiary, 'Outro código', 'Cancelar']
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
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.3)

    # CEP ORIGEM
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    cep_orig = profile.get_value("mdfe", "cep_origem")
    log(f"Preenchendo CEP ORIGEM: {cep_orig}")
    pyautogui.write(cep_orig, interval=0.1)
    time.sleep(0.3)

    # CEP DESTINO
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.15)
    cep_dest = profile.get_value("mdfe", "cep_destino")
    log(f"Preenchendo CEP DESTINO: {cep_dest}")
    pyautogui.write(cep_dest, interval=0.1)
    time.sleep(1.5)


def fill_modal_rodo(profile: ConfigProfile) -> None:
    # ABRIR DADOS DO MODAL RODOVIÁRIO
    time.sleep(1.5)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    log("Procurando por 'modal rodo'")
    pyautogui.write("modal rodo", interval=0.10)
    time.sleep(0.2)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.5)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.5)

    # RNTRC
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.3)
    pyautogui.write("RNTRC", interval=0.10)
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.2)
    rntrc = profile.get_value("modal_rodoviario", "rntrc")
    log(f"Preenchendo RNTRC: {rntrc}")
    pyautogui.write(rntrc, interval=0.10)
    time.sleep(0.2)
    
    # NOME DO CONTRATANTE
    for _ in range(6):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("space")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    contratante = profile.get_value("modal_rodoviario", "contratante_nome")
    log(f"Preenchendo CONTRATANTE: {contratante}")
    pyautogui.write(contratante, interval=0.20)
    time.sleep(0.2)

    # CNPJ DO CONTRATATANTE
    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE: {cnpj_cont}")
    pyautogui.write(cnpj_cont, interval=0.12)
    time.sleep(0.2)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.15)
    pyautogui.press("enter")
    time.sleep(1.5)


def fill_additional_info(profile: ConfigProfile) -> None:
    # ABRIR DADOS DE INFORMAÇÕES OPCIONAIS
    pyautogui.hotkey("ctrl", "f")
    log("Procurando por 'OPCIONAIS'")
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
    log("Procurando por 'ADICIONAIS'")
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
    cnpj_contrib = profile.get_value("informacoes_adicionais", "contribuinte_cnpj")
    log(f"Preenchendo CONTRIBUINTE CNPJ: {cnpj_contrib}")
    pyautogui.write(cnpj_contrib, interval=0.10)
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
    cnpj_cont2 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE (CONTRA): {cnpj_cont2}")
    pyautogui.write(cnpj_cont2, interval=0.12)
    pyautogui.press("tab")
    seguradora = profile.get_value("informacoes_adicionais", "seguradora_nome")
    log(f"Preenchendo SEGURADORA: {seguradora}")
    pyautogui.write(seguradora, interval=0.10)
    pyautogui.press("tab")
    cnpj_seg = profile.get_value("informacoes_adicionais", "seguradora_cnpj")
    log(f"Preenchendo CNPJ SEGURADORA: {cnpj_seg}")
    pyautogui.write(cnpj_seg, interval=0.12)
    pyautogui.press("tab")
    apolice = profile.get_value("informacoes_adicionais", "numero_apolice")
    log(f"Preenchendo NUMERO APOLICE: {apolice}")
    pyautogui.write(apolice, interval=0.10)
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
    contratante2 = profile.get_value("modal_rodoviario", "contratante_nome")
    log(f"Preenchendo CONTRATANTE (2): {contratante2}")
    pyautogui.write(contratante2, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    cnpj_cont3 = profile.get_value("modal_rodoviario", "contratante_cnpj")
    log(f"Preenchendo CNPJ CONTRATANTE (2): {cnpj_cont3}")
    pyautogui.write(cnpj_cont3, interval=0.12)
    time.sleep(0.2)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.3)

    frete_val = profile.get_value("informacoes_adicionais", "frete_valor")
    log(f"Preenchendo FRETE VALOR: {frete_val}")
    pyautogui.write(frete_val, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)
    forma_pag = profile.get_value("informacoes_adicionais", "forma_pagamento")
    log(f"Preenchendo FORMA PAGAMENTO: {forma_pag}")
    pyautogui.write(forma_pag, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("enter")
    time.sleep(0.3)

    # Garantir que o foco chegue no campo "Número do banco" (antes de agência)
    pyautogui.press("tab")
    time.sleep(0.2)
    numero_banco = profile.get_value("informacoes_adicionais", "numero_banco")
    log(f"Preenchendo NUMERO BANCO: {numero_banco}")
    pyautogui.write(numero_banco, interval=0.12)
    time.sleep(0.2)

    pyautogui.press("tab")
    time.sleep(0.2)
    agencia = profile.get_value("informacoes_adicionais", "agencia")
    log(f"Preenchendo AGENCIA: {agencia}")
    pyautogui.write(agencia, interval=0.12)
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
    log("Preenchendo seção FRETE")
    pyautogui.write("FRETE", interval=0.10)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_val2 = profile.get_value("informacoes_adicionais", "frete_valor")
    log(f"Preenchendo FRETE VALOR (seção): {frete_val2}")
    pyautogui.write(frete_val2, interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    frete_tipo = profile.get_value("informacoes_adicionais", "frete_tipo")
    log(f"Preenchendo FRETE TIPO: {frete_tipo}")
    pyautogui.write(frete_tipo, interval=0.12)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)

    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("SELECIONE...", interval=0.10)
    pyautogui.press("esc")
    time.sleep(0.3)

    for _ in range(5):
        pyautogui.press("tab")
        time.sleep(0.2)
    numero_parcelas = profile.get_value("informacoes_adicionais", "numero_parcelas", "1")
    log(f"Preenchendo NUMERO PARCELAS: {numero_parcelas}")
    pyautogui.write(numero_parcelas, interval=0.10)
    time.sleep(0.15)

    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("space")
    time.sleep(0.2)

    for _ in range(7):
        pyautogui.press("tab")
        time.sleep(0.4)
    pyautogui.press("space")
    time.sleep(0.3)
    pyautogui.press("space")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(0.3)
    frete_val3 = profile.get_value("informacoes_adicionais", "frete_valor")
    log(f"Preenchendo FRETE VALOR (final): {frete_val3}")
    pyautogui.write(frete_val3, interval=0.10)
    time.sleep(0.15)
    pyautogui.press("tab")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.5)


def perform_averbacao(numero_cte: str = "", numero_dt: str = "") -> None:
    """Executa a averbação e preenche DT/CT-e no final.
    - Usa CT-e capturado no início quando disponível, sem voltar à primeira página
    - Em fallback, captura CT-e na INVOISYS sem atrelar ou sobrescrever o DT
    """
    # Abrir site/aba de averbação e enviar XML
    pyautogui.hotkey("ctrl", "shift", "a")
    time.sleep(0.5)
    pyautogui.write("ATM", interval=0.1)
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1)

    for search in ("OK", "XML", "ENVIAR"):
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        pyautogui.write(search, interval=0.1)
        time.sleep(0.5)
        pyautogui.press("esc")
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.5)

    upload_latest_xml()
    time.sleep(2)

    # Extrair número de averbação e copiar apenas os dígitos
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.2)
    texto = pyperclip.paste()
    match = re.search(r"Número de Averbação:\s*([\d]+)", texto)
    if match:
        numero_averbacao = match.group(1)
        pyperclip.copy(numero_averbacao)
        print("Número de Averbação copiado:", numero_averbacao)
    else:
        print("Número de Averbação não encontrado")

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

    # Preencher DT/CT-e/NF na área de CONTRIBUINTE
    pyautogui.hotkey("ctrl", "f")
    time.sleep(0.5)
    pyautogui.write("CONTRIBUINTE", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.5)

    # DT sempre vem do prompt (numero_dt) capturado no início
    pyautogui.write("DT: ", interval=0.1)
    time.sleep(0.3)
    if numero_dt:
        pyperclip.copy(numero_dt)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    # CT-e: se já temos do início, usar direto; caso contrário, fallback
    pyautogui.write(" CTE: ", interval=0.1)
    time.sleep(0.5)
    if numero_cte:
        log(f"Usando CT-e capturado no início: {numero_cte}")
        pyperclip.copy(numero_cte)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
    else:
        log("CT-e ausente; iniciando fallback na INVOISYS para capturar.")
        pyautogui.hotkey("ctrl", "shift", "a")
        time.sleep(0.5)
        pyautogui.write("INVOISYS", interval=0.1)
        for _ in range(2):
            pyautogui.press("tab")
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1)

        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        pyautogui.write("e final", interval=0.2)
        pyautogui.press("esc")
        time.sleep(0.2)
        for _ in range(2):
            pyautogui.press("tab")
            time.sleep(0.1)
        time.sleep(0.5)

        # Copiar página e extrair CT-e
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)
        conteudo = pyperclip.paste()
        linhas = conteudo.splitlines()
        numero_cte_local = ""
        for linha in linhas:
            if "100 - Autorizado o uso do CT-e.N" in linha:
                m = re.search(r"100\s*-\s*Autorizado o uso do CT-e\.N[^\d]*(\d{6})", linha)
                if m:
                    numero_cte_local = m.group(1)
                    break

        # Voltar para a aba de preenchimento e colar somente CT-e se encontrado
        pyautogui.hotkey("alt", "tab")
        time.sleep(0.5)
        if numero_cte_local:
            log(f"Número do CT-e encontrado (fallback): {numero_cte_local}")
            pyperclip.copy(numero_cte_local)
            pyautogui.hotkey("ctrl", "v")
        else:
            log("Não foi possível localizar o número do CT-e (fallback); deixar em branco.")
        time.sleep(0.5)

    # Finalizar com rótulo NF
    pyautogui.write(" NF: ", interval=0.1)
    time.sleep(0.5)


def main() -> None:
    try:
        # Sleep inicial para aguardar inicialização
        time.sleep(0.7)
        log("Iniciando automação (main)")
        # Impedir execução duplicada
        ensure_single_instance()
        
        parser = argparse.ArgumentParser()
        parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
        args = parser.parse_args()

        selected = args.profile
        if not selected:
            selected = choose_profile(list_profiles(), DEFAULT_PROFILE)
        log(f"Perfil selecionado: {selected}")
        profile_path = CONFIG_DIR / selected
        if not profile_path.exists():
            log(f"Perfil {profile_path} não encontrado, usando template")
            profile_path = CONFIG_DIR / DEFAULT_PROFILE

        profile = ConfigProfile(profile_path)
        log(f"Perfil carregado de: {profile_path}")

        # Alerta ANTES de abrir navegador (herdado do legado)
        log("Exibindo alerta inicial")
        pyautogui.alert(
            'ANTES DE PROSSEGUIR:\n\n'
            '1. Mantenha 3 abas do Invoisys abertas e o site de averbação logado;\n'
            '2. Deixe o navegador como o primeiro app na barra do Windows;\n'
            '3. Mantenha apenas uma janela do navegador ativa.\n\n'
            'OBS: Para interromper o código, mova o mouse repetidamente para o canto superior direito da tela.'
        )
        time.sleep(1)
        
        # Abrir/focar navegador sem minimizar (usa Win+1 só se não estiver em foco)
        log("Focando navegador (evitando minimizar)")
        focus_browser_if_needed()
        
        # GAP - Pressionar ESC 2x
        log("Enviando ESC x2")
        for _ in range(2):
            pyautogui.press("esc")
            time.sleep(0.3)

        # Recarregar aba 3
        log("Recarregando aba 3 (Ctrl+3, F5)")
        pyautogui.hotkey("ctrl", "3")
        time.sleep(0.5)
        pyautogui.press("f5")
        time.sleep(1)

        # Voltar para aba 1 (uma vez, como no legado)
        log("Voltando para aba 1 e focando em 'empresa'")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(0.5)
        # Buscar "empresa" para garantir foco correto
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.write("empresa", interval=0.10)
        pyautogui.press("esc")
        time.sleep(0.5)

        # Detectar primeira tela (CT-e) como no legado
        log("Detectando tela CT-e (notas emitidas: ct-e)")
        conteudo_cte_pagina = wait_for_form("notas emitidas: ct-e", tempo_maximo=4, intervalo=1, copy_attempts=2)
        log("Página CT-e detectada. Extraindo CT-e do conteúdo capturado...")
        numero_cte = extract_cte_from_content(conteudo_cte_pagina)
        if numero_cte:
            log(f"CT-e extraído com sucesso no início: {numero_cte}")
            pyperclip.copy(numero_cte)
        else:
            log("Aviso: Não foi possível extrair CT-e do conteúdo inicial. Tentaremos novamente ao final.")

        # Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(2)
        pyautogui.write("serie final", interval=0.12)
        pyautogui.press("esc")
        time.sleep(0.5)
        for _ in range(2):
            pyautogui.press("tab")
            time.sleep(1)
        # Prompt para DT
        prompt_text = profile.get_value("general", "dt_prompt_text", "Digite o número do DT:")
        codigo = pyautogui.prompt(text=prompt_text, title="DT")
        if not codigo:
            pyautogui.alert("Nenhum código informado. O script foi pausado.")
            pyautogui.FAILSAFE = True
            return
        
        # Escrever código DT e dar enter 2x
        pyautogui.write(codigo.upper(), interval=0.1)
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(0.5)

        # Alerta com instruções (do perfil)
        pyautogui.alert(profile.get_value("general", "alert_intro", "Antes de prosseguir:\n\n1. Baixe o arquivo XML;\n2. Mantenha 3 abas do Invoisys abertas no começo do navegador;\n3. Mantenha o site de averbação logado.\n\nOBS: Para interromper o processo, deslize o mouse repetidamente em direção ao canto superior direito da tela"))
        time.sleep(1)

        # Desativar Caps Lock
        ensure_caps_off()
        
        # Navegar para MDF-e e detectar formulário (lógica e tempos do legado)
        navigate_to_mdfe()
        wait_for_form("Emissor MDF-e", tempo_maximo=15.0, intervalo=3.0, copy_attempts=3)
        
        # Preencher formulário
        fill_mdfe(profile)
        fill_modal_rodo(profile)
        fill_additional_info(profile)
        perform_averbacao(numero_cte, codigo)
        
        pyautogui.alert("Sucesso! Inclua a NF e os dados do motorista")
    except Exception as e:
        log(f"ERRO FATAL: {e}")
        raise


if __name__ == "__main__":
    main()