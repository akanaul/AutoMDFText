"""
Leitura e gerenciamento de perfis de configuração para AutoMDFText.

Responsabilidades:
- Parsear arquivos .txt de perfil em dicionário de seções (parse_profile)
- Representar um perfil carregado com hot-reload automático (ConfigProfile)
- Listar os perfis disponíveis em scripts/ (list_profiles)

Os perfis seguem a estrutura INI com seções como [MDFE], [MODAL_RODOVIARIO] etc.
"""

from pathlib import Path

# Diretório de perfis relativo à raiz do projeto
_BASE_DIR  = Path(__file__).parent.parent
CONFIG_DIR = _BASE_DIR / "scripts"
CONFIG_DIR.mkdir(exist_ok=True)


def parse_profile(path: Path) -> dict[str, dict[str, str]]:
    """Lê um arquivo .txt de perfil e organiza em dicionário de seções."""
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
    """Perfil de configuração com detecção automática de mudança no arquivo (hot-reload)."""

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
    """Retorna os nomes de todos os arquivos .txt em scripts/, ordenados."""
    return sorted(p.name for p in CONFIG_DIR.glob("*.txt"))


def choose_profile(interactive_list: list[str]) -> str:
    import time
    from constants import SLEEP_ONE_HALF
    from mdfe.logger import log
    from mdfe.dialogs import focused_alert
    from mdfe.console import hide_console_window

    log("Iniciando seleção de perfil")
    if not interactive_list:
        log("Nenhum script disponível; diretório scripts/ está vazio.")
        try:
            focused_alert(
                "Nenhum script encontrado em scripts/.\n\n"
                "Adicione um arquivo .txt de configuração e tente novamente.",
                title="Nenhum script encontrado"
            )
        except Exception:
            pass
        raise SystemExit(1)
    
    # Cores ANSI para destacar o menu
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BOLD = "\033[1m"
    RESET = "\033[0m"
    
    selected_profile = None
    max_attempts = 100  # Proteção contra loops infinitos por erro de lógica
    attempts = 0
    
    # Loop até que uma seleção válida seja feita
    while not selected_profile and attempts < max_attempts:
        attempts += 1
        
        try:
            # Revalidar lista a cada iteração (caso arquivos sejam adicionados/removidos)
            current_list = list_profiles()
            if not current_list:
                log("Lista de scripts ficou vazia durante seleção; sem perfil padrão disponível.")
                raise SystemExit(1)
            
            # Atualizar lista se mudou
            if current_list != interactive_list:
                interactive_list = current_list
                log(f"Lista de scripts atualizada: {len(interactive_list)} scripts disponíveis.")
            
            # Menu destacado com cores e separadores
            print(f"\n{CYAN}{'=' * 60}{RESET}")
            print(f"{BOLD}{GREEN}  SELEÇÃO DE SCRIPT - AUTOMAÇÃO MDF-e{RESET}")
            print(f"{CYAN}{'=' * 60}{RESET}\n")
            
            print(f"{YELLOW}  [0]{RESET} Voltar ao menu anterior\n")
            
            for idx, name in enumerate(interactive_list, start=1):
                print(f"{YELLOW}  [{idx}]{RESET} {name}")
            
            print(f"\n{CYAN}{'─' * 60}{RESET}")
            print(f"{BOLD}Digite o número do script desejado (ou 0 para voltar):{RESET}")
            print(f"{CYAN}{'─' * 60}{RESET}\n")
            
            choice = input(f"{BOLD}Opção: {RESET}").strip()
            
            # Validar entrada
            if not choice:
                print(f"\n{RED}✗ Erro: Você deve selecionar um script!{RESET}")
                log("Entrada vazia; solicitando nova seleção.")
                time.sleep(SLEEP_ONE_HALF)
                continue
            
            # Verificar opção "0" para voltar
            if choice == "0":
                print(f"\n{GREEN}✓ Retornando ao menu anterior...{RESET}\n")
                log("Usuário escolheu retornar ao menu anterior")
                raise SystemExit(99)
            
            # Tentar converter para número
            if choice.isdigit():
                try:
                    index = int(choice) - 1
                    if 0 <= index < len(interactive_list):
                        selected_profile = interactive_list[index]
                        log(f"Script selecionado por índice {choice}: {selected_profile}")
                    else:
                        print(f"\n{RED}✗ Erro: Número inválido! Escolha entre 1 e {len(interactive_list)}.{RESET}")
                        log(f"Opção fora do intervalo: {choice}")
                        time.sleep(SLEEP_ONE_HALF)
                except (ValueError, IndexError) as e:
                    print(f"\n{RED}✗ Erro ao processar número: {str(e)}{RESET}")
                    log(f"Erro ao processar índice {choice}: {e}")
                    time.sleep(SLEEP_ONE_HALF)
            # Aceitar nome exato do arquivo (case-insensitive para maior flexibilidade)
            elif choice.lower() in [s.lower() for s in interactive_list]:
                # Encontrar o nome com case correto
                for script in interactive_list:
                    if script.lower() == choice.lower():
                        selected_profile = script
                        log(f"Script selecionado por nome: {selected_profile}")
                        break
            else:
                print(f"\n{RED}✗ Erro: Opção inválida! Digite um número válido.{RESET}")
                log(f"Opção inválida: {choice}")
                time.sleep(SLEEP_ONE_HALF)
                
        except KeyboardInterrupt:
            print(f"\n\n{RED}✗ Seleção cancelada pelo usuário.{RESET}")
            log("Seleção interrompida por Ctrl+C")
            raise SystemExit(0)
        except Exception as e:
            print(f"\n{RED}✗ Erro inesperado: {str(e)}{RESET}")
            log(f"ERRO durante seleção de perfil: {e}")
            time.sleep(SLEEP_ONE_HALF)
            # Continuar o loop para tentar novamente
            continue
    
    # Verificação de segurança
    if not selected_profile:
        log(f"Número máximo de tentativas atingido ({max_attempts}); nenhum perfil selecionado.")
        raise SystemExit(1)
    
    print(f"\n{GREEN}✓ Script selecionado: {BOLD}{selected_profile}{RESET}\n")
    print(f"{CYAN}{'=' * 60}{RESET}\n")
    
    # Fechar/ocultar o terminal após seleção (com fallback em caso de WM_CLOSE falhar)
    hide_console_window()
    return selected_profile

