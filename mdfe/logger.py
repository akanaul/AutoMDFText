"""
Logging e output de console para a automação AutoMDFText.

Responsabilidades:
- Registrar mensagens em arquivo de log de sessão (log)
- Formatar saída colorida no console (ui_print)
- Inicializar uma nova sessão de log e reiniciar timers (start_automation_session)

Nota de dependência circular:
    resume_automation_timer (mdfe.timing) chama log() deste módulo via import lazy.
    start_automation_session deste módulo mutua globals de mdfe.timing via import lazy.
    Ambos os imports são feitos dentro das funções para evitar ciclo em nível de módulo.
"""

import time
from pathlib import Path

from constants import LOG_TIMESTAMP_FORMAT, LOG_MESSAGE_FORMAT

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent   # raiz do projeto (AutoMDFText/)
LOG_DIR  = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Timestamp e arquivo de log da sessão corrente.
# start_automation_session() sobrescreve ambos ao iniciar a automação de fato.
SESSION_TS: str  = time.strftime(LOG_TIMESTAMP_FORMAT, time.localtime())
LOG_FILE: Path   = LOG_DIR / f"automation_{SESSION_TS}.log"


def log(msg: str) -> None:
    """Registra mensagem apenas no arquivo de log, sem imprimir no console."""
    ts = time.strftime(LOG_MESSAGE_FORMAT, time.localtime())
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Não interromper fluxo por falha de log em disco
        pass


def ui_print(msg: str, style: str = "info") -> None:
    """Imprime mensagem formatada no console estilo GUI."""
    CYAN   = "\033[96m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    BLUE   = "\033[94m"
    BOLD   = "\033[1m"
    RESET  = "\033[0m"

    if style == "success":
        print(f"{GREEN}✓{RESET} {msg}")
    elif style == "error":
        print(f"{RED}✗{RESET} {msg}")
    elif style == "warning":
        print(f"{YELLOW}⚠{RESET} {msg}")
    elif style == "step":
        print(f"{BLUE}▸{RESET} {msg}")
    elif style == "header":
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{msg}{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
    else:
        print(f"  {msg}")


def start_automation_session(selected: str, profile_path: Path) -> float:
    """Cria um novo log de sessao e reinicia os contadores de tempo.

    Retorna o instante monotonic para calculo do tempo total.
    """
    global LOG_FILE, SESSION_TS

    # Import lazy para evitar dependência circular com mdfe.timing
    import mdfe.timing as _timing

    SESSION_TS = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    LOG_FILE   = LOG_DIR / f"automation_{SESSION_TS}.log"

    real_start_time = time.monotonic()
    _timing._automation_start_time  = time.monotonic()
    _timing._automation_time_paused = 0.0

    log("Iniciando automação (main)")
    log(f"[DEBUG] real_start_time={real_start_time}")
    log(f"Perfil selecionado: {selected}")
    log(f"Perfil carregado com sucesso de: {profile_path}")
    log(f"[DEBUG] _automation_start_time iniciado após escolha do perfil: {_timing._automation_start_time}")

    return real_start_time
