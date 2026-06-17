"""
Controle de tempo de automação para AutoMDFText.

Responsabilidades:
- Rastrear tempo total de execução da sessão
- Descontar pausas causadas por prompts e diálogos do operador
- Formatar durações para exibição no resumo final

Nota de dependência:
    resume_automation_timer chama log() de mdfe.logger via import lazy
    para evitar dependência circular em nível de módulo.
"""

import time


# ── Estado de tempo ──────────────────────────────────────────────────────────
_automation_start_time: float  = 0.0
_automation_time_paused: float = 0.0
_pause_start_time: float       = 0.0


def pause_automation_timer() -> None:
    """Pausa o contador de tempo de automação (usado durante prompts)."""
    global _pause_start_time
    _pause_start_time = time.monotonic()


def resume_automation_timer() -> None:
    """Resume o contador de tempo de automação após prompt."""
    global _automation_time_paused, _pause_start_time
    if _pause_start_time > 0:
        _automation_time_paused += time.monotonic() - _pause_start_time
        _pause_start_time = 0.0

    # Import lazy para evitar dependência circular com mdfe.logger
    from mdfe.logger import log
    log(f"[DEBUG PAUSE] Resuming timer. Total paused so far: {_automation_time_paused}s")


def format_duration(seconds: float) -> str:
    """Formata duracao em segundos para o formato MM:SS ou apenas SS."""
    minutes = int(seconds // 60)
    secs    = int(seconds % 60)
    if minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"
