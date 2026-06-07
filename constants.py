"""
Constantes centralizadas para automação MDF-e e extensões.

Este módulo define todos os delays, timeouts e configurações utilizados
pela automação principal e seus scripts de extensão. Centralizar aqui
facilita manutenção e desacoplamentos futuros.
"""

# ═══════════════════════════════════════════════════════════════════════
# DELAYS DE NAVEGAÇÃO (em segundos)
# ═══════════════════════════════════════════════════════════════════════

TAB_DELAY = 0.10
"""Delay padrão entre pressionamentos de Tab"""

TAB_DELAY_LONG = 0.25
"""Delay estendido para Tab (campos mais sensíveis)"""

CTRL_F_DELAY = 0.2
"""Delay para operações com Ctrl+F (procura)"""

DROPDOWN_SETTLE_DELAY = 0.4
"""Delay para dropdowns se estabilizarem após interação"""

# ═══════════════════════════════════════════════════════════════════════
# DELAYS GERAIS DE ESPERA (em segundos)
# ═══════════════════════════════════════════════════════════════════════

SLEEP_SHORT = 0.15
"""Pausa curta entre ações (não-crítico)"""

SLEEP_MEDIUM = 0.25
"""Pausa média entre ações (navegação)"""

SLEEP_LONG = 0.40
"""Pausa longa (carregamento de página, dropdown)"""

SLEEP_LONGER = 0.7
"""Pausa estendida (transições, mudanças de abas)"""

SLEEP_ONE = 1.0
"""Pausa de 1 segundo (operações críticas, espera de página)"""

SLEEP_ONE_HALF = 1.5
"""Pausa de 1.5 segundos (menu selection, long operations)"""

# ═══════════════════════════════════════════════════════════════════════
# TIMEOUTS E LIMITES (em segundos)
# ═══════════════════════════════════════════════════════════════════════

FORM_WAIT_TIMEOUT = 15.0
"""Timeout máximo para aguardar formulário abrir"""

FORM_WAIT_INTERVAL = 1.0
"""Intervalo entre tentativas de detecção de formulário"""

FORM_COPY_ATTEMPTS = 2
"""Número de tentativas ao copiar conteúdo de página"""

VERIFY_FIELD_TIMEOUT = 6.0
"""Timeout para verificação de campo na página"""

# ═══════════════════════════════════════════════════════════════════════
# DIMENSÕES E OFFSETS DE JANELA (pixels)
# ═══════════════════════════════════════════════════════════════════════

EDGE_SEARCHBAR_HEIGHT = 70
"""Altura da barra de pesquisa do Edge (px)"""

EDGE_CLICK_OFFSET = 200
"""Offset para click abaixo da barra de pesquisa (px)"""

# ═══════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES DE ESCRITA INTELIGENTE
# ═══════════════════════════════════════════════════════════════════════

MIN_PASTE_LENGTH = 4
"""Comprimento mínimo de texto para usar clipboard ao invés de digitar"""

WRITE_INTERVAL = 0.10
"""Intervalo entre caracteres ao digitar"""

# ═══════════════════════════════════════════════════════════════════════
# FORMATOS E PADRÕES
# ═══════════════════════════════════════════════════════════════════════

LOG_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"
"""Formato de timestamp para nomes de arquivo de log"""

LOG_MESSAGE_FORMAT = "%Y-%m-%dT%H:%M:%S"
"""Formato de timestamp para mensagens de log"""

# ═══════════════════════════════════════════════════════════════════════
# MUTEX E GERENCIAMENTO DE INSTÂNCIA
# ═══════════════════════════════════════════════════════════════════════

MUTEX_NAME = "Global\\AutoMDFText_Mutex"
"""Nome do Mutex global para garantir single-instance"""
