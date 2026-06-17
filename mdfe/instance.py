"""
Controle de instância única para AutoMDFText.

Responsabilidades:
- Garantir que apenas uma instância do motor de automação rode por vez,
  usando um Mutex nomeado do Windows (Global\\AutoMDFText_Mutex).
- Manter o handle do mutex vivo durante todo o tempo de processo para que
  o garbage collector não o libere prematuramente.

No-op em sistemas não-Windows.
"""

import ctypes
import os
from typing import Optional

from constants import MUTEX_NAME

# Handle mantido vivo pelo tempo de vida do processo.
# O GC não pode coletar enquanto o módulo estiver importado.
_SINGLETON_MUTEX_HANDLE: Optional[int] = None


def ensure_single_instance(name: str = MUTEX_NAME, on_duplicate: str = "warn") -> None:
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
                # Import lazy para evitar dependência circular com mdfe.dialogs (Parte 3)
                from mdfe.dialogs import focused_alert
                focused_alert("Já existe uma instância em execução. O processo será encerrado.")
        except Exception:
            pass
        raise SystemExit(0)
    else:
        # Manter handle vivo para não liberar o mutex
        global _SINGLETON_MUTEX_HANDLE
        _SINGLETON_MUTEX_HANDLE = handle
