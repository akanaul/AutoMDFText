"""
Failsafe de teclado e estado de pausa para AutoMDFText.

Responsabilidades:
- Listener global de F8 (encerrar imediatamente) via pynput
- Listener global de F9 (sinalizar pausa para o próximo ponto seguro)
- Deter estado de pausa compartilhado (_pause_requested, _pause_active, _pause_lock)
  lido por mdfe.pause para executar a pausa efetiva

Nota de design:
    O estado de pausa vive aqui (e não em mdfe.pause) porque o gatilho F9
    é parte do mesmo listener de failsafe. mdfe.pause importa daqui.
"""

import ctypes
import os
import threading
from typing import Optional

# ── Estado do listener ────────────────────────────────────────────────────────
_failsafe_listener: Optional[object] = None

# ── Estado de pausa — lido por mdfe.pause ─────────────────────────────────────
_pause_requested: bool            = False
_pause_active: bool               = False
_pause_lock: threading.Lock       = threading.Lock()


def request_pause() -> None:
    """Sinaliza uma pausa para o próximo ponto seguro (acionada pelo F9)."""
    global _pause_requested
    with _pause_lock:
        _pause_requested = True


def start_failsafe_f8() -> None:
    """Inicia listener global para F8 (encerrar) e F9 (pausar).

    Em Windows, ignora eventos injetados para aceitar apenas teclado fisico.
    """
    global _failsafe_listener
    if _failsafe_listener is not None:
        return
    try:
        from pynput import keyboard
    except Exception as exc:
        from mdfe.logger import log
        log(f"Aviso: pynput nao disponivel; failsafe F8 desativado ({exc})")
        return

    injected = {"value": False}

    def win32_event_filter(_msg, data):
        flags = getattr(data, "flags", 0)
        injected["value"] = bool(flags & 0x10)
        return True

    def show_failsafe_alert() -> None:
        if os.name != "nt":
            return
        try:
            ctypes.windll.user32.MessageBoxW(
                0,
                "A automacao foi encerrada pelo botao de seguranca (F8).",
                "Automacao encerrada",
                0x00000040 | 0x00040000 | 0x00010000,
            )
        except Exception:
            pass

    def on_press(key) -> None:
        if injected["value"]:
            return
        if key == keyboard.Key.f8:
            from mdfe.logger import log
            log("Failsafe F8 acionado. Encerrando automacao.")
            show_failsafe_alert()
            os._exit(1)
        if key == keyboard.Key.f9:
            request_pause()

    listener_kwargs: dict = {"on_press": on_press}
    if os.name == "nt":
        listener_kwargs["win32_event_filter"] = win32_event_filter
    _failsafe_listener = keyboard.Listener(**listener_kwargs)
    _failsafe_listener.start()


def stop_failsafe_f8() -> None:
    """Finaliza o listener de failsafe por F8, se ativo."""
    global _failsafe_listener
    if _failsafe_listener is None:
        return
    try:
        _failsafe_listener.stop()
    except Exception:
        pass
    _failsafe_listener = None
