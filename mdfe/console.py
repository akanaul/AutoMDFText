"""
Gerenciamento da janela de console para AutoMDFText.

Responsabilidades:
- Ocultar o terminal durante a automação (hide_console_window)
- Restaurá-lo como popup ao final, trazendo-o para o topo (restore_console_popup)
- Emitir beep sonoro ao concluir (play_low_beep)

Todas as funções são no-ops em sistemas não-Windows.
"""

import ctypes
import os
import time


def hide_console_window() -> None:
    """Oculta a janela do console sem encerrar o processo."""
    if os.name != "nt":
        return
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if not hwnd:
        return
    # Apenas oculta o console para evitar encerramento do processo durante o debug
    ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE


def restore_console_popup() -> None:
    """Restaura o terminal, trazendo-o ao topo como um popup curto."""
    if os.name != "nt":
        return
    try:
        user32  = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        hwnd = kernel32.GetConsoleWindow()
        if not hwnd:
            return

        SW_SHOW       = 5
        HWND_TOPMOST  = -1
        HWND_NOTOPMOST = -2
        SWP_NOMOVE    = 0x0002
        SWP_NOSIZE    = 0x0001
        SWP_SHOWWINDOW = 0x0040

        # Mostrar e trazer para frente
        user32.ShowWindow(hwnd, SW_SHOW)
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
        user32.BringWindowToTop(hwnd)

        # Tornar topmost brevemente para comportamento de popup
        user32.SetWindowPos(hwnd, HWND_TOPMOST,    0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        time.sleep(0.15)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST,  0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
    except Exception:
        pass


def play_low_beep() -> None:
    """Emite um beep de baixa frequência ao final da automação."""
    try:
        if os.name == "nt":
            # Frequência baixa (~400Hz), duração 180ms
            import winsound
            winsound.Beep(400, 180)
        else:
            # Fallback para bell character em outros sistemas
            print("\a", end="", flush=True)
    except Exception:
        pass
