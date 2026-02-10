@echo off
setlocal enabledelayedexpansion
title AutoMDFText - MDF-e Toolkit

set "ROOT=%~dp0"
set "VENV_DIR=%ROOT%.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "BASE_PY="

for %%i in (py py.exe python) do (
    if not defined BASE_PY (
        where %%i >nul 2>&1 && set "BASE_PY=%%i"
    )
)

if exist "%VENV_PY%" (
    set "PYTHON=%VENV_PY%"
) else (
    if not defined BASE_PY (
        echo Python isn't available on PATH; please install it or consult IT.
        pause
        goto :EOF
    )
    echo Creating local .venv for AutoMDFText...
    %BASE_PY% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Failed to create .venv. Check your Python installation.
        pause
        goto :EOF
    )
    set "PYTHON=%VENV_PY%"
)
set "SCRIPTS_DIR=%~dp0scripts"
if not exist "%SCRIPTS_DIR%" mkdir "%SCRIPTS_DIR%"

REM Startup guard: block duplicate run.bat windows and active automation
set "THISBAT=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $t='AutoMDFText - MDF-e Toolkit'; $wins = Get-Process -EA SilentlyContinue | Where-Object { $_.MainWindowTitle -eq $t }; if ($wins.Count -gt 1) { Write-Host 'Já existe outra janela do toolkit aberta. Fechando este terminal...'; exit 1 } else { exit 0 } } catch { exit 0 }"
IF ERRORLEVEL 1 goto :EOF
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $bat = '%~f0'; $name = [System.IO.Path]::GetFileName($bat); $runs = Get-CimInstance Win32_Process -EA SilentlyContinue ^| Where-Object { $_.Name -match 'cmd(\\d+)?\\.exe' -and $_.CommandLine -like ('*' + $name + '*') }; if ($runs.Count -gt 1) { Write-Host 'Já existe outra instância do run.bat aberta. Fechando este terminal...'; exit 1 }; $auto = Get-CimInstance Win32_Process -EA SilentlyContinue ^| Where-Object { $_.Name -match 'python(\\d+)?\\.exe' -and $_.CommandLine -match 'modular_mdfe\\.py' }; if ($auto) { Write-Host 'Automação já em execução. Use a janela existente.'; exit 1 }; exit 0 } catch { exit 0 }"
IF ERRORLEVEL 1 goto :EOF

:prompt
cls
echo Select an action for the MDF-e automation toolkit:
echo 1 ^| Run modular MDF-e filler
echo 2 ^| Launch the script template editor
echo 3 ^| Install/upgrade Python dependencies
echo 4 ^| Exit
choice /C 1234 /N /M "Your choice"
set "CHOICE_RESULT=!errorlevel!"

if %CHOICE_RESULT%==4 goto :EOF
if %CHOICE_RESULT%==3 goto install
if %CHOICE_RESULT%==2 goto editor
if %CHOICE_RESULT%==1 goto modular

:install
echo Installing required Python packages...^ (pyautogui pyperclip Pillow pywin32 pynput^)
%PYTHON% -m pip install --upgrade pip >nul
%PYTHON% -m pip install pyautogui pyperclip Pillow pywin32 pynput >nul
if errorlevel 1 (
    echo Installation failed, check the output above.
    pause
    goto prompt
)
echo Dependencies installed successfully.
pause
goto prompt

:modular
echo Running modular MDF-e filler. Press Ctrl+C to abort.
REM Guard before launching, ensure no existing automation is running
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $p = Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -match 'modular_mdfe\.py' }; if ($p) { Write-Host 'Automação já em execução. Retornando ao menu...'; exit 1 } else { exit 0 } } catch { exit 0 }"
IF ERRORLEVEL 1 goto prompt
%PYTHON% modular_mdfe.py
REM Check if user requested to return to menu (exit code 99)
IF ERRORLEVEL 99 goto prompt
goto modular

:editor
echo Opening the profile editor.
%PYTHON% script_editor.py
goto prompt