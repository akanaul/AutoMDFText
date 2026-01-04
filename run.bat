@echo off
setlocal enabledelayedexpansion

set "PYTHON=python"
for %%i in (py py.exe) do (
    if not defined PYTHON_FOUND (
        where %%i >nul 2>&1 && set "PYTHON=%%i" && set "PYTHON_FOUND=1"
    )
)
if not defined PYTHON_FOUND (
    where python >nul 2>&1 && set "PYTHON=python" && set "PYTHON_FOUND=1"
)
if not defined PYTHON_FOUND (
    echo Python isn't available on PATH; please install it or consult IT.
    pause
    goto :EOF
)
set "SCRIPTS_DIR=%~dp0scripts"
if not exist "%SCRIPTS_DIR%" mkdir "%SCRIPTS_DIR%"

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
echo Installing required Python packages...^ (pyautogui pyperclip PyGetWindow Pillow numpy pyperclip win32gui^)
%PYTHON% -m pip install --upgrade pip >nul
%PYTHON% -m pip install pyautogui pyperclip Pillow pywin32 >nul
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
%PYTHON% modular_mdfe.py
goto prompt

:editor
echo Opening the profile editor.
%PYTHON% script_editor.py
goto prompt