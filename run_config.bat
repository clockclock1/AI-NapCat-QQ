@echo off
setlocal EnableExtensions

cd /d "%~dp0"
set "VENV_DIR=.venv"
set "REQ_FILE=requirements.txt"
set "CFG_SCRIPT=config.py"

echo [1/4] Checking Python...
set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD (
    where python >nul 2>nul && set "PY_CMD=python"
)
if not defined PY_CMD (
    echo [ERROR] Python not found. Install Python 3 first.
    goto :fail
)

echo [2/4] Creating virtual environment if needed...
if not exist "%VENV_DIR%\Scripts\python.exe" (
    %PY_CMD% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        goto :fail
    )
)

echo [3/4] Installing/updating dependencies...
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if exist "%REQ_FILE%" (
    "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%REQ_FILE%"
)

echo [4/4] Running window picker...
"%VENV_DIR%\Scripts\python.exe" "%CFG_SCRIPT%"
if errorlevel 1 goto :fail

echo.
echo Done.
pause
exit /b 0

:fail
echo.
echo Failed.
pause
exit /b 1
