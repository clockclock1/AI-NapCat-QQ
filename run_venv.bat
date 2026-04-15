@echo off
setlocal EnableExtensions

cd /d "%~dp0"
set "VENV_DIR=.venv"
set "APP_SCRIPT=napcat_screenshot_ai.py"
set "REQ_FILE=requirements.txt"

echo [0/5] Stopping previous napcat_screenshot_ai.py instances...
for /f "usebackq delims=" %%p in (`powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -match 'python' -and $_.CommandLine -match 'napcat_screenshot_ai.py' } | Select-Object -ExpandProperty ProcessId"`) do (
    if not "%%p"=="" taskkill /PID %%p /F >nul 2>nul
)

echo [1/5] Checking Python...
set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD (
    where python >nul 2>nul && set "PY_CMD=python"
)
if not defined PY_CMD (
    echo [ERROR] Python not found. Install Python 3 first.
    goto :fail
)

echo [2/5] Creating virtual environment if needed...
if not exist "%VENV_DIR%\Scripts\python.exe" (
    %PY_CMD% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        goto :fail
    )
)

echo [3/5] Installing/updating dependencies...
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if exist "%REQ_FILE%" (
    "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%REQ_FILE%"
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        goto :fail
    )
) else (
    echo [WARN] requirements.txt not found, skipping dependency install.
)

echo [4/5] Preparing config...
if not exist "config.json" (
    if exist "config.json.example" (
        copy /Y "config.json.example" "config.json" >nul
        echo [INFO] config.json created from config.json.example
        echo [INFO] Please edit config.json with your real settings.
    ) else (
        echo [WARN] config.json and config.json.example are both missing.
    )
)

if /I "%~1"=="--setup-only" (
    echo [5/5] Setup complete. You can run this file again to start the app.
    goto :ok
)

echo [5/5] Starting app with virtual environment...
"%VENV_DIR%\Scripts\python.exe" "%APP_SCRIPT%"
if errorlevel 1 (
    echo [ERROR] App exited with error code %errorlevel%.
    goto :fail
)

goto :ok

:fail
echo.
echo Setup/start failed.
pause
exit /b 1

:ok
echo.
echo Done.
if /I not "%~1"=="--no-pause" pause
exit /b 0
