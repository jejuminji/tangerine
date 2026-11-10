@echo off
setlocal
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Python 3.13 virtual environment was not found.
    echo Run these commands in PowerShell from this folder:
    echo   py -3.13 -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -c "import sys; print('Python:', sys.version.split()[0]); sys.exit(0 if sys.version_info[:2] == (3, 13) else 1)"
if errorlevel 1 (
    echo [ERROR] This virtual environment must use Python 3.13.
    echo Install Python 3.13 first, then close Tangerine and run in PowerShell:
    echo   Rename-Item -LiteralPath .venv -NewName .venv-backup
    echo   py -3.13 -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    echo Keep .venv-backup until the new environment works.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -c "import openai; print('OpenAI SDK:', openai.__version__)"
if errorlevel 1 (
    echo [ERROR] Python or OpenAI SDK check failed.
    pause
    exit /b 1
)

if not exist "logs\streamlit" mkdir "logs\streamlit"
if not exist "logs\streamlit" (
    echo [ERROR] Could not create the Streamlit log folder.
    pause
    exit /b 1
)

echo Starting Tangerine. Close this window or press Ctrl+C to stop.
echo Each run saves its logs in a separate folder under logs\streamlit.
".venv\Scripts\python.exe" -u run_streamlit.py
if errorlevel 1 (
    echo.
    echo [ERROR] Tangerine could not start.
    echo Check the message above and the latest folder under logs\streamlit.
    pause
    exit /b 1
)
