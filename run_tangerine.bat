@echo off
setlocal
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

".venv\Scripts\python.exe" -m streamlit run main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Tangerine could not start. Check the error above.
    pause
    exit /b 1
)
