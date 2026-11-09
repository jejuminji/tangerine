@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Python environment not found. Run these commands in this folder first:
    echo   py -3.10 -m venv .venv
    echo   .venv\Scripts\python -m pip install -r requirements.txt
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m streamlit run main.py
if errorlevel 1 (
    echo.
    echo Tangerine could not start. Check the error above.
    pause
    exit /b 1
)
