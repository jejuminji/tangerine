@echo off
setlocal
cd /d "%~dp0"

py -3.10 --version >nul 2>&1
if errorlevel 1 (
    echo Python 3.10 was not found. Install Python 3.10 with the Python Launcher.
    pause
    exit /b 1
)

py -3.10 -m streamlit run main.py
if errorlevel 1 (
    echo.
    echo Tangerine could not start. Check the error above.
    echo If packages are missing, run this command in the project folder:
    echo   py -3.10 -m pip install -r requirements.txt
    pause
    exit /b 1
)
