@echo off
setlocal
set "TANGERINE_DIR=%~dp0"
powershell.exe -NoProfile -Command "$ErrorActionPreference = 'Stop'; $root = $env:TANGERINE_DIR; $shell = New-Object -ComObject WScript.Shell; $link = $shell.CreateShortcut((Join-Path $root 'Tangerine.lnk')); $link.TargetPath = Join-Path $root 'run_tangerine.bat'; $link.WorkingDirectory = $root; $link.IconLocation = (Join-Path $root 'tangerine.ico') + ',0'; $link.Description = 'Run Tangerine'; $link.Save()"
if errorlevel 1 (
    echo Could not create the shortcut.
    pause
    exit /b 1
)
echo Tangerine.lnk created. Uses installed Python 3.10 without a virtual environment.
echo Double-click its orange icon to start.
pause
