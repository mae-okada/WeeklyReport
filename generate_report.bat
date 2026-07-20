@echo off

cd /d "%~dp0"

if not exist requirements.txt (
    echo requirements.txt not found.
    pause
    exit /b 1
)

where py >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Python was not found on PATH. Install Python first.
    pause
    exit /b 1
)

py -m pip install -r requirements.txt
py main.py 2> log.txt

pause