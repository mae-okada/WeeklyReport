@echo off
setlocal

cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Python launcher not found. Install Python first and ensure it is on PATH.
    pause
    exit /b 1
)

py -m pip install --upgrade pip
py -m pip install -r requirements.txt

echo.
echo Dependencies installed successfully.
pause
