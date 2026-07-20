@echo off

cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Python was not found on PATH. Install Python first.
    pause
    exit /b 1
)

echo Running weekly report generator...
py main.py 2> log.txt

echo.
echo Report generation complete.
pause