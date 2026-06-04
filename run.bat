@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
echo Starting Excel Comparator...
REM Try python first, then py (Windows Launcher), then python3
python main.py 2>nul || py main.py 2>nul || python3 main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to start. Make sure Python is installed and in PATH.
    echo         Install dependencies: pip install -r requirements.txt
)
pause
