@echo off
cd /d "%~dp0"

echo [1/3] Installing packages...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [ERROR] pip install failed. Check Python is installed.
    pause
    exit /b 1
)

if not exist ".env" (
    echo ANTHROPIC_API_KEY=sk-ant-여기에_API키_입력> .env
    echo .env file created. Opening for editing...
    notepad .env
    echo After saving, press any key to continue...
    pause >nul
)

echo [2/3] Starting server...
start "" python web_app.py
timeout /t 5 /nobreak >nul

echo [3/3] Opening browser...
start "" http://localhost:8000
