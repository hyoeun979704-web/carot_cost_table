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
    echo [ERROR] .env file not found.
    echo Please create .env file and set ANTHROPIC_API_KEY=sk-ant-...
    pause
    exit /b 1
)

echo [2/3] Starting server...
start "" python web_app.py
timeout /t 5 /nobreak >nul

echo [3/3] Opening browser...
start "" http://localhost:8000
