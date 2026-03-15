@echo off
cd /d "%~dp0"

echo [1/3] Installing packages...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

if not exist "%~dp0.env" (
    echo.
    echo Enter your Anthropic API key (sk-ant-...):
    set /p APIKEY=^>
    (echo ANTHROPIC_API_KEY=%APIKEY%)> "%~dp0.env"
    echo .env file created.
    echo.
)

for /f "usebackq tokens=1,* delims==" %%a in ("%~dp0.env") do (
    if "%%a"=="ANTHROPIC_API_KEY" set ANTHROPIC_API_KEY=%%b
)

if "%ANTHROPIC_API_KEY%"=="" (
    echo [ERROR] ANTHROPIC_API_KEY not found in .env
    pause
    exit /b 1
)

echo API key loaded OK.
echo [2/3] Starting server...
start "" python "%~dp0web_app.py"
timeout /t 5 /nobreak >nul

echo [3/3] Opening browser...
start "" http://localhost:8000
