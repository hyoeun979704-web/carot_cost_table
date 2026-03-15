@echo off
cd /d "%~dp0"

:: .env 파일 존재 확인
if not exist ".env" (
    echo [오류] .env 파일이 없습니다.
    echo .env.example 파일을 복사하여 .env 를 만들고 API 키를 입력하세요.
    pause
    exit /b 1
)

start "" python web_app.py
timeout /t 5 /nobreak >nul
start "" http://localhost:8000
