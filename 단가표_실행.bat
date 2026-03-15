@echo off
cd /d "%~dp0"

:: 패키지 설치 확인
echo [1/3] 필요 패키지 설치 중...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [오류] 패키지 설치 실패. Python/pip이 설치되어 있는지 확인하세요.
    pause
    exit /b 1
)

:: .env 파일 존재 확인
if not exist ".env" (
    echo [오류] .env 파일이 없습니다.
    echo .env.example 파일을 복사해서 .env 를 만들고 API 키를 입력하세요.
    pause
    exit /b 1
)

echo [2/3] 서버 시작 중...
start "" python web_app.py
timeout /t 5 /nobreak >nul

echo [3/3] 브라우저 열기...
start "" http://localhost:8000
