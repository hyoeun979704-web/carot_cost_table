@echo off
cd /d "%~dp0"
start "" python web_app.py
timeout /t 2 /nobreak >nul
start "" http://localhost:8000
