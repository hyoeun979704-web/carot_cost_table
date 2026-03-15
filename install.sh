#!/usr/bin/env bash
# ────────────────────────────────────────────────────────────
# 단가표 웹 서비스 설치 스크립트 (최초 1회만 실행)
#
# 실행 후:
#   - PC를 켜면 서버가 자동 시작됩니다.
#   - 브라우저에서 http://localhost:8000 으로 바로 접속하면 됩니다.
# ────────────────────────────────────────────────────────────
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SERVICE_NAME="carot-web"
SERVICE_FILE="$HOME/.config/systemd/user/$SERVICE_NAME.service"
AUTOSTART_FILE="$HOME/.config/autostart/$SERVICE_NAME-browser.desktop"
ENV_FILE="$SCRIPT_DIR/.env"

echo "=============================="
echo " 단가표 웹 서비스 설치"
echo "=============================="
echo ""

# ── 1. ANTHROPIC_API_KEY 설정 ─────────────────────────────
if [ -f "$ENV_FILE" ] && grep -q "ANTHROPIC_API_KEY=sk-" "$ENV_FILE" 2>/dev/null; then
    echo "✓ .env 파일이 이미 있습니다. API 키 설정을 건너뜁니다."
else
    echo "Anthropic API 키를 입력하세요 (sk-ant-... 형태):"
    read -rsp "  API_KEY: " API_KEY
    echo ""
    if [[ ! "$API_KEY" == sk-* ]]; then
        echo "❌ 올바른 API 키 형식이 아닙니다 (sk-로 시작해야 합니다)."
        exit 1
    fi
    echo "ANTHROPIC_API_KEY=$API_KEY" > "$ENV_FILE"
    echo "✓ .env 파일 저장 완료"
fi

# ── 2. Python 의존성 설치 ─────────────────────────────────
echo ""
echo "패키지 설치 중..."
pip install -r "$SCRIPT_DIR/requirements.txt" --quiet
echo "✓ 패키지 설치 완료"

# ── 3. systemd 유저 서비스 생성 ───────────────────────────
UVICORN_BIN="$(python -m site --user-base)/bin/uvicorn"
# PATH 내 uvicorn 우선, 없으면 site-packages 경로 사용
if command -v uvicorn &>/dev/null; then
    UVICORN_BIN="$(command -v uvicorn)"
fi

mkdir -p "$HOME/.config/systemd/user"
cat > "$SERVICE_FILE" << EOF
[Unit]
Description=단가표 자동 업데이트 웹 서비스
After=network.target

[Service]
Type=simple
WorkingDirectory=$SCRIPT_DIR
EnvironmentFile=$ENV_FILE
ExecStart=$UVICORN_BIN web_app:app --host 127.0.0.1 --port 8000
Restart=on-failure
RestartSec=5

[Install]
WantedBy=default.target
EOF

systemctl --user daemon-reload
systemctl --user enable "$SERVICE_NAME"
systemctl --user start  "$SERVICE_NAME"
echo "✓ 서비스 등록 및 시작 완료"

# ── 4. 로그인 시 브라우저 자동 열기 (선택) ────────────────
echo ""
read -rp "로그인할 때 브라우저를 자동으로 열까요? [y/N]: " AUTO_BROWSER
if [[ "${AUTO_BROWSER,,}" == "y" ]]; then
    mkdir -p "$HOME/.config/autostart"
    cat > "$AUTOSTART_FILE" << EOF
[Desktop Entry]
Type=Application
Name=단가표 웹 서비스 열기
Exec=bash -c 'sleep 5 && xdg-open http://localhost:8000'
Hidden=false
NoDisplay=false
X-GNOME-Autostart-enabled=true
EOF
    echo "✓ 로그인 시 자동으로 브라우저가 열립니다."
else
    echo "  → 브라우저 자동 열기를 건너뜁니다."
    echo "     접속 주소: http://localhost:8000"
fi

# ── 5. 완료 ──────────────────────────────────────────────
echo ""
echo "=============================="
echo "✅ 설치 완료!"
echo ""
echo "  접속 주소: http://localhost:8000"
echo ""
echo "  서비스 관리:"
echo "    상태 확인: systemctl --user status $SERVICE_NAME"
echo "    로그 보기: journalctl --user -u $SERVICE_NAME -f"
echo "    수동 중지: systemctl --user stop $SERVICE_NAME"
echo "    제거:      systemctl --user disable --now $SERVICE_NAME"
echo "=============================="
