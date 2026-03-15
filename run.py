#!/usr/bin/env python3
"""서버 실행 스크립트 - 단가표_실행.bat 에서 호출됩니다."""

import os
import sys
import subprocess
from pathlib import Path

BASE = Path(__file__).parent
ENV_FILE = BASE / ".env"


def load_env():
    """env 파일에서 API 키 로드"""
    if not ENV_FILE.exists():
        return False
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("ANTHROPIC_API_KEY="):
            key = line.split("=", 1)[1].strip()
            if key:
                os.environ["ANTHROPIC_API_KEY"] = key
                return True
    return False


def main():
    print("[1/3] Installing packages...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r",
                    str(BASE / "requirements.txt"), "-q"], check=True)

    if not load_env():
        print()
        print("=" * 50)
        print("ANTHROPIC_API_KEY not found.")
        print("Enter your API key (sk-ant-...):")
        key = input("> ").strip()
        ENV_FILE.write_text(f"ANTHROPIC_API_KEY={key}\n", encoding="utf-8")
        os.environ["ANTHROPIC_API_KEY"] = key
        print(".env file created.")
        print("=" * 50)
        print()

    print("[2/3] Starting server...")
    import webbrowser, threading, time

    def open_browser():
        time.sleep(5)
        webbrowser.open("http://localhost:8000")

    threading.Thread(target=open_browser, daemon=True).start()

    print("[3/3] Browser will open in 5 seconds...")
    print("Press Ctrl+C to stop the server.")
    print()

    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8000,
                reload=False, app_dir=str(BASE))


if __name__ == "__main__":
    main()
