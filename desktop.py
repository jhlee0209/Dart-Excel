"""맥 앱 실행기: Streamlit 서버를 띄우고 브라우저 대신 전용 창에 보여 준다.

'DART 데이터북.app'이 이 파일을 실행한다. 창을 닫으면 서버도 함께 끝난다.
"""
from __future__ import annotations

import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import webview

ROOT = Path(__file__).resolve().parent
LOG = Path.home() / "Library" / "Logs" / "DART Databook" / "server.log"
TITLE = "DART 미니 데이터북"
STARTUP_SECONDS = 90

PAGE = """<html><body style="margin:0;height:100vh;display:flex;align-items:center;justify-content:center;
background:#FAF8F4;font-family:-apple-system,'Apple SD Gothic Neo',sans-serif;color:#18222F">
<div style="text-align:center"><div style="font-size:12px;letter-spacing:.16em;color:#9A7B4F;font-weight:700">
DART MINI DATABOOK</div><div style="font-size:22px;font-weight:700;margin-top:10px">{title}</div>
<div style="font-size:14px;color:#6B7480;margin-top:8px;line-height:1.7">{body}</div></div></body></html>"""


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_until_up(url: str, server: subprocess.Popen) -> bool:
    deadline = time.time() + STARTUP_SECONDS
    while time.time() < deadline and server.poll() is None:
        try:
            with urllib.request.urlopen(f"{url}/_stcore/health", timeout=1) as res:
                if res.status == 200:
                    return True
        except OSError:
            time.sleep(0.3)
    return False


def main() -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    url = f"http://127.0.0.1:{free_port()}"
    with LOG.open("w") as log:
        server = subprocess.Popen(
            [
                sys.executable, "-m", "streamlit", "run", str(ROOT / "app.py"),
                "--server.address", "127.0.0.1", "--server.port", url.rsplit(":", 1)[1],
                "--server.headless", "true", "--server.fileWatcherType", "none",
                "--browser.gatherUsageStats", "false",
            ],
            cwd=ROOT, stdout=log, stderr=subprocess.STDOUT,
        )
        try:
            webview.settings["ALLOW_DOWNLOADS"] = True  # 엑셀 파일 내려받기 버튼이 저장 창을 띄우도록
            window = webview.create_window(
                TITLE, html=PAGE.format(title="여는 중입니다…", body="잠시만 기다려 주세요."),
                width=1440, height=920, min_size=(1040, 700),
            )

            def boot() -> None:
                if wait_until_up(url, server):
                    window.load_url(url)
                else:
                    window.load_html(PAGE.format(
                        title="앱을 시작하지 못했습니다",
                        body=f"자세한 내용은 아래 기록 파일에 있습니다.<br>{LOG}",
                    ))

            webview.start(boot)
        finally:
            server.terminate()
            try:
                server.wait(timeout=5)
            except subprocess.TimeoutExpired:
                server.kill()


if __name__ == "__main__":
    main()
