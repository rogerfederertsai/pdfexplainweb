# -*- coding: utf-8 -*-
"""
Headless launcher: ensure uvicorn on port 8000 (start detached if needed), then open URL in
Chrome/Edge (new tab when browser already running). Launcher exits immediately; server keeps
running. Re-run anytime to open another tab. No success dialog.
"""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
HOST = "127.0.0.1"
PORT = 8000
URL = f"http://{HOST}:{PORT}/"

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

_uvicorn_proc: subprocess.Popen | None = None


def _log(msg: str) -> None:
    try:
        log_path = Path(os.environ.get("TEMP", ".")) / "pdf_web_launcher.log"
        log_path.write_text(msg, encoding="utf-8")
    except OSError:
        pass


def _port_listening(timeout: float = 0.35) -> bool:
    try:
        with socket.create_connection((HOST, PORT), timeout=timeout):
            return True
    except OSError:
        return False


def _wait_port(timeout_s: float = 120.0, interval: float = 0.25) -> bool:
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        try:
            with socket.create_connection((HOST, PORT), timeout=1.0):
                return True
        except OSError:
            time.sleep(interval)
    return False


def _uvicorn_subprocess_kwargs() -> dict:
    """Child outlives this launcher so double-click run_web.bat again only opens a new tab."""
    if sys.platform == "win32":
        detached = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
        newpg = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
        return {
            "creationflags": int(CREATE_NO_WINDOW | detached | newpg),
        }
    return {"start_new_session": True}


def _kill_tree_win(pid: int) -> None:
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        capture_output=True,
        creationflags=CREATE_NO_WINDOW,
        stdin=subprocess.DEVNULL,
    )


def _browser_candidates() -> list[Path]:
    """At most one Chrome and one Edge path."""
    env = os.environ
    chrome_path: Path | None = None
    edge_path: Path | None = None
    for key in ("ProgramFiles(x86)", "ProgramFiles", "LocalAppData"):
        base = env.get(key)
        if not base:
            continue
        b = Path(base)
        cexe = b / "Google" / "Chrome" / "Application" / "chrome.exe"
        eexe = b / "Microsoft" / "Edge" / "Application" / "msedge.exe"
        if chrome_path is None and cexe.is_file():
            chrome_path = cexe
        if edge_path is None and eexe.is_file():
            edge_path = eexe
    out: list[Path] = []
    if chrome_path is not None:
        out.append(chrome_path)
    if edge_path is not None:
        out.append(edge_path)
    return out


def _open_browser() -> None:
    candidates = _browser_candidates()
    for exe in candidates:
        try:
            subprocess.Popen(
                [str(exe), URL],
                cwd=str(ROOT),
                creationflags=CREATE_NO_WINDOW,
            )
            return
        except OSError:
            continue
    try:
        import webbrowser

        webbrowser.open(URL)
    except Exception:
        _log("webbrowser.open failed")


def main() -> int:
    global _uvicorn_proc
    py = sys.executable
    env = os.environ.copy()
    env["NO_COLOR"] = "1"
    # 停用 HuggingFace XET 傳輸協定，改用穩定的 HTTP 下載，避免多次中斷重啟
    env["HF_HUB_DISABLE_XET"] = "1"
    args = [
        py,
        "-m",
        "uvicorn",
        "web.app:app",
        "--host",
        "0.0.0.0",
        "--port",
        str(PORT),
        "--no-use-colors",
    ]

    we_started = False
    if not _port_listening():
        extra = _uvicorn_subprocess_kwargs()
        try:
            _uvicorn_proc = subprocess.Popen(
                args,
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                **extra,
            )
            we_started = True
        except OSError as e:
            _log(f"uvicorn spawn failed: {e!r}")
            # Race: another launcher may have bound the port first
            time.sleep(0.5)
            if not _port_listening(2.0):
                return 1

    if not _wait_port():
        _log("timeout waiting for port 8000")
        if (
            we_started
            and _uvicorn_proc is not None
            and _uvicorn_proc.poll() is None
            and sys.platform == "win32"
        ):
            _kill_tree_win(_uvicorn_proc.pid)
        return 1

    _open_browser()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        _log(repr(exc))
        raise SystemExit(1)
