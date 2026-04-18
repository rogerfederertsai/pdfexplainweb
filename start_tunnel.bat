@echo off
cd /d "%~dp0"

echo ====================================================
echo   Cloudflare Tunnel (WAN access)
echo ====================================================
echo.
echo When ready, this window shows a URL like https://xxxx.trycloudflare.com
echo Send that URL to remote users. They need nothing installed.
echo.
echo If you close this window, WAN access stops. LAN http://IP:8000/ still works if server runs.
echo ====================================================
echo.

where cloudflared >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] cloudflared not found in PATH.
    echo.
    echo Install steps:
    echo   1. https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
    echo   2. Download Windows amd64 build (cloudflared-windows-amd64.exe)
    echo   3. Rename to cloudflared.exe and put in a folder on PATH (e.g. System32) or add its folder to PATH
    echo.
    pause
    exit /b 1
)

cloudflared tunnel --url http://localhost:8000
pause
