@echo off
cd /d "%~dp0"

:: Load WEB_LOGIN_* from config.bat (not committed to git)
call "%~dp0config.bat"

echo ====================================================
echo   PDF web server + Cloudflare Tunnel (LAN + WAN)
echo ====================================================

set "LAN_IP="
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do (
    set "CANDIDATE=%%A"
    setlocal EnableDelayedExpansion
    set "CANDIDATE=!CANDIDATE: =!"
    if not "!CANDIDATE!"=="127.0.0.1" (
        endlocal
        set "LAN_IP=%%A"
        set "LAN_IP=!LAN_IP: =!"
        goto :found_ip
    )
    endlocal
)
:found_ip

echo.
if defined LAN_IP (
    echo [LAN] http://%LAN_IP%:8000/
) else (
    echo [LAN] Could not detect IPv4. Check Network settings for this PC IP.
)
echo [WAN] Open the new "Cloudflare Tunnel" window and copy the https URL shown there.
echo.
echo ====================================================
echo.

start "PDF Web Server" cmd /c "cd /d "%~dp0project" && python run_web_launcher.py"

timeout /t 4 /nobreak >nul

start "Cloudflare Tunnel" cmd /k "cloudflared tunnel --url http://localhost:8000"

echo Both started.
echo   - PDF Web Server: detached (re-run this .bat to open browser again if needed)
echo   - Cloudflare Tunnel: see the [Cloudflare Tunnel] window for the public URL
echo.
echo Username: %WEB_LOGIN_USERNAME%
echo Password: see config.bat on this machine
echo.
pause
