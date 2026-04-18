@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

call "%~dp0config.bat"

echo ====================================================
echo   PDF web server + Cloudflare Tunnel (LAN + WAN)
echo ====================================================

set "LAN_IP="
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do (
    set "_ip=%%A"
    set "_ip=!_ip: =!"
    if not "!_ip!"=="127.0.0.1" (
        set "LAN_IP=!_ip!"
        goto :ip_done
    )
)
:ip_done

echo.
if defined LAN_IP (
    echo [LAN] http://!LAN_IP!:8000/
) else (
    echo [LAN] Could not detect IPv4. Run ipconfig to find this PC address.
)
echo.

set "HAS_CF=0"
where cloudflared >nul 2>&1
if not errorlevel 1 set "HAS_CF=1"

if "!HAS_CF!"=="0" (
    echo [WAN] cloudflared NOT in PATH - no public trycloudflare URL until you install it.
    echo       Try: winget install --id Cloudflare.cloudflared -e
    echo       Or:  https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
    echo.
) else (
    echo [WAN] A window titled Cloudflare Tunnel will open - copy the https URL from that window.
    echo.
)

echo ====================================================
echo.

start "PDF Web Server" cmd /c "cd /d "%~dp0project" && python run_web_launcher.py"

timeout /t 4 /nobreak >nul

if "!HAS_CF!"=="1" (
    start "Cloudflare Tunnel" cmd /k "cloudflared tunnel --url http://localhost:8000"
)

echo.
if "!HAS_CF!"=="1" (
    echo OK: Web server started. WAN URL is in the Cloudflare Tunnel window.
) else (
    echo OK: Web server started. WAN skipped - install cloudflared, then run start_all.bat again.
)
echo.
echo Username: %WEB_LOGIN_USERNAME%
echo Password: see config.bat on this machine
echo.
pause
endlocal
