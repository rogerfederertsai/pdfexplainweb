@echo off
cd /d "%~dp0"

call "%~dp0config.bat"

echo ====================================================
echo   PDF web server starting...
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
if defined LAN_IP (
    echo.
    echo [LAN users] Open in browser:
    echo    http://%LAN_IP%:8000/
    echo.
) else (
    echo Could not auto-detect LAN IPv4. Check Control Panel -^> Network for this PC IP.
)

echo ====================================================
echo Username: %WEB_LOGIN_USERNAME%
echo Password: see config.bat on this machine
echo ====================================================
echo.

cd /d "%~dp0project"
python run_web_launcher.py
