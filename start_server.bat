@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

call "%~dp0config.bat"

echo ====================================================
echo   PDF web server starting...
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

if defined LAN_IP (
    echo.
    echo [LAN users] Open in browser:
    echo    http://!LAN_IP!:8000/
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
endlocal
