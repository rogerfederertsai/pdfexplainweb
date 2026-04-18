@echo off
cd /d "%~dp0"

:: 載入帳密與環境設定
call "%~dp0config.bat"

echo ====================================================
echo   PDF 解析伺服器 - 全功能啟動（區網 + 外網並用）
echo ====================================================

:: 偵測區網 IP
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
    echo [區網使用者] http://%LAN_IP%:8000/
) else (
    echo [區網] 請手動查詢本機 IP（控制台 → 網路）
)
echo [外網使用者] 請查看新開的 Cloudflare Tunnel 視窗取得網址
echo.
echo ====================================================
echo.

:: 啟動 Web Server（在新視窗，完成後自動關閉）
start "PDF Web Server" cmd /c "cd /d "%~dp0project" && python run_web_launcher.py"

:: 等候伺服器初步就緒
timeout /t 4 /nobreak >nul

:: 啟動 Cloudflare Tunnel（在新視窗，保持開啟以顯示外網 URL）
start "Cloudflare Tunnel" cmd /k "cloudflared tunnel --url http://localhost:8000"

echo 兩個服務已啟動！
echo   - PDF Web Server：在背景執行（如需重啟，請再執行此批次檔）
echo   - Cloudflare Tunnel：請查看新開的 [Cloudflare Tunnel] 視窗
echo.
echo 帳號：%WEB_LOGIN_USERNAME%
echo 密碼：（請查看 config.bat）
echo.
pause
