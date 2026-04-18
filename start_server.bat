@echo off
cd /d "%~dp0"

:: 載入帳密與環境設定
call "%~dp0config.bat"

echo ====================================================
echo   PDF 解析伺服器啟動中...
echo ====================================================

:: 自動偵測本機區網 IP 並顯示給使用者
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
    echo [同區網使用者] 請在瀏覽器輸入：
    echo    http://%LAN_IP%:8000/
    echo.
) else (
    echo 無法自動偵測區網 IP，請至「控制台 → 網路」查詢。
)

echo ====================================================
echo 帳號：%WEB_LOGIN_USERNAME%
echo 密碼：（請查看 config.bat）
echo ====================================================
echo.

cd /d "%~dp0project"
python run_web_launcher.py
