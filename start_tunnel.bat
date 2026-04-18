@echo off
cd /d "%~dp0"

echo ====================================================
echo   Cloudflare Tunnel 啟動中（外網連線用）
echo ====================================================
echo.
echo   完成後，下方會顯示一個 https://xxxx.trycloudflare.com 網址。
echo   將此網址傳給外網使用者，他們即可連線。
echo.
echo   注意：此視窗關閉後，外網連線即中斷。
echo         區網使用者不受影響，仍可直接用區網 IP 連線。
echo ====================================================
echo.

:: 檢查 cloudflared 是否已安裝
where cloudflared >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [錯誤] 找不到 cloudflared 指令！
    echo.
    echo 請先下載並安裝 cloudflared：
    echo   1. 前往 https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
    echo   2. 下載 Windows 版（cloudflared-windows-amd64.exe）
    echo   3. 將檔案改名為 cloudflared.exe 並放到 C:\Windows\System32\ 或任意 PATH 目錄
    echo.
    pause
    exit /b 1
)

cloudflared tunnel --url http://localhost:8000
pause
