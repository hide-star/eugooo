@echo off
title EUGOO Server Launcher

:: Build HTML pages first
cd /d "%~dp0"
echo [1/4] Building HTML pages...
python build_v7_v3.py
echo.

:: Get WLAN IP
for /f "tokens=2 delims=:" %%a in ('netsh interface ip show address "WLAN" ^| findstr "IP "') do set LOCAL_IP=%%a
set LOCAL_IP=%LOCAL_IP: =%

:: Start Web Server
start "EUGOO-Web" cmd /k "cd /d "%~dp0server" && node server.js"

timeout /t 2 >nul

:: Start Sync Service
start "EUGOO-Sync" cmd /k "cd /d "%~dp0" && python sync_wecom.py"

timeout /t 2 >nul

echo.
echo ============================================
echo   Services started!
echo.
echo   Local:   http://localhost:3000
if defined LOCAL_IP (
echo   LAN:     http://%LOCAL_IP%:3000
)
echo.
echo   Close the three cmd windows to stop.
echo ============================================
echo.
pause
