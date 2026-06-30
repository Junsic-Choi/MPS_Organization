@chcp 65001 >nul
@echo off
cd /d "%~dp0"
title [MPS Dashboard Launcher]
cls
echo ==========================================================
echo   MPS Dashboard Server를 백그라운드에서 실행합니다.
echo   브라우저 창이 실행되면 이 콘솔 창은 자동으로 닫힙니다.
echo ==========================================================
echo.

taskkill /f /im mps_dashboard.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul

where node >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Node.js가 감지되어 백그라운드로 안전하게 실행합니다.
    powershell -Command "Start-Process -FilePath 'node' -ArgumentList 'server.js' -WindowStyle Hidden"
) else (
    if exist mps_dashboard.exe (
        echo [WARNING] Node.js가 없어 mps_dashboard.exe를 실행합니다.
        powershell -Command "Start-Process -FilePath 'mps_dashboard.exe' -WindowStyle Hidden"
    ) else (
        echo [ERROR] 실행 가능한 서버 파일이 없습니다. (node 또는 mps_dashboard.exe 필요)
        pause
        exit
    )
)

echo [INFO] 브라우저를 통해 대시보드를 엽니다...
timeout /t 2 /nobreak >nul
start "" "http://localhost:8890"

echo [SUCCESS] 서버가 백그라운드에서 정상적으로 실행되었습니다.
timeout /t 1 >nul
exit
