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

wscript.exe run_dashboard_silent.vbs

echo [SUCCESS] 서버가 백그라운드에서 정상적으로 실행되었습니다.
timeout /t 2 >nul
exit
