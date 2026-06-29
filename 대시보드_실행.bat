@chcp 65001 >nul
@echo off
cd /d "%~dp0"
title [MPS Dashboard Server]
cls
echo ==========================================================
echo   MPS Dashboard Server is running.
echo   Please do not close this window.
echo   If you want to stop the server, press Ctrl+C or close this window.
echo ==========================================================
echo.

taskkill /f /im mps_dashboard.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul

if exist mps_dashboard.exe (
    mps_dashboard.exe
) else (
    echo [INFO] mps_dashboard.exe not found. Running with Node.js...
    node server.js
)

echo.
echo [DEBUG] Server process terminated.
pause
