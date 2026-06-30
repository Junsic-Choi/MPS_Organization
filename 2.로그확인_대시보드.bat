@echo off
rem Foreground Server Start
taskkill /f /im mps_dashboard.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul
echo [INFO] Starting MPS Server in Terminal Mode...
echo [INFO] 브라우저를 통해 대시보드를 엽니다...
start http://localhost:8890
node server.js
pause
