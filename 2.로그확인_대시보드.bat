@echo off
rem Foreground Server Start
taskkill /f /im mps_dashboard_app.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
ping 127.0.0.1 -n 2 >nul
echo [INFO] Starting MPS Server in Terminal Mode...
echo [INFO] 브라우저를 통해 대시보드를 엽니다...
start http://localhost:8890
node server.js
pause
