@echo off
rem Foreground Server Start
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul
echo [INFO] Starting MPS Server in Terminal Mode...
node server.js
pause
