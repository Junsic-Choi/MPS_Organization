@echo off
cd /d "%~dp0"
taskkill /f /im mps_dashboard_app.exe >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8890 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1
ping 127.0.0.1 -n 2 >nul

where node >nul 2>&1
if %errorlevel% equ 0 (
    powershell -Command "Start-Process -FilePath 'node' -ArgumentList 'server.js' -WindowStyle Hidden"
) else (
    powershell -Command "Start-Process -FilePath 'mps_dashboard_app.exe' -WindowStyle Hidden"
)

echo [SUCCESS] Server started in background.
echo [INFO] Access URL: http://localhost:8890
ping 127.0.0.1 -n 3 >nul
exit
