@echo off
cd /d "%~dp0"
rem Background Server Start
taskkill /f /im mps_dashboard.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul

where node >nul 2>&1
if %errorlevel% equ 0 (
    powershell -Command "Start-Process -FilePath 'node' -ArgumentList 'server.js' -WindowStyle Hidden"
) else (
    powershell -Command "Start-Process -FilePath 'mps_dashboard.exe' -WindowStyle Hidden"
)

echo [SUCCESS] Server started in background.
echo [INFO] Access URL: http://localhost:8890
timeout /t 2 >nul
exit
