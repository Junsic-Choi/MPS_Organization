@echo off
cd /d "%~dp0"
rem Background Server Start
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul
wscript.exe run_dashboard_silent.vbs
echo [SUCCESS] Server started in background.
echo [INFO] Access URL: http://localhost:8890
timeout /t 3 >nul
exit
