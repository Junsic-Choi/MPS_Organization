@echo off
pushd "%~dp0"
echo [1/3] Stopping ghost processes...
taskkill /F /IM node.exe /T >nul 2>&1
timeout /t 2 /nobreak >nul

echo [2/3] Starting Fixed Server (Port: 8890) in Background...
cscript run_background_8890.vbs

echo [3/3] Success!
echo 1. Open Dashboard: http://localhost:8890/dashboard.html
echo 2. Click [Data Extraction] button once.
echo 3. Verify 'Hutec' LYNX XG800 results.
pause
