@echo off
pushd "%~dp0"
echo [1/2] Cleaning up old processes...
:: Kill all node processes silently, ignore errors if none found
taskkill /F /IM node.exe /T >nul 2>&1
timeout /t 2 /nobreak >nul

echo [2/2] Starting New Dashboard Server...
echo Directory: %CD%
node server.js
pause
