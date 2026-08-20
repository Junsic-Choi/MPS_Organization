@echo off
cd /d "%~dp0"
title [Namsan MC Shopfloor Board Launcher]

for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8895 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1
ping 127.0.0.1 -n 2 >nul

if exist "%~dp0node.exe" (
    start "MC_Shopfloor_Server_8895" "%~dp0node.exe" shopfloor_server.js
) else (
    start "MC_Shopfloor_Server_8895" node shopfloor_server.js
)

ping 127.0.0.1 -n 2 >nul
start "" "http://localhost:8895"
exit
