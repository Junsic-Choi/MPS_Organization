@echo off
pushd "%~dp0"
echo ------------------------------------------
echo Port Bypass: Starting Dashboard on 8890
echo ------------------------------------------
node server.js
pause
