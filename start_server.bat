@echo off
cd /d "%~dp0"
powershell -Command "$p = Get-NetTCPConnection -LocalPort 8888 -ErrorAction SilentlyContinue; if ($p) { Stop-Process -Id $p.OwningProcess -Force }" 2>NUL
node server.js > server_execution.log 2>&1
