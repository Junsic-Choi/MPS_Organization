@echo off
echo Stopping MPS Dashboard Server on port 8888...
powershell -Command "$p = Get-NetTCPConnection -LocalPort 8888 -ErrorAction SilentlyContinue; if ($p) { Stop-Process -Id $p.OwningProcess -Force; echo 'Server stopped.' } else { echo 'Server is not running.' }"
pause
