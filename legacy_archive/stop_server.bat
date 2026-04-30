@echo off
echo Stopping existing Node.js server processes...
taskkill /F /IM node.exe /T
echo Done. You can now run the new server.
pause
