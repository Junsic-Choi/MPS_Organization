@echo off
cd /d "%~dp0"
:: This is the actual server process called by the background script
node server.js > server_start.log 2>&1

