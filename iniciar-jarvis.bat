@echo off
cd /d "%~dp0"
start "Antena do Jarvis" cmd /k node server.js
timeout /t 2 >nul
start "" "%~dp0jarvis.html"
