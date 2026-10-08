@echo off
cd /d "%~dp0"
start "Servidor PauloIA" /min python -m http.server 8080 --bind 127.0.0.1
timeout /t 2 >nul
start chrome http://localhost:8080/rissa.html
