@echo off
cd /d "%~dp0"
start "Antena PauloIA" /min python antena.py
timeout /t 2 >nul
start chrome http://localhost:8080/pauloia.html
