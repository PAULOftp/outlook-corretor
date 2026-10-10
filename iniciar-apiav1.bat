@echo off
cd /d "%~dp0"
start "Antena A.P.IAv1" /min python antena.py
timeout /t 2 >nul
start chrome http://localhost:8080/apiav1.html
