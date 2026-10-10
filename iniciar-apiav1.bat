@echo off
cd /d "%~dp0"
rem Antena em segundo plano, sem janela (se ja estiver a correr, sai sozinha)
start "" pythonw antena.py
timeout /t 2 >nul
rem Abre a A.P.IAv1 como aplicacao: janela propria do Microsoft Edge, com a voz Natural de Portugal
start "" msedge --app=http://localhost:8080/apiav1.html
