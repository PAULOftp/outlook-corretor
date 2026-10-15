@echo off
cd /d "%~dp0"
rem Fecha antenas antigas da A.P.IAv1 que tenham ficado a correr (so processos python com antena.py)
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -like 'python*' -and $_.CommandLine -like '*antena.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1
rem Antena em segundo plano, sem janela
start "" pythonw antena.py
timeout /t 2 >nul
rem Abre a A.P.IAv1 como aplicacao: janela propria do Microsoft Edge, com a voz Natural de Portugal
start "" msedge --app=http://localhost:8080/apiav1.html
