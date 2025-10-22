@echo off
setlocal
rem Auto-elevacao
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
cd /d "%~dp0"
set SVC_NAME=AudacesDBAPI
sc.exe stop %SVC_NAME% >nul 2>&1
sc.exe delete %SVC_NAME% >nul 2>&1
echo Servico %SVC_NAME% removido (se existia).
endlocal

