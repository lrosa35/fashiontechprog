@echo off
setlocal
cd /d "%~dp0"

set TASK=AudacesDBAPI
set RUN="%CD%\run_api_db.cmd"

REM Cria tarefa para iniciar na inicializacao do sistema como SYSTEM
schtasks /Create /TN %TASK% /SC ONSTART /RU SYSTEM /TR %RUN% /F
echo Iniciando tarefa...
schtasks /Run /TN %TASK%
echo Status atual:
schtasks /Query /TN %TASK% /V /FO LIST
endlocal

