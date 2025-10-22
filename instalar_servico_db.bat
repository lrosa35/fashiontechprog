@echo off
setlocal
rem Auto-elevacao
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
cd /d "%~dp0"

set SVC_NAME=AudacesDBAPI
set SVC_DESC=Servidor FastAPI (DB local via SQLAlchemy)
set PORT=8000
set PYEXE=%CD%\.venv\Scripts\python.exe

if not exist %PYEXE% (
  echo ERRO: venv nao encontrada. Rode server_db_setup.bat primeiro.
  pause & exit /b 1
)

set RUN=%CD%\run_api_db.cmd

sc.exe create %SVC_NAME% binPath= "%RUN%" start= auto DisplayName= "%SVC_NAME%"
if errorlevel 1 goto :err
sc.exe description %SVC_NAME% "%SVC_DESC%"
sc.exe failure %SVC_NAME% reset= 86400 actions= restart/5000/restart/5000/restart/5000
sc.exe start %SVC_NAME%
echo Servico %SVC_NAME% instalado e iniciado.
goto :eof

:err
echo Falha ao criar o servico. Verifique se ja existe ou se o comando esta correto.
sc.exe query %SVC_NAME%
endlocal
