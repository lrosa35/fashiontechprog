@echo off
setlocal
rem Auto-elevacao
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
  echo Criando venv...
  python -m venv .venv || (echo ERRO ao criar venv & pause & exit /b 1)
)
set PY=.venv\Scripts\python.exe

echo Instalando pywin32...
"%PY%" -m pip install -U pywin32 || (echo ERRO ao instalar pywin32 & pause & exit /b 1)

echo Instalando servico Windows AudacesDBAPI...
"%PY%" svc_db.py install
sc.exe config AudacesDBAPI start= auto >nul 2>&1
sc.exe start AudacesDBAPI
sc.exe query AudacesDBAPI

echo Servico instalado. Para parar/remover:
echo   "%PY%" svc_db.py stop
echo   "%PY%" svc_db.py remove
pause
endlocal

