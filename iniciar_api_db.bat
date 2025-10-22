@echo off
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo ERRO: venv nao encontrada. Rode server_db_setup.bat primeiro.
  pause & exit /b 1
)
set PY=".venv\Scripts\python.exe"
%PY% -m uvicorn server_db:app --host 0.0.0.0 --port 8000
endlocal

