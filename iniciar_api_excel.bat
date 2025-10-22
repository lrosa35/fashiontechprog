@echo off
setlocal
REM Inicia a API Excel (server:app) usando a venv local
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
  echo ERRO: venv nao encontrada em .venv\Scripts\python.exe
  echo Crie a venv e instale dependencias: python -m venv .venv ^&^& .venv\Scripts\pip install -U pip uvicorn fastapi httpx msal python-dotenv
  pause
  exit /b 1
)

set PYTHONUTF8=1
".venv\Scripts\python.exe" -m uvicorn server:app --host 0.0.0.0 --port 8000
endlocal
