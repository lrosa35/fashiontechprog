@echo off
setlocal
cd /d "%~dp0"
echo === Configurando servidor DB (SQLite/Postgres) ===

if not exist .venv\Scripts\python.exe (
  echo Criando venv...
  python -m venv .venv || (echo ERRO ao criar venv & pause & exit /b 1)
)

echo Instalando dependencias...
".venv\Scripts\pip.exe" install -U pip uvicorn fastapi python-dotenv sqlalchemy || (echo ERRO & pause & exit /b 1)

if not exist .env (
  if exist .env.example copy /Y .env.example .env >nul
)
if exist .env (
  powershell -NoProfile -Command "(Get-Content .env) -replace '^STORAGE_BACKEND=.*','STORAGE_BACKEND=db' | Set-Content .env"
)

call abrir_firewall_8000.bat

echo Iniciando servidor DB em http://0.0.0.0:8000 ...
".venv\Scripts\python.exe" -m uvicorn server_db:app --host 0.0.0.0 --port 8000

endlocal

