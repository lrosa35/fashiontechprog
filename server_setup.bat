@echo off
setlocal
cd /d "%~dp0"
echo === Configurando servidor da API Excel (OneDrive) ===

if not exist .env (
  echo ERRO: Arquivo .env nao encontrado. Copie .env.example para .env e preencha TENANT_ID, CLIENT_ID, EXCEL_RELATIVE_PATH, EXCEL_TABLE_NAME.
  pause
  exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo Criando venv...
  python -m venv .venv || (echo ERRO ao criar venv && pause && exit /b 1)
)

echo Instalando dependencias...
".venv\Scripts\pip.exe" install -U pip uvicorn fastapi httpx msal python-dotenv || (echo ERRO ao instalar dependencias && pause && exit /b 1)

call abrir_firewall_8000.bat
call abrir_firewall_udp_discovery.bat

echo Iniciando servidor (primeira vez pode solicitar login pelo Device Code)...
".venv\Scripts\python.exe" -m uvicorn server:app --host 0.0.0.0 --port 8000

endlocal
