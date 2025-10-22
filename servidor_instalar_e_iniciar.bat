@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"

echo === Instalar/Atualizar Servidor Orcamentos (API Excel) ===

if not exist .env (
  if exist .env.example (
    echo Criando .env a partir de .env.example...
    copy /Y .env.example .env >nul
    echo Abra o arquivo .env e preencha TENANT_ID, CLIENT_ID e EXCEL_RELATIVE_PATH.
  ) else (
    echo ATENCAO: Crie um arquivo .env com as variaveis necessarias.
  )
)

where python >nul 2>&1 || (
  echo ERRO: Python nao encontrado no PATH. Instale Python 3.11+ e rode novamente.
  pause & exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo Criando venv...
  python -m venv .venv || (echo ERRO ao criar venv & pause & exit /b 1)
)

set PY=".venv\Scripts\python.exe"

echo Atualizando pip e instalando dependencias do servidor...
%PY% -m pip install -U pip setuptools wheel || (echo ERRO ao atualizar pip & pause & exit /b 1)
%PY% -m pip install -U uvicorn fastapi httpx msal python-dotenv || (echo ERRO ao instalar dependencias & pause & exit /b 1)

echo Abrindo portas no firewall (TCP/8000 e UDP/56789)...
call abrir_firewall_8000.bat >nul 2>&1
call abrir_firewall_udp_discovery.bat >nul 2>&1

echo.
echo Servidor pronto. Opcoes:
echo   [1] Iniciar agora (console)
echo   [2] Instalar como Servico do Windows
echo   [3] Sair
set /p OPT=Escolha [1/2/3]: 
if "%OPT%"=="2" goto :svc
if "%OPT%"=="1" goto :run
goto :end

:svc
call instalar_servico_excel.bat
goto :end

:run
echo Iniciando em http://0.0.0.0:8000 ...
%PY% -m uvicorn server:app --host 0.0.0.0 --port 8000
goto :end

:end
endlocal

