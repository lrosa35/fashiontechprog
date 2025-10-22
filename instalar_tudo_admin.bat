@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul

rem Auto-elevar
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)

cd /d "%~dp0"
echo === Instalacao completa (Servidor + Regras + Build) ===

where python >nul 2>&1 || (
  echo ERRO: Python nao encontrado. Instale Python 3.11+ e rode novamente.
  pause & exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo Criando venv...
  python -m venv .venv || (echo ERRO ao criar venv & pause & exit /b 1)
)

set PY=".venv\Scripts\python.exe"
%PY% -m pip install -U pip setuptools wheel || (echo ERRO ao atualizar pip & pause & exit /b 1)

echo Instalando dependencias do cliente e servidor...
%PY% -m pip install -U flet reportlab openpyxl python-docx pillow pyinstaller ^
  uvicorn fastapi httpx msal python-dotenv || (echo ERRO ao instalar deps & pause & exit /b 1)

echo Abrindo portas de firewall...
call abrir_firewall_8000.bat
call abrir_firewall_udp_discovery.bat
call abrir_firewall_8090.bat

if not exist .env (
  if exist .env.example copy /Y .env.example .env >nul
  echo Abra o arquivo .env e preencha TENANT_ID, CLIENT_ID e EXCEL_RELATIVE_PATH.
  start notepad.exe .env
  pause
)

echo Deseja instalar a API como Servico do Windows agora?
set /p ANS=[S/N]: 
if /I "%ANS%"=="S" call instalar_servico_excel.bat

echo Gerando executavel do cliente...
call build_cliente.bat

echo Opcional: servir a pasta dist via HTTP para download na rede local.
echo Deseja iniciar o servidor de downloads (porta 8090)?
set /p ANS2=[S/N]: 
if /I "%ANS2%"=="S" call servir_downloads.bat

echo Concluido.
pause
endlocal

