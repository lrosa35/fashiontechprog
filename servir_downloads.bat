@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"

set PORT=8090
set DIR=dist

if not exist "%DIR%" (
  echo Pasta "%DIR%" não encontrada. Gere primeiro o executável com build_cliente.bat.
  pause
  exit /b 1
)

call abrir_firewall_8090.bat

for /f "tokens=14" %%I in ('ipconfig ^| findstr /R /C:"IPv4.*:"') do set IP=%%I
echo.
echo Servindo "%DIR%" em http://%IP%:%PORT%/  (acesso na rede local)
echo Pressione Ctrl+C para encerrar.

set PY=python
if exist .venv\Scripts\python.exe set PY=.venv\Scripts\python.exe

"%PY%" -m http.server %PORT% --directory "%DIR%"
endlocal

