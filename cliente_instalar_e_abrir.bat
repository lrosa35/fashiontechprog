@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"

title Orcamentos - Instalador do Cliente

if not exist Orcamentos.exe (
  echo ERRO: Orcamentos.exe nao encontrado nesta pasta.
  echo Copie Orcamentos.exe para esta mesma pasta e rode novamente.
  pause
  exit /b 1
)

echo ====== Configurar URL da API ======
set DEFAULT_URL=
if exist cloud_api_url.txt (
  for /f "usebackq delims=" %%A in ("cloud_api_url.txt") do set DEFAULT_URL=%%A
)
echo URL atual: %DEFAULT_URL%
set /p API_URL=Informe a URL da API [ex.: http://SEU_IP:8000] (Enter para manter atual): 
if not "%API_URL%"=="" set DEFAULT_URL=%API_URL%
if not "%DEFAULT_URL%"=="" (
  echo %DEFAULT_URL%> cloud_api_url.txt
  echo Salvo em cloud_api_url.txt: %DEFAULT_URL%
)

echo ====== Verificando/Instalando WebView2 Runtime ======
REM Muitos PCs ja possuem. Tentamos instalar/update silencioso.
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $url='https://go.microsoft.com/fwlink/p/?LinkId=2124703'; $tmp=Join-Path $env:TEMP 'wv2setup.exe'; Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing; Start-Process -FilePath $tmp -ArgumentList '/silent','/install' -Wait; 'WebView2 OK' } catch { 'Falha ao baixar/instalar WebView2: ' + $_ }" 1>nul 2>nul

echo ====== Iniciando o programa ======
start "" "%CD%\Orcamentos.exe"
echo Se nada abrir, verifique o firewall ou permissao do Windows SmartScreen.
pause
endlocal

