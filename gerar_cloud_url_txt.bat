@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"

set PORT=8000
for /f "tokens=14" %%I in ('ipconfig ^| findstr /R /C:"IPv4.*:"') do set IP=%%I
set URL=http://%IP%:%PORT%
echo %URL%> cloud_api_url.txt
echo Gerado cloud_api_url.txt com: %URL%
echo Inclua este arquivo ao lado do executável para pré-preencher a URL.
pause
endlocal

