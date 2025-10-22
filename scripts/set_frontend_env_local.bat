@echo off
setlocal
rem Gera frontend/js/config.js apontando para backend local
set API=http://localhost:8000
echo window.API_BASE="%API%";> "%~dp0..\frontend\js\config.js"
echo Gerado frontend\js\config.js com %API%
endlocal
