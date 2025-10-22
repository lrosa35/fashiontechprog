@echo off
setlocal
rem Gera frontend/js/config.js apontando para backend no Heroku
if "%1"=="" (
  echo Uso: %~n0 https://SEU-HEROKU-APP.herokuapp.com
  exit /b 1
)
set API=%~1
echo window.API_BASE="%API%";> "%~dp0..\frontend\js\config.js"
echo Gerado frontend\js\config.js com %API%
endlocal
