@echo off
setlocal ENABLEDELAYEDEXPANSION
title Migrar Orçamentos para pasta local
chcp 65001 >nul

set SRC="%~dp0"
set DEFAULT_DST=C:\OrcamentosApp
set /p DST=Destino para instalar localmente [C:\OrcamentosApp]: 
if "%DST%"=="" set DST=%DEFAULT_DST%

echo.
echo Criando pasta destino: "%DST%"
mkdir "%DST%" 2>nul

echo.
echo Copiando arquivos (ignorando pastas temporarias e venv)...
rem Requer Windows Vista+; usa ROBOCOPY para copiar com filtros
robocopy %SRC% "%DST%" ^
  /E ^
  /XD .git .venv __pycache__ build dist antigo ^
  /XF orcamento_copy.txt __check_compile.py install_deps.ps1 instalar_exe.bat ^
  /R:2 /W:1 >nul

if errorlevel 8 (
  echo ATENCAO: Robocopy retornou codigo %errorlevel%. Verifique a copia.
)

echo.
echo Concluido.
echo Agora execute:
echo   1) "%DST%\server_setup.bat"  ^<-- no seu PC (servidor) para subir a API Excel
echo   2) "%DST%\build_cliente.bat" ^<-- para gerar o EXE e enviar ao cliente

echo.
echo Abrindo a pasta destino...
start "" "%DST%"
pause
endlocal

