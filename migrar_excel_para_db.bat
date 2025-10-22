@echo off
setlocal
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
  echo ERRO: venv nao encontrada. Rode server_db_setup.bat primeiro.
  pause & exit /b 1
)

set PY=".venv\Scripts\python.exe"
if "%~1"=="" (
  %PY% migrate_from_excel.py
) else (
  %PY% migrate_from_excel.py "%~1"
)

echo.
echo Migracao concluida (veja contagem acima).
pause
endlocal

