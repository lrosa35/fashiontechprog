@echo off
setlocal
REM Auto-elevacao para Administrador
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
cd /d "%~dp0"

set SVC_NAME=AudacesExcelAPI
set SVC_DESC=Servidor FastAPI Excel (OneDrive via Graph)
set PORT=8000

set PYEXE="%CD%\.venv\Scripts\python.exe"
if not exist %PYEXE% (
  echo ERRO: venv nao encontrada em %PYEXE%
  echo Crie a venv: python -m venv .venv
  echo Depois: .venv\Scripts\pip install -U pip uvicorn fastapi httpx msal python-dotenv
  pause
  exit /b 1
)

REM Comando do servico (executa uvicorn com server:app)
set CMD=cmd.exe /c "cd /d %CD% && %PYEXE% -m uvicorn server:app --host 0.0.0.0 --port %PORT%"

sc.exe query %SVC_NAME% >nul 2>&1 && (
  echo Servico %SVC_NAME% ja existe. Remova primeiro com remover_servico_excel.bat
  goto :eof
)

sc.exe create %SVC_NAME% binPath= "%CMD%" start= auto DisplayName= "%SVC_NAME%" >nul 2>&1
sc.exe description %SVC_NAME% "%SVC_DESC%" >nul 2>&1
sc.exe failure %SVC_NAME% reset= 86400 actions= restart/5000/restart/5000/restart/5000 >nul 2>&1

echo Iniciando o servico...
sc.exe start %SVC_NAME%
echo OK.
endlocal
