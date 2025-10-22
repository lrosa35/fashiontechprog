@echo off
setlocal

REM Activate venv if exists
if exist .venv\Scripts\activate.bat (
  call .venv\Scripts\activate.bat
)

set PORT=8090
echo Iniciando Web UI em http://localhost:%PORT%
start http://localhost:%PORT%
python -m uvicorn ui_app:app --host 0.0.0.0 --port %PORT% --reload
