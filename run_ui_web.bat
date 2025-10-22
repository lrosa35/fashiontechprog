@echo off
setlocal

REM Activate venv if exists
if exist .venv\Scripts\activate.bat (
  call .venv\Scripts\activate.bat
)

set PORT=8090
set STORAGE_BACKEND=db
set DATABASE_URL=sqlite:///local.db
set PYTHONIOENCODING=utf-8
echo Iniciando Web UI em http://localhost:%PORT%
start http://localhost:%PORT%
python -m uvicorn ui_app:app --host 0.0.0.0 --port %PORT% --reload
