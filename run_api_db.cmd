@echo off
cd /d "%~dp0"
"%~dp0.venv\Scripts\python.exe" -m uvicorn server_db:app --host 0.0.0.0 --port 8000

