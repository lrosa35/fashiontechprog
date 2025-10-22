@echo off
setlocal

if not exist .venv (
  echo Criando ambiente virtual...
  py -3 -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo Ambiente pronto. Para iniciar a UI: run_ui_web.bat

