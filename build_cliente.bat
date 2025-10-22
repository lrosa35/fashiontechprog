@echo on
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

set LOG=build_log.txt
if exist "%LOG%" del "%LOG%"

where python >>"%LOG%" 2>&1
if errorlevel 1 (
  echo ERRO: Python nao encontrado no PATH. Instale Python 3.11+.
  pause
  exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo Criando venv...>>"%LOG%"
  python -m venv .venv >>"%LOG%" 2>&1 || (echo ERRO ao criar venv & type "%LOG%" & pause & exit /b 1)
)

set PY=".venv\Scripts\python.exe"

echo Atualizando pip...>>"%LOG%"
"%PY%" -m pip install -U pip >>"%LOG%" 2>&1 || (echo ERRO ao atualizar pip & type "%LOG%" & pause & exit /b 1)

echo Instalando dependencias do cliente...>>"%LOG%"
"%PY%" -m pip install -U flet reportlab openpyxl python-docx pillow pyinstaller >>"%LOG%" 2>&1 || (echo ERRO ao instalar dependencias & type "%LOG%" & pause & exit /b 1)

if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__

echo Gerando executavel (onefile)...>>"%LOG%"
"%PY%" -m PyInstaller --clean --noconfirm --onefile --noconsole ^
  --collect-all flet --collect-all flet_core --collect-submodules flet ^
  --name Orcamentos orcamento.py >>"%LOG%" 2>&1
if errorlevel 1 (
  echo ERRO ao gerar o executavel. Veja o log: %LOG%
  type "%LOG%"
  pause
  exit /b 1
)

if not exist dist\Orcamentos.exe (
  echo ERRO: Build finalizou mas o executavel nao foi encontrado. Log:
  type "%LOG%"
  pause
  exit /b 1
)

echo.
echo SUCESSO! Executavel: dist\Orcamentos.exe
for %%A in (dist\Orcamentos.exe) do echo Tamanho: %%~zA bytes

echo.
echo Envie apenas este arquivo ao cliente.
pause
endlocal

