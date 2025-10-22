@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0.."

rem === Config ===
set REPO_URL=https://github.com/lrosa35/fashiontechprog.git
set TARGET_DIR=fashiontechprog
set BRANCH=main

echo Atualizando repositório: %REPO_URL%

where git >nul 2>&1
if %ERRORLEVEL%==0 (
  if not exist "%TARGET_DIR%\.git" (
    echo [git] Clonando em %TARGET_DIR% ...
    git clone "%REPO_URL" "%TARGET_DIR%" || goto :try_zip
  ) else (
    echo [git] Já existe .git em %TARGET_DIR%, atualizando...
    pushd "%TARGET_DIR%"
    git fetch --all || goto :try_zip_pop
    git checkout %BRANCH% || goto :try_zip_pop
    git pull --ff-only origin %BRANCH% || goto :try_zip_pop
    popd
  )
  echo [ok] Repositório pronto em %TARGET_DIR%
  goto :eof
)

:try_zip
echo [info] Git não disponível ou falhou. Baixando ZIP do GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar_github.ps1" -RepoUrl "%REPO_URL" -TargetDir "%TARGET_DIR" -Branch "%BRANCH%"
if %ERRORLEVEL% NEQ 0 (
  echo [erro] Falha ao baixar e extrair o ZIP.
  exit /b 1
)
echo [ok] Repositório pronto em %TARGET_DIR%
exit /b 0

:try_zip_pop
  popd
  goto :try_zip

endlocal

