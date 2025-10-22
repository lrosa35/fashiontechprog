@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0.."

set "REPO_URL=https://github.com/lrosa35/fashiontechprog.git"
set "BRANCH=main"

where git >nul 2>&1 || (echo [error] Git not found. Install Git and retry.& exit /b 1)

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo [git] Initializing local repository...
  git init || exit /b 1
  git checkout -B %BRANCH% 1>nul 2>nul
)

git symbolic-ref --short HEAD | findstr /i ^%BRANCH%$ >nul || git checkout -B %BRANCH%

git remote get-url origin >nul 2>&1 && (
  echo [git] Updating remote origin to: %REPO_URL%
  git remote set-url origin %REPO_URL% || exit /b 1
) || (
  echo [git] Adding remote origin: %REPO_URL%
  git remote add origin %REPO_URL% || exit /b 1
)

rem Ensure local author is set to allow commit
for /f "tokens=*" %%A in ('git config user.name 2^>nul') do set _HAS_NAME=%%A
for /f "tokens=*" %%A in ('git config user.email 2^>nul') do set _HAS_EMAIL=%%A
if not defined _HAS_NAME git config user.name "lrosa35"
if not defined _HAS_EMAIL git config user.email "no-reply@example.com"

echo [git] Adding files...
git add -A || exit /b 1

set "NOW=%date% %time%"
set "MSG=deploy: publish to GitHub - %NOW%"

git diff --cached --quiet && (
  echo [git] Nothing to commit.
) || (
  git commit -m "%MSG%" || exit /b 1
)

echo [git] Pushing to origin/%BRANCH% ...
git push -u origin %BRANCH%
if errorlevel 1 (
  echo [error] Push failed. Check GitHub credentials or token.
  exit /b 1
)

echo [ok] Published to %REPO_URL% (branch %BRANCH%).
exit /b 0
