@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"

REM Wrapper simples chamando o script existente
call build_cliente.bat

endlocal

