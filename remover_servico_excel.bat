@echo off
setlocal
REM Requer executar como Administrador
set SVC_NAME=AudacesExcelAPI
sc.exe stop %SVC_NAME% >nul 2>&1
sc.exe delete %SVC_NAME% >nul 2>&1
echo Servico %SVC_NAME% removido (se existia).
endlocal
