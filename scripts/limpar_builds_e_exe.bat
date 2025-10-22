@echo off
chcp 65001 >nul
echo [info] Removendo builds/EXE locais (n?o afeta o GitHub)
for %%D in (dist build antigo) do (
  if exist "%%D" (
    echo - Apagando pasta %%D ...
    rmdir /s /q "%%D"
  )
)
for %%F in (*.exe *.spec) do (
  if exist "%%F" (
    echo - Apagando arquivo %%F
    del /q "%%F"
  )
)
echo [ok] Limpeza conclu?da.
exit /b 0
