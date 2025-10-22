@echo off
setlocal
chcp 65001 >nul
set OD=%OneDriveCommercial%
if "%OD%"=="" set OD=%OneDrive%
if "%OD%"=="" (
  echo Variavel OneDrive nao encontrada. Abra o arquivo via Explorador e marque "Sempre manter neste dispositivo".
  pause & exit /b 1
)
set FILE="%OD%\01 LEANDRO\IMPRESSÕES\BANCO_DE_DADOS_ORCAMENTO.xlsx"
echo Fixando em "Sempre manter neste dispositivo":
echo   %FILE%
attrib +P -U %FILE%
if errorlevel 1 (
  echo Nao foi possivel aplicar o atributo. Tente clicar com o botao direito no arquivo no OneDrive e escolha "Manter neste dispositivo".
) else (
  echo OK. Caso o arquivo esteja apenas online, aguarde baixar e tente abrir o app novamente.
)
pause
endlocal

