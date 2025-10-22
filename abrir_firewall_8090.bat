@echo off
setlocal
rem Elevar para Administrador se necessario
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
echo Criando regra de firewall para liberar TCP/8090 (downloads)...
netsh advfirewall firewall add rule name="Orcamentos Downloads 8090" dir=in action=allow protocol=TCP localport=8090
echo Concluido.
endlocal
