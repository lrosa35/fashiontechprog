@echo off
setlocal
rem Elevar para Administrador se necessario
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
echo Criando regra de firewall para liberar TCP/8000...
netsh advfirewall firewall add rule name="Audaces Excel API 8000" dir=in action=allow protocol=TCP localport=8000
echo Concluido.
endlocal
