@echo off
setlocal
rem Elevar para Administrador se necessario
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
echo Liberando porta UDP 56789 para descoberta...
netsh advfirewall firewall add rule name="Audaces Discovery 56789" dir=in action=allow protocol=UDP localport=56789
echo Concluido.
endlocal
