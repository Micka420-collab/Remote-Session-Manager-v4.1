@echo off
:: Lance le script en Admin et cache la fenêtre console bleue (-WindowStyle Hidden)

cd /d "%~dp0"
powershell.exe -Command "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%~dp0RemoteSessionManager.ps1""' -Verb RunAs"
