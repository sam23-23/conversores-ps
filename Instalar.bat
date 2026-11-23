@echo off
title Instalador Suite Conversores PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
echo Ya puedes cerrar esta ventana y abrir PowerShell para usar 2md, 2docx, 2pdf y 2code.
pause
