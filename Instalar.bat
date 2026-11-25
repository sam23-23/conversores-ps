@echo off
chcp 65001 >nul
title Instalador Suite Conversores PowerShell (2md, 2docx, 2pdf, 2code)
echo ============================================================================
echo   INSTALANDO SUITE DE CONVERSIONES (2md, 2docx, 2pdf, 2code)
echo ============================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
if %ERRORLEVEL% NEQ 0 (
    echo [!] Hubo un aviso o error durante la instalacion. Revisa los mensajes de arriba.
) else (
    echo [OK] Instalacion completada. Ya puedes cerrar esta ventana y abrir PowerShell.
)
pause