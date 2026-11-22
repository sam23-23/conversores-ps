$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "1/3 Instalando dependencias (uv, Pandoc, Tesseract OCR, wkhtmltopdf)..." -ForegroundColor Cyan
winget install -e --id astral-sh.uv --accept-source-agreements --accept-package-agreements
winget install -e --id JohnMacFarlane.Pandoc --accept-source-agreements --accept-package-agreements
winget install -e --id UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements
winget install -e --id wkhtmltopdf.wkhtmltox --accept-source-agreements --accept-package-agreements

Write-Host "2/3 Copiando pdf2md.py y conversores.ps1 a $env:USERPROFILE..." -ForegroundColor Cyan
Copy-Item -Path "$scriptDir\pdf2md.py" -Destination "$env:USERPROFILE\pdf2md.py" -Force
Copy-Item -Path "$scriptDir\conversores.ps1" -Destination "$env:USERPROFILE\conversores.ps1" -Force

Write-Host "3/3 Enlazando conversores en el perfil de PowerShell ($PROFILE)..." -ForegroundColor Cyan
$perfilDir = Split-Path -Parent $PROFILE
if (-not (Test-Path $perfilDir)) { New-Item -ItemType Directory -Path $perfilDir -Force | Out-Null }
if (-not (Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }

$lineaCarga = '. "$env:USERPROFILE\conversores.ps1"'
$perfilActual = Get-Content $PROFILE -Raw -ErrorAction SilentlyContinue
if ($perfilActual -notmatch [regex]::Escape($lineaCarga)) {
    Add-Content -Path $PROFILE -Value "`n$lineaCarga" -Encoding UTF8
}

. $PROFILE
Write-Host "¡Listo! Ya tienes 2md, 2docx, 2pdf y 2code instalados sin tocar el resto del perfil." -ForegroundColor Green
