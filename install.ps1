$ErrorActionPreference = "Continue"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  Instalador Automático - Suite Conversores PS" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

Write-Host "`n[1/4] Instalando dependencias (uv, Pandoc, Tesseract OCR, wkhtmltopdf)..." -ForegroundColor Cyan
$paquetes = @(
    "astral-sh.uv",
    "JohnMacFarlane.Pandoc",
    "UB-Mannheim.TesseractOCR",
    "wkhtmltopdf.wkhtmltox"
)
foreach ($pkg in $paquetes) {
    winget install -e --id $pkg --accept-source-agreements --accept-package-agreements --silent | Out-Null
}

# Recargar PATH del sistema y usuario en la sesión actual
$env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")

Write-Host "[2/4] Configurando PATH y paquete de idioma español para OCR..." -ForegroundColor Cyan
$rutasExtra = @(
    "C:\Program Files\Tesseract-OCR",
    "C:\Program Files (x86)\Tesseract-OCR",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR",
    "C:\Program Files\wkhtmltopdf\bin",
    "$env:LOCALAPPDATA\Programs\wkhtmltopdf\bin",
    "$env:USERPROFILE\.local\bin"
) | Where-Object { Test-Path $_ }

$pathUsuario = [Environment]::GetEnvironmentVariable("Path", "User")
if (-not $pathUsuario) { $pathUsuario = "" }
$pathUsuario = $pathUsuario -replace ';C$', ''

foreach ($ruta in $rutasExtra) {
    if ($pathUsuario -notlike "*$ruta*") {
        $pathUsuario = ($pathUsuario.TrimEnd(';') + ";" + $ruta).TrimStart(';')
    }
    if ($env:Path -notlike "*$ruta*") {
        $env:Path += ";$ruta"
    }
}
[Environment]::SetEnvironmentVariable("Path", $pathUsuario, "User")

# Descargar spa.traineddata si Tesseract está instalado y no lo tiene
$rutaTesseract = @(
    "C:\Program Files\Tesseract-OCR",
    "C:\Program Files (x86)\Tesseract-OCR",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($rutaTesseract) {
    $tessdataDir = Join-Path $rutaTesseract "tessdata"
    $spaFile = Join-Path $tessdataDir "spa.traineddata"
    if ((Test-Path $tessdataDir) -and (-not (Test-Path $spaFile))) {
        Write-Host "      Descargando modelo OCR en español (spa.traineddata)..." -ForegroundColor DarkCyan
        try {
            Invoke-WebRequest -Uri "https://github.com/tesseract-ocr/tessdata_fast/raw/main/spa.traineddata" -OutFile $spaFile -UseBasicParsing
        } catch {
            Write-Host "      Aviso: No se pudo descargar spa.traineddata (se usará inglés por defecto)." -ForegroundColor Yellow
        }
    }
}

Write-Host "[3/4] Copiando y desbloqueando scripts en $env:USERPROFILE..." -ForegroundColor Cyan
$destPy = Join-Path $env:USERPROFILE "pdf2md.py"
$destPs1 = Join-Path $env:USERPROFILE "conversores.ps1"

Copy-Item -Path (Join-Path $scriptDir "pdf2md.py") -Destination $destPy -Force
Copy-Item -Path (Join-Path $scriptDir "conversores.ps1") -Destination $destPs1 -Force
Unblock-File -Path $destPy, $destPs1 -ErrorAction SilentlyContinue

Write-Host "[4/4] Sustituyendo/añadiendo solo las funciones de esta suite en `$PROFILE..." -ForegroundColor Cyan
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force -ErrorAction SilentlyContinue

$perfilDir = Split-Path -Parent $PROFILE
if (-not (Test-Path $perfilDir)) { New-Item -ItemType Directory -Path $perfilDir -Force | Out-Null }
if (-not (Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }

# Nombres exactos de las funciones y alias gestionados por esta suite
$funcionesGestionadas = @(
    "Buscar-ArchivoLocal",
    "2md",
    "convertirmd",
    "pdf2md",
    "Reemplazar-TextoWord",
    "Convertir-ConPlantillaWord",
    "2docx",
    "2pdf",
    "2code"
)

$contenidoActual = Get-Content $PROFILE -Raw -ErrorAction SilentlyContinue
if ($null -eq $contenidoActual) { $contenidoActual = "" }

if ($contenidoActual.Trim().Length -gt 0) {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($contenidoActual, [ref]$tokens, [ref]$errors)

    # Localizar bloques AST de funciones/filtros y llamadas Set-Alias que coincidan con nuestros nombres
    $rangosEliminar = @()

    $funcDefs = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        ($funcionesGestionadas -contains $node.Name)
    }, $true)

    foreach ($fn in $funcDefs) {
        $rangosEliminar += [PSCustomObject]@{
            Start  = $fn.Extent.StartOffset
            End    = $fn.Extent.EndOffset
            Nombre = $fn.Name
        }
    }

    $aliasCmds = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.CommandAst] -and
        ($node.GetCommandName() -eq "Set-Alias") -and
        ($node.Extent.Text -match '\b(convertirmd|pdf2md|2md|2docx|2pdf|2code)\b')
    }, $true)

    foreach ($al in $aliasCmds) {
        $rangosEliminar += [PSCustomObject]@{
            Start  = $al.Extent.StartOffset
            End    = $al.Extent.EndOffset
            Nombre = "Alias"
        }
    }

    # Eliminar los bloques antiguos de atrás hacia adelante para conservar los índices exactos
    $rangosOrdenados = $rangosEliminar | Sort-Object Start -Descending
    foreach ($r in $rangosOrdenados) {
        Write-Host "      Sobrescribiendo definición previa detectada: $($r.Nombre)" -ForegroundColor Yellow
        $contenidoActual = $contenidoActual.Remove($r.Start, $r.End - $r.Start)
    }

    # Eliminar líneas antiguas de carga de conversores.ps1 para evitar duplicados
    $lineasRestantes = ($contenidoActual -split "`r?`n") | Where-Object {
        $_ -notmatch 'conversores\.ps1'
    }
    $contenidoLimpio = ($lineasRestantes -join "`r`n").Trim()
} else {
    $contenidoLimpio = ""
}

$lineaCarga = 'Invoke-Expression (Get-Content "$env:USERPROFILE\conversores.ps1" -Raw -Encoding UTF8)'
if ($contenidoLimpio.Length -gt 0) {
    $nuevoPerfil = $contenidoLimpio + "`r`n`r`n" + $lineaCarga + "`r`n"
} else {
    $nuevoPerfil = $lineaCarga + "`r`n"
}

Set-Content -Path $PROFILE -Value $nuevoPerfil -Encoding UTF8
Unblock-File -Path $PROFILE -ErrorAction SilentlyContinue

# Cargar en memoria inmediatamente
Invoke-Expression (Get-Content $destPs1 -Raw -Encoding UTF8)

Write-Host "`n¡Instalación completada con éxito!" -ForegroundColor Green
Write-Host "Funciones sustituidas/añadidas: 2md, 2docx, 2pdf, 2code (el resto de tu perfil sigue intacto)." -ForegroundColor Green
