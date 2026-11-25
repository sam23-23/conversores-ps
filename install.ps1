#Requires -Version 5.1
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$totalPasos = 6
$dirOrigen = Split-Path -Parent $MyInvocation.MyCommand.Path
$utf8Bom = New-Object System.Text.UTF8Encoding $true

function Mostrar-Progreso {
    param([int]$Paso, [string]$Actividad, [string]$Estado)
    $porcentaje = [math]::Round(($Paso / $totalPasos) * 100)
    Write-Progress -Activity $Actividad -Status $Estado -PercentComplete $porcentaje
    Write-Host "[$Paso/$totalPasos] $Estado" -ForegroundColor Cyan
}

# 1. Comprobacion de restricciones de PowerShell (ordenadores de instituto / GPO)
Mostrar-Progreso -Paso 1 -Actividad "Instalando conversores-ps" -Estado "Comprobando politicas de ejecucion de PowerShell..."

Get-ChildItem -Path $dirOrigen -File | ForEach-Object {
    try { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue } catch {}
}

$politicas = Get-ExecutionPolicy -List
$gpoMachine = ($politicas | Where-Object { $_.Scope -eq "MachinePolicy" }).ExecutionPolicy
$gpoUser    = ($politicas | Where-Object { $_.Scope -eq "UserPolicy" }).ExecutionPolicy
$bloqueadoPorAdmin = $false

if (($gpoMachine -in @("Restricted", "AllSigned")) -or ($gpoUser -in @("Restricted", "AllSigned"))) {
    $bloqueadoPorAdmin = $true
    Write-Host "  [AVISO] Este equipo tiene una directiva de grupo restrictiva ($gpoMachine / $gpoUser)." -ForegroundColor Yellow
} else {
    try {
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force -ErrorAction Stop
        Write-Host "  [OK] Politica CurrentUser configurada en RemoteSigned." -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -notmatch "ExecutionPolicyOverride") {
            Write-Host "  [AVISO] No se pudo cambiar la politica automaticamente: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# 2. Instalacion de dependencias mediante winget (solo si faltan)
Mostrar-Progreso -Paso 2 -Actividad "Instalando conversores-ps" -Estado "Verificando dependencias (uv, Pandoc, Tesseract, wkhtmltopdf)..."

$dependencias = @(
    @{ Id = "astral-sh.uv";             Nombre = "uv (Python)";   Check = { $null -ne (Get-Command uv -ErrorAction SilentlyContinue) } },
    @{ Id = "JohnMacFarlane.Pandoc";    Nombre = "Pandoc";        Check = { ($null -ne (Get-Command pandoc -ErrorAction SilentlyContinue)) -or (Test-Path "$env:LOCALAPPDATA\Pandoc\pandoc.exe") -or (Test-Path "C:\Program Files\Pandoc\pandoc.exe") } },
    @{ Id = "UB-Mannheim.TesseractOCR"; Nombre = "Tesseract OCR"; Check = { ($null -ne (Get-Command tesseract -ErrorAction SilentlyContinue)) -or (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe") -or (Test-Path "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe") } },
    @{ Id = "wkhtmltopdf.wkhtmltox";    Nombre = "wkhtmltopdf";   Check = { ($null -ne (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue)) -or (Test-Path "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe") } }
)

$tieneWinget = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)

foreach ($dep in $dependencias) {
    if (& $dep.Check) {
        Write-Host "  [OK] $($dep.Nombre) ya esta instalado." -ForegroundColor DarkGreen
    } elseif ($tieneWinget) {
        Write-Host "  [..] Instalando $($dep.Nombre)..." -ForegroundColor Gray
        winget install --id $dep.Id -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
    } else {
        Write-Host "  [AVISO] winget no disponible. Instala manualmente $($dep.Nombre) ($($dep.Id))." -ForegroundColor Yellow
    }
}

# 3. Configuracion del PATH del usuario
Mostrar-Progreso -Paso 3 -Actividad "Instalando conversores-ps" -Estado "Configurando variables de entorno PATH..."

$rutasExtra = @(
    "$env:USERPROFILE\.local\bin",
    "$env:LOCALAPPDATA\Pandoc",
    "C:\Program Files\Pandoc",
    "C:\Program Files\Tesseract-OCR",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR",
    "C:\Program Files\wkhtmltopdf\bin",
    "C:\Program Files (x86)\wkhtmltopdf\bin"
)

$pathUsuario = [Environment]::GetEnvironmentVariable("Path", "User")
if (-not $pathUsuario) { $pathUsuario = "" }
$listaPath = @($pathUsuario -split ";" | Where-Object { $_ -ne "" })

foreach ($ruta in $rutasExtra) {
    if ((Test-Path $ruta) -and ($listaPath -notcontains $ruta)) {
        $listaPath += $ruta
    }
}

$nuevoPathUsuario = ($listaPath -join ";")
[Environment]::SetEnvironmentVariable("Path", $nuevoPathUsuario, "User")
$env:Path = "$nuevoPathUsuario;$env:Path"

# 4. Idioma espanol para Tesseract OCR
Mostrar-Progreso -Paso 4 -Actividad "Instalando conversores-ps" -Estado "Comprobando paquete OCR en espanol (spa.traineddata)..."

$posiblesTessdata = @(
    "C:\Program Files\Tesseract-OCR\tessdata",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tessdata"
)

foreach ($dirTess in $posiblesTessdata) {
    if (Test-Path $dirTess) {
        $spaFile = Join-Path $dirTess "spa.traineddata"
        if (-not (Test-Path $spaFile)) {
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri "https://github.com/tesseract-ocr/tessdata_fast/raw/main/spa.traineddata" -OutFile $spaFile -UseBasicParsing
                Write-Host "  [OK] Idioma espanol descargado en $dirTess" -ForegroundColor Green
            } catch {
                Write-Host "  [AVISO] No se pudo descargar spa.traineddata (sin permisos en Program Files o sin red)." -ForegroundColor Yellow
            }
        }
    }
}

# 5. Copia de scripts con codificacion segura (UTF-8 con BOM para PowerShell 5.1)
Mostrar-Progreso -Paso 5 -Actividad "Instalando conversores-ps" -Estado "Copiando conversores.ps1 y pdf2md.py a $env:USERPROFILE..."

$origenPs1 = Join-Path $dirOrigen "conversores.ps1"
$origenPy  = Join-Path $dirOrigen "pdf2md.py"
$destPs1   = Join-Path $env:USERPROFILE "conversores.ps1"
$destPy    = Join-Path $env:USERPROFILE "pdf2md.py"

if (-not (Test-Path $origenPs1) -or -not (Test-Path $origenPy)) {
    throw "No se encontraron 'conversores.ps1' o 'pdf2md.py' junto a install.ps1."
}

$contenidoPs1 = [System.IO.File]::ReadAllText($origenPs1, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($destPs1, $contenidoPs1, $utf8Bom)

$contenidoPy = [System.IO.File]::ReadAllText($origenPy, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($destPy, $contenidoPy, (New-Object System.Text.UTF8Encoding $false))

Unblock-File -Path $destPs1 -ErrorAction SilentlyContinue
Unblock-File -Path $destPy -ErrorAction SilentlyContinue

# 6. Actualizacion selectiva del $PROFILE mediante AST (sin borrar otras funciones del usuario)
Mostrar-Progreso -Paso 6 -Actividad "Instalando conversores-ps" -Estado "Configurando perfil de PowerShell ($PROFILE)..."

$dirPerfil = Split-Path -Parent $PROFILE
if (-not (Test-Path $dirPerfil)) {
    New-Item -ItemType Directory -Path $dirPerfil -Force | Out-Null
}

$funcionesSuite = @(
    "Buscar-ArchivoLocal", "Ejecutar-ConBarraCarga",
    "2md", "convertirmd", "pdf2md",
    "Preparar-PlantillaWord", "Convertir-ConPlantillaWord",
    "2docx", "2pdf", "2code"
)

$contenidoPerfil = ""
if (Test-Path $PROFILE) {
    $contenidoPerfil = [System.IO.File]::ReadAllText($PROFILE, [System.Text.Encoding]::UTF8)
    $tokens = $null
    $erroresAst = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($contenidoPerfil, [ref]$tokens, [ref]$erroresAst)

    $rangosEliminar = @()

    # Eliminar definiciones de funciones previas de la suite
    $nodosFunc = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    foreach ($fn in $nodosFunc) {
        if ($funcionesSuite -contains $fn.Name) {
            $rangosEliminar += [PSCustomObject]@{ Start = $fn.Extent.StartOffset; End = $fn.Extent.EndOffset }
        }
    }

    # Eliminar bloques if completos que carguen conversores.ps1
    $nodosIf = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] }, $true)
    foreach ($ifNode in $nodosIf) {
        if ($ifNode.Extent.Text -match 'conversores\.ps1') {
            $rangosEliminar += [PSCustomObject]@{ Start = $ifNode.Extent.StartOffset; End = $ifNode.Extent.EndOffset }
        }
    }

    $rangosOrdenados = @($rangosEliminar | Sort-Object Start -Descending)
    foreach ($r in $rangosOrdenados) {
        if ($r.Start -ge 0 -and $r.End -le $contenidoPerfil.Length -and $r.End -gt $r.Start) {
            $contenidoPerfil = $contenidoPerfil.Remove($r.Start, ($r.End - $r.Start))
        }
    }

    # Limpiar lineas sueltas antiguas de carga o alias de la suite
    $lineas = @($contenidoPerfil -split "`r?`n" | Where-Object {
        $_ -notmatch 'conversores\.ps1' -and
        $_ -notmatch 'Set-Alias\s+(convertirmd|pdf2md)\s+2md' -and
        $_ -notmatch '# --- Suite conversores-ps ---'
    })
    $contenidoPerfil = ($lineas -join "`r`n").Trim()

    # Reparar cualquier llave huerfana si hubiera quedado de una version anterior
    for ($intento = 0; $intento -lt 10; $intento++) {
        $tCheck = $null
        $eCheck = $null
        [System.Management.Automation.Language.Parser]::ParseInput($contenidoPerfil, [ref]$tCheck, [ref]$eCheck) | Out-Null
        $llaveSobrante = $eCheck | Where-Object { $_.ErrorId -eq "UnexpectedToken" -and $_.Extent.Text -eq "}" } | Select-Object -First 1
        if ($null -ne $llaveSobrante) {
            $contenidoPerfil = $contenidoPerfil.Remove($llaveSobrante.Extent.StartOffset, ($llaveSobrante.Extent.EndOffset - $llaveSobrante.Extent.StartOffset)).Trim()
        } else {
            break
        }
    }
}

$bloqueCarga = @"
# --- Suite conversores-ps ---
if (Test-Path "`$env:USERPROFILE\conversores.ps1") {
    Invoke-Expression (Get-Content "`$env:USERPROFILE\conversores.ps1" -Raw -Encoding UTF8)
}
"@

if ($contenidoPerfil.Length -gt 0) {
    $perfilFinal = $contenidoPerfil + "`r`n`r`n" + $bloqueCarga + "`r`n"
} else {
    $perfilFinal = $bloqueCarga + "`r`n"
}

[System.IO.File]::WriteAllText($PROFILE, $perfilFinal, $utf8Bom)
Write-Progress -Activity "Instalando conversores-ps" -Completed

Write-Host "`n============================================================================" -ForegroundColor Green
Write-Host "  SUITE INSTALADA CORRECTAMENTE: 2md | 2docx | 2pdf | 2code" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Green

if ($bloqueadoPorAdmin) {
    Write-Host "`n[IMPORTANTE - ORDENADOR DE INSTITUTO CON POLITICA RESTRINGIDA]" -ForegroundColor Yellow
    Write-Host "Como el administrador bloquea la carga automatica del perfil, cuando abras PowerShell" -ForegroundColor Yellow
    Write-Host "ejecuta una sola vez este comando para activar los comandos en esa ventana:" -ForegroundColor Yellow
    Write-Host '  iex (gc "$env:USERPROFILE\conversores.ps1" -Raw -Encoding UTF8)' -ForegroundColor White
}