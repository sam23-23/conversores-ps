# ============================================================================
# SUITE PORTABLE DE CONVERSIONES PARA POWERSHELL
# Comandos: 2md | 2docx | 2pdf | 2code
# ============================================================================

function Buscar-ArchivoLocal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$NombreOPath
    )

    if (Test-Path -LiteralPath $NombreOPath) {
        return (Resolve-Path -LiteralPath $NombreOPath).Path
    }

    $rutasBusqueda = @(
        (Get-Location).Path,
        "$env:USERPROFILE\Downloads",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Pictures",
        "$env:USERPROFILE\OneDrive",
        "$env:USERPROFILE\OneDrive\Documentos",
        "$env:USERPROFILE\OneDrive\Escritorio"
    ) | Where-Object { Test-Path $_ } | Select-Object -Unique

    $encontradosTotal = @()
    foreach ($carpeta in $rutasBusqueda) {
        $encontrados = @(Get-ChildItem -Path $carpeta -Filter $NombreOPath -File -Recurse -ErrorAction SilentlyContinue)
        if ($encontrados.Count -gt 0) {
            $encontradosTotal += $encontrados
        }
    }

    $coincidencias = @($encontradosTotal | ForEach-Object { $_.FullName } | Select-Object -Unique)

    if ($coincidencias.Count -eq 0) {
        throw "No se encontro el archivo '$NombreOPath' en la carpeta actual ni en Downloads/Documents/Desktop/OneDrive."
    }

    if ($coincidencias.Count -eq 1) {
        Write-Host "Archivo encontrado: $($coincidencias[0])" -ForegroundColor Cyan
        return $coincidencias[0]
    }

    Write-Host "Se encontraron varias coincidencias para '$NombreOPath':" -ForegroundColor Yellow
    for ($i = 0; $i -lt $coincidencias.Count; $i++) {
        Write-Host "  [$($i + 1)] $($coincidencias[$i])" -ForegroundColor White
    }
    $sel = Read-Host "Elige el numero del archivo (1-$($coincidencias.Count))"
    $idx = 0
    if ([int]::TryParse($sel, [ref]$idx) -and $idx -ge 1 -and $idx -le $coincidencias.Count) {
        return $coincidencias[$idx - 1]
    }
    throw "Seleccion no valida."
}

function Ejecutar-ConBarraCarga {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mensaje,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Accion,
        [object[]]$Argumentos = @()
    )

    $pathActual = $env:Path
    $job = Start-Job -ScriptBlock {
        param($sbText, $argsList, $pathEnv)
        $env:Path = $pathEnv
        $env:PYTHONWARNINGS = "ignore"
        $sb = [scriptblock]::Create($sbText)
        $salida = & $sb @argsList 2>&1
        $codigo = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 0 }
        return [PSCustomObject]@{
            Output   = $salida
            ExitCode = $codigo
        }
    } -ArgumentList $Accion.ToString(), $Argumentos, $pathActual

    $pct = 10
    while ($job.State -eq "Running") {
        Write-Progress -Activity "Procesando conversion" -Status $Mensaje -PercentComplete $pct
        Start-Sleep -Milliseconds 250
        $pct = if ($pct -ge 90) { 15 } else { $pct + 5 }
    }

    Write-Progress -Activity "Procesando conversion" -Completed
    $resultado = Receive-Job -Job $job -ErrorAction SilentlyContinue
    $erroresJob = $job.ChildJobs[0].Error
    Remove-Job -Job $job -Force

    if ($erroresJob -and $erroresJob.Count -gt 0) {
        throw ($erroresJob[0].Exception.Message)
    }

    if ($resultado -and $resultado.ExitCode -ne 0) {
        $detalle = ($resultado.Output | Out-String).Trim()
        throw "Error en el proceso (codigo $($resultado.ExitCode)):`n$detalle"
    }

    if ($resultado -and $resultado.Output) {
        $textoSalida = ($resultado.Output | Out-String).Trim()
        if ($textoSalida) { Write-Host $textoSalida -ForegroundColor Gray }
    }
}

# ----------------------------------------------------------------------------
# 1. 2md: Convierte PDF (digital o escaneado con OCR), DOCX, PPTX, XLSX, HTML,
#         MSG, ZIP o codigo fuente a Markdown (.md)
# ----------------------------------------------------------------------------
function 2md {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Archivo,

        [Parameter(Position = 1)]
        [string]$Salida
    )

    try {
        $rutaEntrada = Buscar-ArchivoLocal -NombreOPath $Archivo
        if (-not $Salida) {
            $Salida = [System.IO.Path]::ChangeExtension($rutaEntrada, ".md")
        } else {
            $Salida = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Salida)
        }

        $scriptPy = Join-Path $env:USERPROFILE "pdf2md.py"
        if (-not (Test-Path $scriptPy)) {
            throw "No se encuentra '$scriptPy'. Ejecuta Instalar.bat nuevamente."
        }

        Ejecutar-ConBarraCarga -Mensaje "Convirtiendo '$(Split-Path -Leaf $rutaEntrada)' a Markdown..." -Argumentos @($scriptPy, $rutaEntrada, $Salida) -Accion {
            param($py, $inFile, $outFile)
            uv run --quiet --with "markitdown[all],pymupdf,pytesseract,pillow,python-docx,python-pptx,openpyxl,beautifulsoup4,mammoth,extract-msg" $py $inFile $outFile
        }

        if (Test-Path $Salida) {
            Write-Host "[OK] Markdown generado: $Salida" -ForegroundColor Green
        } else {
            throw "No se genero el archivo de salida '$Salida'."
        }
    } catch {
        Write-Host "[ERROR 2md] $($_.Exception.Message)" -ForegroundColor Red
    }
}

Set-Alias -Name convertirmd -Value 2md -Force
Set-Alias -Name pdf2md -Value 2md -Force

# ----------------------------------------------------------------------------
# Auxiliar: Rellenar o limpiar marcadores de Titulo/Subtitulo en plantilla Word
# ----------------------------------------------------------------------------
function Preparar-PlantillaWord {
    param(
        [Parameter(Mandatory = $true)]$DocWord,
        [string]$Titulo,
        [string]$Subtitulo,
        [switch]$SinTitulo
    )

    $usarTitulo = (-not $SinTitulo) -and (-not [string]::IsNullOrWhiteSpace($Titulo))
    $usarSubtitulo = (-not $SinTitulo) -and (-not [string]::IsNullOrWhiteSpace($Subtitulo))

    try {
        if ($usarTitulo) {
            $DocWord.BuiltInDocumentProperties("Title").Value = $Titulo
        } else {
            $DocWord.BuiltInDocumentProperties("Title").Value = ""
        }
    } catch {}

    $reemplazos = @(
        @{ Tags = @("{{TITULO}}", "[TITULO]", "<TITULO>", "{{TITLE}}", "[TITLE]", "<TITLE>"); Valor = $(if ($usarTitulo) { $Titulo } else { "" }) },
        @{ Tags = @("{{SUBTITULO}}", "[SUBTITULO]", "<SUBTITULO>", "{{SUBTITLE}}", "[SUBTITLE]", "<SUBTITLE>"); Valor = $(if ($usarSubtitulo) { $Subtitulo } else { "" }) }
    )

    foreach ($item in $reemplazos) {
        foreach ($tag in $item.Tags) {
            $find = $DocWord.Content.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Execute($tag, $false, $false, $false, $false, $false, $true, 1, $false, $item.Valor, 2) | Out-Null

            foreach ($sec in $DocWord.Sections) {
                foreach ($hdr in $sec.Headers) {
                    $fH = $hdr.Range.Find
                    $fH.Execute($tag, $false, $false, $false, $false, $false, $true, 1, $false, $item.Valor, 2) | Out-Null
                }
                foreach ($ftr in $sec.Footers) {
                    $fF = $ftr.Range.Find
                    $fF.Execute($tag, $false, $false, $false, $false, $false, $true, 1, $false, $item.Valor, 2) | Out-Null
                }
            }
        }
    }

    try {
        foreach ($shape in $DocWord.Shapes) {
            if ($shape.TextFrame.HasText) {
                foreach ($item in $reemplazos) {
                    foreach ($tag in $item.Tags) {
                        $fS = $shape.TextFrame.TextRange.Find
                        $fS.Execute($tag, $false, $false, $false, $false, $false, $true, 1, $false, $item.Valor, 2) | Out-Null
                    }
                }
            }
        }
    } catch {}

    if ($SinTitulo) {
        for ($i = $DocWord.Paragraphs.Count; $i -ge 1; $i--) {
            $p = $DocWord.Paragraphs.Item($i)
            $nombreEstilo = ""
            try { $nombreEstilo = $p.Style.NameLocal } catch {}
            if ($nombreEstilo -match "^(Title|Subtitle|T[ií]tulo|Subt[ií]tulo)$") {
                $p.Range.Delete()
            }
        }
    }
}

function Convertir-ConPlantillaWord {
    param(
        [Parameter(Mandatory = $true)][string]$RutaMd,
        [Parameter(Mandatory = $true)][string]$RutaPlantilla,
        [Parameter(Mandatory = $true)][string]$RutaSalida,
        [Parameter(Mandatory = $true)][ValidateSet("docx", "pdf")][string]$Formato,
        [string]$Titulo,
        [string]$Subtitulo,
        [switch]$SinTitulo
    )

    $tempBodyDocx = Join-Path $env:TEMP ("body_" + [guid]::NewGuid().ToString("N") + ".docx")
    $tempWorkDocx = Join-Path $env:TEMP ("work_" + [guid]::NewGuid().ToString("N") + ".docx")

    try {
        Ejecutar-ConBarraCarga -Mensaje "Generando estructura documental con Pandoc..." -Argumentos @($RutaMd, $RutaPlantilla, $tempBodyDocx) -Accion {
            param($md, $tpl, $outDoc)
            pandoc $md -f markdown -t docx --reference-doc=$tpl -o $outDoc
        }

        Copy-Item -LiteralPath $RutaPlantilla -Destination $tempWorkDocx -Force

        Write-Progress -Activity "Aplicando plantilla de Word" -Status "Combinando estilos, portada y contenido..." -PercentComplete 65

        $word = $null
        $docPlantilla = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0

            $docPlantilla = $word.Documents.Open($tempWorkDocx)

            Preparar-PlantillaWord -DocWord $docPlantilla -Titulo $Titulo -Subtitulo $Subtitulo -SinTitulo:$SinTitulo

            $marcadorContenido = $null
            foreach ($tagC in @("{{CONTENIDO}}", "[CONTENIDO]", "{{CONTENT}}", "[CONTENT]")) {
                $rng = $docPlantilla.Content
                $fC = $rng.Find
                if ($fC.Execute($tagC)) {
                    $marcadorContenido = $rng
                    break
                }
            }

            if ($null -ne $marcadorContenido) {
                $marcadorContenido.Text = ""
                $marcadorContenido.InsertFile($tempBodyDocx)
            } else {
                $fin = $docPlantilla.Content
                $fin.Collapse(0)
                if ($docPlantilla.Content.Text.Trim().Length -gt 0) {
                    $fin.InsertBreak(7)
                }
                $fin.InsertFile($tempBodyDocx)
            }

            if ($Formato -eq "docx") {
                $docPlantilla.SaveAs([ref]$RutaSalida, [ref]16)
            } else {
                $docPlantilla.ExportAsFixedFormat($RutaSalida, 17)
            }
        } finally {
            if ($null -ne $docPlantilla) {
                try { $docPlantilla.Close($false) } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($docPlantilla) | Out-Null
            }
            if ($null -ne $word) {
                try { $word.Quit() } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Write-Progress -Activity "Aplicando plantilla de Word" -Completed
        }
    } finally {
        Remove-Item -LiteralPath $tempBodyDocx -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempWorkDocx -Force -ErrorAction SilentlyContinue
    }
}

# ----------------------------------------------------------------------------
# 2. 2docx: Convierte Markdown (.md) a Word (.docx) con o sin plantilla/titulo
# ----------------------------------------------------------------------------
function 2docx {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Archivo,

        [Parameter(Position = 1)]
        [string]$Salida,

        [string]$Plantilla,
        [string]$Titulo,
        [string]$Subtitulo,
        [switch]$SinTitulo
    )

    try {
        $rutaEntrada = Buscar-ArchivoLocal -NombreOPath $Archivo
        if (-not $Salida) {
            $Salida = [System.IO.Path]::ChangeExtension($rutaEntrada, ".docx")
        } else {
            $Salida = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Salida)
        }

        if ($Plantilla) {
            $rutaPlantilla = Buscar-ArchivoLocal -NombreOPath $Plantilla
            Convertir-ConPlantillaWord -RutaMd $rutaEntrada -RutaPlantilla $rutaPlantilla -RutaSalida $Salida -Formato "docx" `
                -Titulo $Titulo -Subtitulo $Subtitulo -SinTitulo:$SinTitulo
        } else {
            $argsPandoc = @($rutaEntrada, "-f", "markdown", "-t", "docx", "-o", $Salida)
            if (-not $SinTitulo) {
                if (-not [string]::IsNullOrWhiteSpace($Titulo))    { $argsPandoc += @("-M", "title=$Titulo") }
                if (-not [string]::IsNullOrWhiteSpace($Subtitulo)) { $argsPandoc += @("-M", "subtitle=$Subtitulo") }
            }

            Ejecutar-ConBarraCarga -Mensaje "Convirtiendo '$(Split-Path -Leaf $rutaEntrada)' a DOCX..." -Argumentos @(,$argsPandoc) -Accion {
                param($pArgs)
                pandoc @pArgs
            }
        }

        if (Test-Path $Salida) {
            Write-Host "[OK] DOCX generado: $Salida" -ForegroundColor Green
        } else {
            throw "No se genero el archivo '$Salida'."
        }
    } catch {
        Write-Host "[ERROR 2docx] $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ----------------------------------------------------------------------------
# 3. 2pdf: Convierte Markdown (.md) a PDF (.pdf) con o sin plantilla/titulo
# ----------------------------------------------------------------------------
function 2pdf {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Archivo,

        [Parameter(Position = 1)]
        [string]$Salida,

        [string]$Plantilla,
        [string]$Titulo,
        [string]$Subtitulo,
        [switch]$SinTitulo
    )

    try {
        $rutaEntrada = Buscar-ArchivoLocal -NombreOPath $Archivo
        if (-not $Salida) {
            $Salida = [System.IO.Path]::ChangeExtension($rutaEntrada, ".pdf")
        } else {
            $Salida = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Salida)
        }

        if ($Plantilla) {
            $rutaPlantilla = Buscar-ArchivoLocal -NombreOPath $Plantilla
            Convertir-ConPlantillaWord -RutaMd $rutaEntrada -RutaPlantilla $rutaPlantilla -RutaSalida $Salida -Formato "pdf" `
                -Titulo $Titulo -Subtitulo $Subtitulo -SinTitulo:$SinTitulo
        } else {
            $tempHtml = Join-Path $env:TEMP ("doc_" + [guid]::NewGuid().ToString("N") + ".html")
            $tempCss  = Join-Path $env:TEMP ("style_" + [guid]::NewGuid().ToString("N") + ".css")

            try {
                $css = @"
body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 11pt; line-height: 1.55; color: #1f2937; margin: 22mm 20mm; }
h1.title { font-size: 22pt; color: #111827; border-bottom: 2px solid #2563eb; padding-bottom: 6px; margin-bottom: 4px; }
p.subtitle { font-size: 13pt; color: #4b5563; margin-top: 0; margin-bottom: 20px; }
h1, h2, h3 { color: #1e3a8a; margin-top: 1.3em; }
pre, code { font-family: Consolas, 'Courier New', monospace; background: #f3f4f6; border-radius: 4px; }
pre { padding: 10px 12px; border: 1px solid #e5e7eb; white-space: pre-wrap; word-wrap: break-word; }
code { padding: 2px 4px; }
table { border-collapse: collapse; width: 100%; margin: 14px 0; }
th, td { border: 1px solid #d1d5db; padding: 7px 10px; text-align: left; }
th { background: #eff6ff; font-weight: 600; }
blockquote { border-left: 4px solid #3b82f6; margin: 12px 0; padding: 4px 14px; color: #374151; background: #f9fafb; }
"@
                [System.IO.File]::WriteAllText($tempCss, $css, (New-Object System.Text.UTF8Encoding $false))

                $argsPandoc = @($rutaEntrada, "-f", "markdown", "-t", "html5", "--standalone", "--css=$tempCss", "-o", $tempHtml)
                if (-not $SinTitulo) {
                    if (-not [string]::IsNullOrWhiteSpace($Titulo))    { $argsPandoc += @("-M", "title=$Titulo") }
                    if (-not [string]::IsNullOrWhiteSpace($Subtitulo)) { $argsPandoc += @("-M", "subtitle=$Subtitulo") }
                }

                Ejecutar-ConBarraCarga -Mensaje "Renderizando PDF desde '$(Split-Path -Leaf $rutaEntrada)'..." -Argumentos @(,$argsPandoc, $tempHtml, $Salida) -Accion {
                    param($pArgs, $htmlFile, $pdfOut)
                    pandoc @pArgs
                    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
                    wkhtmltopdf --quiet --enable-local-file-access --encoding utf-8 $htmlFile $pdfOut
                }
            } finally {
                Remove-Item -LiteralPath $tempHtml -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $tempCss -Force -ErrorAction SilentlyContinue
            }
        }

        if (Test-Path $Salida) {
            Write-Host "[OK] PDF generado: $Salida" -ForegroundColor Green
        } else {
            throw "No se genero el archivo '$Salida'."
        }
    } catch {
        Write-Host "[ERROR 2pdf] $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ----------------------------------------------------------------------------
# 4. 2code: Extrae bloques de codigo de un .md y crea los ficheros reales
# ----------------------------------------------------------------------------
function 2code {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Archivo,

        [Parameter(Position = 1)]
        [string]$CarpetaDestino
    )

    try {
        $rutaEntrada = Buscar-ArchivoLocal -NombreOPath $Archivo
        if (-not $CarpetaDestino) {
            $CarpetaDestino = (Get-Location).Path
        } else {
            $CarpetaDestino = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CarpetaDestino)
            if (-not (Test-Path $CarpetaDestino)) {
                New-Item -ItemType Directory -Path $CarpetaDestino -Force | Out-Null
            }
        }

        $contenido = [System.IO.File]::ReadAllText($rutaEntrada, [System.Text.Encoding]::UTF8)

        $extMap = @{
            "java" = ".java"; "python" = ".py"; "py" = ".py"; "powershell" = ".ps1"; "ps1" = ".ps1";
            "javascript" = ".js"; "js" = ".js"; "typescript" = ".ts"; "ts" = ".ts";
            "html" = ".html"; "css" = ".css"; "xml" = ".xml"; "json" = ".json";
            "sql" = ".sql"; "c" = ".c"; "cpp" = ".cpp"; "csharp" = ".cs"; "cs" = ".cs";
            "bash" = ".sh"; "sh" = ".sh"; "bat" = ".bat"; "cmd" = ".bat"; "yaml" = ".yml"; "yml" = ".yml"
        }

        $patron = '(?ms)(?:^#{1,6}\s+`?([^\r\n`]+\.[a-zA-Z0-9]{1,8})`?\s*\r?\n+|^\*\*`?([^\r\n`]+\.[a-zA-Z0-9]{1,8})`?\*\*\s*\r?\n+)?```([a-zA-Z0-9_-]*)\r?\n(.*?)\r?\n```'
        $coincidenciasRegex = @([regex]::Matches($contenido, $patron))

        if ($coincidenciasRegex.Count -eq 0) {
            Write-Host "[AVISO] No se encontraron bloques de codigo ```...``` en '$rutaEntrada'." -ForegroundColor Yellow
            return
        }

        $utf8SinBom = New-Object System.Text.UTF8Encoding $false
        $contador = 0

        foreach ($m in $coincidenciasRegex) {
            $contador++
            $nombreHeader = if ($m.Groups[1].Success) { $m.Groups[1].Value.Trim() } elseif ($m.Groups[2].Success) { $m.Groups[2].Value.Trim() } else { "" }
            $lang = $m.Groups[3].Value.Trim().ToLower()
            $codigo = $m.Groups[4].Value

            $lineas = @($codigo -split "`r?`n")
            $nombrePrimeraLinea = ""
            if ($lineas.Count -gt 0 -and $lineas[0] -match '^\s*(?://|#|/\*|<!--)\s*([a-zA-Z0-9_./\\-]+\.[a-zA-Z0-9]{1,8})\s*(?:\*/|-->)?\s*$') {
                $nombrePrimeraLinea = $Matches[1].Trim()
                $codigo = ($lineas | Select-Object -Skip 1) -join "`r`n"
            }

            $nombreRelativo = if ($nombrePrimeraLinea) {
                $nombrePrimeraLinea
            } elseif ($nombreHeader) {
                $nombreHeader
            } else {
                $ext = if ($extMap.ContainsKey($lang)) { $extMap[$lang] } else { ".txt" }
                "bloque_${contador}${ext}"
            }

            $nombreRelativo = $nombreRelativo -replace '^[/\\]+', ''
            $rutaFinal = Join-Path $CarpetaDestino $nombreRelativo
            $dirPadre = Split-Path -Parent $rutaFinal
            if ($dirPadre -and -not (Test-Path $dirPadre)) {
                New-Item -ItemType Directory -Path $dirPadre -Force | Out-Null
            }

            [System.IO.File]::WriteAllText($rutaFinal, $codigo, $utf8SinBom)
            Write-Host "  [+] Creado: $rutaFinal" -ForegroundColor Green
        }

        Write-Host "[OK] Se extrajeron $contador archivo(s) de codigo." -ForegroundColor Cyan
    } catch {
        Write-Host "[ERROR 2code] $($_.Exception.Message)" -ForegroundColor Red
    }
}