function Buscar-ArchivoLocal {
    param([string]$InputFile)
    $coincidencias = @()

    if (Test-Path $InputFile) {
        $coincidencias += Get-Item (Resolve-Path $InputFile).Path
    } else {
        $misImagenes = [Environment]::GetFolderPath("MyPictures")
        $misDocumentos = [Environment]::GetFolderPath("MyDocuments")
        $miEscritorio = [Environment]::GetFolderPath("Desktop")

        $rutas = @(
            "$env:USERPROFILE\Downloads\Documents",
            "$env:USERPROFILE\Downloads",
            $misDocumentos,
            "$env:USERPROFILE\Documents",
            $miEscritorio,
            "$env:USERPROFILE\Desktop",
            $misImagenes,
            "$misImagenes\Screenshots",
            "$misImagenes\Capturas de pantalla",
            "$env:USERPROFILE\Pictures",
            "$env:USERPROFILE\Pictures\Screenshots",
            "$env:USERPROFILE\Pictures\Capturas de pantalla",
            "$env:USERPROFILE\OneDrive\Imágenes",
            "$env:USERPROFILE\OneDrive\Imágenes\Capturas de pantalla",
            "$env:USERPROFILE\OneDrive\Pictures",
            "$env:USERPROFILE\OneDrive\Pictures\Screenshots"
        ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

        foreach ($carpeta in $rutas) {
            $coincidencias += Get-ChildItem -Path $carpeta -Filter $InputFile -Recurse -File -ErrorAction SilentlyContinue
        }
        $coincidencias = $coincidencias | Sort-Object FullName -Unique
    }

    if ($coincidencias.Count -eq 0) {
        Write-Host "No se encontró el archivo '$InputFile' en tus carpetas habituales." -ForegroundColor Red
        return $null
    }
    if ($coincidencias.Count -eq 1) {
        Write-Host "Archivo encontrado: $($coincidencias[0].FullName)" -ForegroundColor Cyan
        return $coincidencias[0].FullName
    }

    Write-Host "`nSe encontraron $($coincidencias.Count) archivos con ese nombre:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $coincidencias.Count; $i++) {
        $fecha = $coincidencias[$i].LastWriteTime.ToString("dd/MM/yyyy HH:mm")
        Write-Host "  [$($i + 1)] $($coincidencias[$i].FullName) (Modificado: $fecha)"
    }
    do {
        $seleccion = Read-Host "`nElige el número del archivo (1-$($coincidencias.Count))"
        $indice = 0
        $valido = [int]::TryParse($seleccion, [ref]$indice) -and ($indice -ge 1) -and ($indice -le $coincidencias.Count)
    } while (-not $valido)
    return $coincidencias[$indice - 1].FullName
}

function 2md {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    $resolved = Buscar-ArchivoLocal $InputFile
    if (-not $resolved) { return }
    if (-not $OutputFile) {
        $OutputFile = [System.IO.Path]::ChangeExtension($resolved, ".md")
    }
    uv run --with "markitdown,mammoth,python-docx,python-pptx,openpyxl,beautifulsoup4,pymupdf,pytesseract,pillow,extract-msg" "$env:USERPROFILE\pdf2md.py" "$resolved" "$OutputFile"
}

Set-Alias -Name convertirmd -Value 2md
Set-Alias -Name pdf2md -Value 2md

function Reemplazar-TextoWord {
    param($Doc, [string]$Buscar, [string]$Reemplazar)
    if (-not $Buscar -or -not $Reemplazar) { return }
    foreach ($story in $Doc.StoryRanges) {
        $rango = $story
        while ($rango -ne $null) {
            $find = $rango.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Execute($Buscar, $true, $false, $false, $false, $false, $true, 1, $false, $Reemplazar, 2) | Out-Null
            $rango = $rango.NextStoryRange
        }
    }
    foreach ($shape in $Doc.Shapes) {
        if ($shape.TextFrame.HasText) {
            $find = $shape.TextFrame.TextRange.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Execute($Buscar, $true, $false, $false, $false, $false, $true, 1, $false, $Reemplazar, 2) | Out-Null
        }
    }
}

function Convertir-ConPlantillaWord {
    param(
        [string]$MdPath,
        [string]$PlantillaPath,
        [string]$DestinoPath,
        [bool]$ExportarPdf,
        [string]$Titulo,
        [string]$Subtitulo
    )

    $lineas = Get-Content $MdPath -Encoding UTF8
    $primerIndiceNoVacio = -1
    for ($i = 0; $i -lt $lineas.Count; $i++) {
        if ($lineas[$i].Trim() -ne "") {
            $primerIndiceNoVacio = $i
            break
        }
    }

    if ($primerIndiceNoVacio -ge 0 -and $lineas[$primerIndiceNoVacio] -match '^#\s+(.+)$') {
        $cabecera = $Matches[1].Trim()
        if (-not $Titulo) { $lineas[$primerIndiceNoVacio] = "" }

        if (-not $Titulo) {
            if ($cabecera -match '^(.+?)\s*(?:·|-|:)\s*(.+)$') {
                $Titulo = $Matches[1].Trim()
                if (-not $Subtitulo) { $Subtitulo = $Matches[2].Trim() }
            } else {
                $Titulo = $cabecera
            }
        }
    }

    if (-not $Titulo) {
        $Titulo = [System.IO.Path]::GetFileNameWithoutExtension($MdPath) -replace '_', ' '
    }
    if (-not $Subtitulo) {
        $Subtitulo = "Sistemas de Gestión Empresarial"
    }

    $tempMdLimpio = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "md_clean_" + [guid]::NewGuid().ToString() + ".md")
    Set-Content -Path $tempMdLimpio -Value $lineas -Encoding UTF8

    $tempMdDocx = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "pandoc_" + [guid]::NewGuid().ToString() + ".docx")
    pandoc "$tempMdLimpio" -o "$tempMdDocx"
    Remove-Item $tempMdLimpio -Force -ErrorAction SilentlyContinue

    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $tempMdDocx)) {
        Write-Host "Error al convertir el Markdown con Pandoc." -ForegroundColor Red
        return $false
    }

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($PlantillaPath, $false, $true)

        Reemplazar-TextoWord $doc "Subtítulo del documento a varias líneas" $Subtitulo
        Reemplazar-TextoWord $doc "Subtítulo del encabezado" $Subtitulo
        Reemplazar-TextoWord $doc "Título del documento a varias líneas" $Titulo
        Reemplazar-TextoWord $doc "TÍTULO DEL ENCABEZADO" $Titulo.ToUpper()

        $selection = $word.Selection
        $selection.EndKey(6)
        $selection.InsertFile($tempMdDocx)

        if ($ExportarPdf) {
            $doc.SaveAs([ref]$DestinoPath, [ref]17)
        } else {
            $doc.SaveAs([ref]$DestinoPath)
        }
        return $true
    } catch {
        Write-Host "Error en Word: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    } finally {
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
        if (Test-Path $tempMdDocx) { Remove-Item $tempMdDocx -Force }
    }
}

function 2docx {
    param(
        [string]$InputFile,
        [string]$Plantilla,
        [string]$Titulo,
        [string]$Subtitulo,
        [string]$OutputFile
    )
    $resolved = Buscar-ArchivoLocal $InputFile
    if (-not $resolved) { return }
    if ((Get-Item $resolved).Length -eq 0) {
        Write-Host "El archivo '$resolved' está vacío (0 bytes)." -ForegroundColor Red
        return
    }
    if (-not $OutputFile) { $OutputFile = [System.IO.Path]::ChangeExtension($resolved, ".docx") }

    if ($Plantilla) {
        $resolvedPlantilla = Buscar-ArchivoLocal $Plantilla
        if (-not $resolvedPlantilla) { return }
        if (Convertir-ConPlantillaWord -MdPath $resolved -PlantillaPath $resolvedPlantilla -DestinoPath $OutputFile -ExportarPdf $false -Titulo $Titulo -Subtitulo $Subtitulo) {
            Write-Host "Convertido con éxito en: $OutputFile" -ForegroundColor Green
        }
    } else {
        pandoc "$resolved" -o "$OutputFile"
        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputFile)) {
            Write-Host "Convertido con éxito en: $OutputFile" -ForegroundColor Green
        }
    }
}

function 2pdf {
    param(
        [string]$InputFile,
        [string]$Plantilla,
        [string]$Titulo,
        [string]$Subtitulo,
        [string]$OutputFile
    )
    $resolved = Buscar-ArchivoLocal $InputFile
    if (-not $resolved) { return }
    if ((Get-Item $resolved).Length -eq 0) {
        Write-Host "El archivo '$resolved' está vacío (0 bytes)." -ForegroundColor Red
        return
    }
    if (-not $OutputFile) { $OutputFile = [System.IO.Path]::ChangeExtension($resolved, ".pdf") }

    if ($Plantilla) {
        $resolvedPlantilla = Buscar-ArchivoLocal $Plantilla
        if (-not $resolvedPlantilla) { return }
        if (Convertir-ConPlantillaWord -MdPath $resolved -PlantillaPath $resolvedPlantilla -DestinoPath $OutputFile -ExportarPdf $true -Titulo $Titulo -Subtitulo $Subtitulo) {
            Write-Host "Convertido con éxito en: $OutputFile" -ForegroundColor Green
        }
    } else {
        $wkPath = @(
            "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "$env:LOCALAPPDATA\Programs\wkhtmltopdf\bin\wkhtmltopdf.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($wkPath) {
            pandoc "$resolved" --pdf-engine="$wkPath" -o "$OutputFile"
        } else {
            pandoc "$resolved" --pdf-engine="wkhtmltopdf" -o "$OutputFile"
        }
        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputFile)) {
            Write-Host "Convertido con éxito en: $OutputFile" -ForegroundColor Green
        }
    }
}

function 2code {
    param(
        [string]$InputFile,
        [string]$OutputDir
    )
    $resolved = Buscar-ArchivoLocal $InputFile
    if (-not $resolved) { return }

    if (-not $OutputDir) {
        $OutputDir = Split-Path -Parent $resolved
    } elseif (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $texto = Get-Content $resolved -Raw -Encoding UTF8

    $extMap = @{
        "java" = ".java"; "xml" = ".xml"; "python" = ".py"; "py" = ".py";
        "sql" = ".sql"; "json" = ".json"; "yaml" = ".yaml"; "yml" = ".yml";
        "javascript" = ".js"; "js" = ".js"; "typescript" = ".ts"; "ts" = ".ts";
        "html" = ".html"; "css" = ".css"; "csharp" = ".cs"; "cs" = ".cs";
        "cpp" = ".cpp"; "c" = ".c"; "bash" = ".sh"; "sh" = ".sh";
        "powershell" = ".ps1"; "ps1" = ".ps1"; "properties" = ".properties";
        "groovy" = ".gradle"; "gradle" = ".gradle"
    }

    $patron = '(?ms)(?:(?:###?\s*(?:Archivo:\s*)?`?([^\r\n`]+\.[a-zA-Z0-9]+)`?|(?://|#|<!--)\s*([^\r\n]+?\.[a-zA-Z0-9]+)\s*(?:-->)?)\s*\r?\n)?```([a-zA-Z0-9_-]*)\r?\n(.*?)\r?\n```'
    $matches = [regex]::Matches($texto, $patron)

    if ($matches.Count -eq 0) {
        Write-Host "No se encontraron bloques de código en '$resolved'." -ForegroundColor Yellow
        return
    }

    $contador = 1
    foreach ($m in $matches) {
        $nombreExplicito = if ($m.Groups[1].Value) { $m.Groups[1].Value.Trim() } else { $m.Groups[2].Value.Trim() }
        $lang = $m.Groups[3].Value.Trim().ToLower()
        $codigo = $m.Groups[4].Value

        if (-not $nombreExplicito -and $codigo -match '^\s*(?://|#|<!--)\s*([a-zA-Z0-9_.-]+\.[a-zA-Z0-9]+)\s*(?:-->)?\r?\n') {
            $nombreExplicito = $Matches[1].Trim()
            $codigo = $codigo -replace '^\s*(?://|#|<!--)\s*[a-zA-Z0-9_.-]+\.[a-zA-Z0-9]+\s*(?:-->)?\r?\n', ''
        }

        if ($nombreExplicito) {
            $nombreArchivo = [System.IO.Path]::GetFileName($nombreExplicito)
        } else {
            $ext = if ($extMap.ContainsKey($lang)) { $extMap[$lang] } elseif ($lang) { ".$lang" } else { ".txt" }
            if ($lang -eq "java" -and $codigo -match 'public\s+(?:class|interface|enum|record)\s+([A-Za-z0-9_]+)') {
                $nombreArchivo = "$($Matches[1]).java"
            } else {
                $nombreArchivo = "codigo_$contador$ext"
                $contador++
            }
        }

        $rutaSalida = Join-Path $OutputDir $nombreArchivo
        Set-Content -Path $rutaSalida -Value $codigo -Encoding UTF8
        Write-Host "Creado: $rutaSalida" -ForegroundColor Green
    }
}
