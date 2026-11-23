# Suite de Conversores Documentales y de Código para PowerShell (`conversores-ps`)

Conjunto de herramientas portables para Windows 11 que añade cuatro comandos rápidos a tu terminal PowerShell (`2md`, `2docx`, `2pdf` y `2code`) con búsqueda automática de archivos en tus carpetas locales y soporte OCR en español.

---

## Instalación rápida en cualquier ordenador con Windows

### Método 1: Descargando el ZIP (sin usar Git ni iniciar sesión)

1. Descarga el repositorio en ZIP desde GitHub (**Code → Download ZIP**) y descomprímelo.
2. Haz **doble clic en `Instalar.bat`** (o abre PowerShell en esa carpeta y ejecuta `powershell -ExecutionPolicy Bypass -File .\install.ps1`).
3. Espera a que termine la instalación automática y abre una ventana nueva de PowerShell.

### Método 2: En un solo comando desde PowerShell (si el repositorio es público)

```powershell
$zip = "$env:TEMP\conv.zip"
Invoke-WebRequest -Uri "https://github.com/sam23-23/conversores-ps/archive/refs/heads/main.zip" -OutFile $zip
Expand-Archive -Path $zip -DestinationPath "$env:TEMP\conv" -Force
$inst = Get-ChildItem "$env:TEMP\conv" -Filter "install.ps1" -Recurse | Select-Object -First 1
powershell -ExecutionPolicy Bypass -File $inst.FullName
Remove-Item $zip, "$env:TEMP\conv" -Recurse -Force
```

> **Comportamiento inteligente sobre `$PROFILE`:** El instalador analiza el árbol sintáctico (AST) de tu archivo `$PROFILE` antes de escribir. Si ya existían definiciones previas de `2md`, `2docx`, `2pdf`, `2code` o sus funciones auxiliares/alias, **elimina únicamente esos bloques antiguos** y enlaza la nueva versión sin tocar ninguna otra función que tuvieras configurada (como scripts de IA, `apagar`, `update-app`, etc.).

---

## Guía de uso de los comandos

Todos los comandos buscan el archivo automáticamente en la carpeta actual, `Descargas`, `Documentos`, `Escritorio`, `Capturas de pantalla` y `OneDrive` (no hace falta escribir la ruta completa).

### 1. `2md` — Convierte cualquier archivo a Markdown limpio
Alias disponibles: `convertirmd`, `pdf2md`.

```powershell
2md "Tema 9 Práctica.pdf"
2md "apuntes.docx"
2md "captura_pantalla.png"
2md "ejercicios_clase.zip"
```

* **Soporta:** PDF (nativo y escaneado mediante OCR en español), Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`), HTML, correos `.msg`, imágenes (`.png`, `.jpg`, etc.), código fuente y archivos `.zip` con múltiples ficheros dentro.

### 2. `2docx` — Convierte Markdown a Word (`.docx`)

```powershell
# Conversión estándar
2docx "resumen.md"

# Rellenando automáticamente portada y encabezados de una plantilla Word
2docx "memoria.md" -Plantilla "Plantilla_Clase.docx" -Titulo "Práctica 6" -Subtitulo "Sistemas de Gestión"
```

### 3. `2pdf` — Convierte Markdown a PDF

```powershell
# Conversión rápida con wkhtmltopdf
2pdf "resumen.md"

# Exportación profesional a través de una plantilla Word
2pdf "memoria.md" -Plantilla "Plantilla_Clase.docx"
```

### 4. `2code` — Extrae archivos de código desde un Markdown

```powershell
2code "respuesta_ia.md"
2code "respuesta_ia.md" -OutputDir ".\src"
```

* Detecta todos los bloques de código (` ```java `, ` ```xml `, ` ```py `, etc.) dentro del Markdown, deduce el nombre de archivo a partir del encabezado, comentario o declaración `public class ...` y genera cada fichero físico en disco.

---

## Arquitectura del repositorio (Qué hace cada archivo)

| Archivo | Función técnica |
| :--- | :--- |
| **`Instalar.bat`** | Lanzador de doble clic que ejecuta `install.ps1` con `-ExecutionPolicy Bypass` para evitar bloqueos de seguridad de Windows al descargar el ZIP. |
| **`install.ps1`** | Instala las 4 dependencias mediante `winget` (`uv`, `Pandoc`, `Tesseract OCR` y `wkhtmltopdf`), añade sus rutas al `PATH` del usuario, descarga el modelo OCR en español (`spa.traineddata`), desbloquea los scripts con `Unblock-File` y utiliza el parser AST nativo de PowerShell para sustituir selectivamente solo las funciones de esta suite en `$PROFILE` sin alterar el resto del archivo. |
| **`conversores.ps1`** | Contiene las funciones de PowerShell: `Buscar-ArchivoLocal` (búsqueda recursiva multi-carpeta con selector interactivo), `2md` (invoca `uv run` con entorno efímero), `2docx` / `2pdf` (con automatización COM de Microsoft Word para rellenar portadas/encabezados de plantillas) y `2code` (parser regex para extraer ficheros de código). |
| **`pdf2md.py`** | Motor en Python ejecutado por `uv`. Combina `MarkItDown` para documentos Office con `PyMuPDF` (`pymupdf`) + `Tesseract OCR` (`spa+eng` a 250 DPI con filtro de ruido de iconos) para PDFs escaneados o basados en imágenes, además de empaquetar código fuente y descomprimir `.zip` recursivamente. |
