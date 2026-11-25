import os
import re
import sys
import zipfile
import tempfile
from pathlib import Path

os.environ["PYTHONWARNINGS"] = "ignore"

import pymupdf  # PyMuPDF
import pytesseract
from PIL import Image
from markitdown import MarkItDown


def configurar_tesseract() -> bool:
    posibles = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for ruta in posibles:
        if ruta and os.path.isfile(ruta):
            pytesseract.pytesseract.tesseract_cmd = ruta
            return True
    return False


def limpiar_ruido_ocr(texto: str) -> str:
    lineas_limpias = []
    for linea in texto.splitlines():
        s = linea.strip()
        if not s:
            lineas_limpias.append("")
            continue
        if len(s) <= 3 and not re.search(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]", s):
            continue
        lineas_limpias.append(linea)
    resultado = "\n".join(lineas_limpias)
    return re.sub(r"\n{3,}", "\n\n", resultado).strip()


def convertir_pdf_a_md(ruta_pdf: Path) -> str:
    tiene_ocr = configurar_tesseract()
    doc = pymupdf.open(str(ruta_pdf))
    partes = [f"# {ruta_pdf.stem}\n"]

    for num_pag, pagina in enumerate(doc, start=1):
        texto = pagina.get_text("text").strip()

        if len(texto) < 30 and len(pagina.get_images()) > 0 and tiene_ocr:
            try:
                pix = pagina.get_pixmap(dpi=250)
                modo = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(modo, [pix.width, pix.height], pix.samples)
                texto_ocr = pytesseract.image_to_string(img, lang="spa+eng")
                texto = limpiar_ruido_ocr(texto_ocr)
            except Exception as e:
                texto = f"> *[Aviso OCR en página {num_pag}: {e}]*"

        if texto:
            partes.append(f"## Página {num_pag}\n\n{texto}")

    doc.close()
    return "\n\n---\n\n".join(partes).strip() + "\n"


def convertir_archivo(ruta: Path, md_engine: MarkItDown) -> str:
    ext = ruta.suffix.lower()

    if ext == ".pdf":
        return convertir_pdf_a_md(ruta)

    if ext == ".zip":
        secciones = [f"# Archivo ZIP: {ruta.name}\n"]
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(ruta, "r") as zf:
                zf.extractall(tmpdir)
            for sub in sorted(Path(tmpdir).rglob("*")):
                if sub.is_file() and not sub.name.startswith("."):
                    rel = sub.relative_to(tmpdir)
                    try:
                        contenido_sub = convertir_archivo(sub, md_engine)
                        secciones.append(f"## {rel}\n\n{contenido_sub}")
                    except Exception as e:
                        secciones.append(f"## {rel}\n\n> *[No se pudo convertir: {e}]*")
        return "\n\n".join(secciones)

    extensiones_codigo = {
        ".java": "java", ".py": "python", ".ps1": "powershell", ".js": "javascript",
        ".ts": "typescript", ".html": "html", ".css": "css", ".xml": "xml",
        ".json": "json", ".sql": "sql", ".c": "c", ".cpp": "cpp", ".cs": "csharp",
        ".sh": "bash", ".bat": "bat", ".yml": "yaml", ".yaml": "yaml", ".md": "markdown",
        ".txt": ""
    }
    if ext in extensiones_codigo:
        raw = ruta.read_text(encoding="utf-8", errors="replace")
        lang = extensiones_codigo[ext]
        if ext in {".md", ".txt"}:
            return raw
        return f"### `{ruta.name}`\n\n```{lang}\n{raw}\n```\n"

    res = md_engine.convert(str(ruta))
    return (res.text_content or "").strip() + "\n"


def main():
    if len(sys.argv) < 3:
        print("Uso: python pdf2md.py <archivo_entrada> <archivo_salida.md>", file=sys.stderr)
        sys.exit(1)

    ruta_entrada = Path(sys.argv[1]).resolve()
    ruta_salida = Path(sys.argv[2]).resolve()

    if not ruta_entrada.is_file():
        print(f"Error: No existe '{ruta_entrada}'", file=sys.stderr)
        sys.exit(1)

    md_engine = MarkItDown(enable_plugins=False)
    contenido_md = convertir_archivo(ruta_entrada, md_engine)
    ruta_salida.write_text(contenido_md, encoding="utf-8")


if __name__ == "__main__":
    main()