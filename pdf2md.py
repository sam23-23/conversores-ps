import sys
import os
import zipfile
from pathlib import Path
from PIL import Image
import pytesseract
import fitz  # PyMuPDF
from markitdown import MarkItDown

tesseract_default = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(tesseract_default):
    pytesseract.pytesseract.tesseract_cmd = tesseract_default

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"}

CODE_LANGS = {
    ".java": "java", ".py": "python", ".sql": "sql", ".xml": "xml",
    ".json": "json", ".yaml": "yaml", ".yml": "yaml", ".js": "javascript",
    ".ts": "typescript", ".html": "html", ".css": "css", ".cs": "csharp",
    ".cpp": "cpp", ".c": "c", ".h": "c", ".sh": "bash", ".ps1": "powershell",
    ".bat": "bat", ".properties": "properties", ".ini": "ini", ".gradle": "groovy"
}

def ocr_image(img: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(img, lang="spa+eng").strip()
    except pytesseract.TesseractError:
        return pytesseract.image_to_string(img).strip()

def ocr_pdf(pdf_path: Path) -> str:
    doc = fitz.open(str(pdf_path))
    paginas = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=200)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        texto_pag = ocr_image(img)
        if texto_pag:
            paginas.append(f"## Página {i}\n\n{texto_pag}")
    return "\n\n---\n\n".join(paginas)

def convert_single_file(file_path: Path, md_converter: MarkItDown) -> str:
    ext = file_path.suffix.lower()

    # 1. Archivos de código fuente o configuración
    if ext in CODE_LANGS:
        contenido = file_path.read_text(encoding="utf-8", errors="replace").strip()
        lang = CODE_LANGS[ext]
        return f"### Archivo: `{file_path.name}`\n\n```{lang}\n{contenido}\n```"

    # 2. Imágenes sueltas
    if ext in IMAGE_EXTS:
        img = Image.open(file_path)
        return ocr_image(img)

    # 3. PDFs (con respaldo automático de OCR si viene como imagen)
    if ext == ".pdf":
        result = md_converter.convert(str(file_path))
        texto = (result.text_content or "").strip()
        if len(texto) < 50:  # PDF escaneado o basado en imágenes
            texto = ocr_pdf(file_path)
        return texto

    # 4. Resto de formatos (Office, CSV, EPUB, MSG, etc.)
    result = md_converter.convert(str(file_path))
    return (result.text_content or "").strip()

def main():
    if len(sys.argv) < 3:
        print("Uso: pdf2md.py <entrada> <salida>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    md_converter = MarkItDown()

    try:
        # Soporte para archivos .zip con múltiples ejercicios/clases dentro
        if input_path.suffix.lower() == ".zip":
            import tempfile
            partes = []
            with tempfile.TemporaryDirectory() as tmpdir:
                with zipfile.ZipFile(input_path, 'r') as zf:
                    zf.extractall(tmpdir)
                for f in sorted(Path(tmpdir).rglob("*")):
                    if f.is_file() and not f.name.startswith("."):
                        try:
                            res = convert_single_file(f, md_converter)
                            if res:
                                partes.append(res)
                        except Exception:
                            pass
            final_text = "\n\n---\n\n".join(partes)
        else:
            final_text = convert_single_file(input_path, md_converter)

        if not final_text.strip():
            print(f"Aviso: No se pudo extraer texto de '{input_path.name}'.")
        else:
            output_path.write_text(final_text.strip() + "\n", encoding="utf-8")
            print(f"Convertido con éxito en: {output_path}")

    except Exception as e:
        print(f"Error al procesar '{input_path.name}': {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
