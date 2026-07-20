"""pdf_converter.py — convert DOCX to PDF."""
import os
import subprocess
import shutil


def docx_to_pdf(docx_path, pdf_path):
    """Try multiple conversion methods. Returns (success, message)."""
    # Method 1: docx2pdf (requires MS Word on Windows)
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 100:
            return True, "Microsoft Word"
    except Exception:
        pass

    # Method 2: LibreOffice (works on Linux/Mac)
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo:
        try:
            out_dir = os.path.dirname(pdf_path)
            result = subprocess.run(
                [lo, "--headless", "--convert-to", "pdf",
                 "--outdir", out_dir, docx_path],
                capture_output=True, timeout=60)
            # LibreOffice names the output after the input file
            expected = os.path.join(
                out_dir,
                os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
            if expected != pdf_path and os.path.exists(expected):
                os.rename(expected, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 100:
                return True, "LibreOffice"
        except Exception:
            pass

    return False, "PDF conversion unavailable. Open the DOCX in Word and save as PDF."
