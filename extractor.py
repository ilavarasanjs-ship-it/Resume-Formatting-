"""extractor.py — read raw text from PDF or DOCX."""
import io


def extract_text(file_bytes, filename):
    fn = filename.lower()
    if fn.endswith('.pdf'):
        return _from_pdf(file_bytes)
    elif fn.endswith('.docx'):
        return _from_docx(file_bytes)
    elif fn.endswith('.doc'):
        return _from_docx(file_bytes)  # best effort
    return file_bytes.decode('utf-8', errors='ignore')


def _from_pdf(file_bytes):
    text_parts = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_parts.append(t)
        if text_parts:
            return "\n".join(text_parts)
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
    except Exception:
        pass
    return "\n".join(text_parts)


def _from_docx(file_bytes):
    try:
        import docx
        doc = docx.Document(io.BytesIO(file_bytes))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(parts)
    except Exception as e:
        return f"[Could not read file: {e}]"
