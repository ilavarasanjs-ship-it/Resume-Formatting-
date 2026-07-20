"""
resume_builder.py
Builds a formatted DOCX from structured candidate data, exactly matching
the client template format observed across all 3 sample pairs.

Format rules (hardcoded from template analysis):
- Name: 16pt Bold Center
- Additional fields (location, commute, availability): 12pt Bold Center
- Section headings (Summary, Education and Certifications, Skills, Experience): Bold Left
- Summary text: Justified
- Education institution: Plain Left
- Education degree / certifications: Bold Bullet (List Paragraph)
- Skills: Plain Bullet (List Paragraph)
- Company line: Bold Left, right-aligned tab at 6.5" for dates
- Job title: Bold Italic Left
- Bullet points: Plain Bullet (List Paragraph)
- Empty spacer paragraphs between sections
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy
import io
import os


def _make_tab_stop_element(position_twips, alignment="right"):
    """Create a w:tab XML element for use in w:tabs."""
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), alignment)
    tab.set(qn('w:pos'), str(position_twips))
    return tab


def _add_paragraph(doc, text="", bold=False, italic=False,
                   font_size_pt=None, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                   style_name="Normal"):
    """Add a paragraph with consistent formatting."""
    p = doc.add_paragraph(style=style_name)
    p.alignment = alignment

    if not text:
        return p

    run = p.add_run(text)
    run.bold  = bold
    run.italic = italic
    if font_size_pt:
        run.font.size = Pt(font_size_pt)
    return p


def _add_company_line(doc, company, dates):
    """
    Add company name (left) with dates right-aligned via tab stop at 6.5".
    Bold throughout. Tab stop mirrors the template XML exactly.
    """
    p   = doc.add_paragraph(style="Normal")
    pPr = p._p.get_or_add_pPr()

    # Add tab stop: right-aligned at 9360 twips (6.5")
    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), '9360')
    tabs.append(tab)
    pPr.append(tabs)

    # Also set bold in paragraph-level rPr
    p_rPr = OxmlElement('w:rPr')
    b_el  = OxmlElement('w:b')
    p_rPr.append(b_el)
    pPr.append(p_rPr)

    def make_bold_run(text_val):
        r   = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        b   = OxmlElement('w:b')
        rPr.append(b)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = text_val
        if text_val.startswith(' ') or text_val.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        return r

    def make_tab_run():
        r   = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        b   = OxmlElement('w:b')
        rPr.append(b)
        r.append(rPr)
        r.append(OxmlElement('w:tab'))
        return r

    p._p.append(make_bold_run(company))
    p._p.append(make_tab_run())
    p._p.append(make_bold_run(dates))
    return p


def _add_bullet(doc, text, bold=False, style_name="List Paragraph"):
    """Add a bullet point paragraph."""
    p   = doc.add_paragraph(style=style_name)
    run = p.add_run(text)
    run.bold = bold
    return p


def _spacer(doc):
    """Add an empty spacer paragraph."""
    doc.add_paragraph(style="Normal")


def build_resume_docx(data):
    """
    Build a formatted resume DOCX from structured data.

    data keys:
        name             str
        location         str  (additional field 1)
        commute_info     str  (additional field 2, optional)
        availability     str  (additional field 3, optional)
        summary          str
        education        list of {"institution": str, "degrees": [str]}
        certifications   list of str
        skills           list of str
        experience       list of {"company": str, "dates": str, "title": str,
                                   "bullets": [str]}
    """
    doc = Document()

    # ── Page setup ──────────────────────────────────────────────────────────
    sec = doc.sections[0]
    sec.page_width   = Pt(612)   # 8.5"
    sec.page_height  = Pt(792)   # 11"
    sec.left_margin  = Inches(1)
    sec.right_margin = Inches(1)
    sec.top_margin   = Inches(1)
    sec.bottom_margin = Inches(1)

    # ── Default font (Times New Roman matches template) ──────────────────────
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE
    styles = doc.styles
    normal = styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(11)

    # ── NAME ────────────────────────────────────────────────────────────────
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(data.get("name", ""))
    run.bold      = True
    run.font.size = Pt(16)

    # ── ADDITIONAL FIELDS (location, commute, availability) ─────────────────
    fields = [
        data.get("location", ""),
        data.get("commute_info", ""),
        data.get("availability", ""),
    ]
    for field in fields:
        if field and field.strip():
            p = doc.add_paragraph(style="Normal")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(field.strip())
            run.bold      = True
            run.font.size = Pt(12)

    _spacer(doc)

    # ── SUMMARY ─────────────────────────────────────────────────────────────
    heading = doc.add_paragraph(style="Normal")
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = heading.add_run("Summary")
    r.bold = True

    if data.get("summary"):
        p = doc.add_paragraph(style="Normal")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run(data["summary"])

    _spacer(doc)

    # ── EDUCATION AND CERTIFICATIONS ────────────────────────────────────────
    heading = doc.add_paragraph(style="Normal")
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = heading.add_run("Education and Certifications")
    r.bold = True

    for edu in data.get("education", []):
        # Institution name — plain
        if edu.get("institution"):
            p = doc.add_paragraph(style="Normal")
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(edu["institution"])

        # Degrees — bold bullets
        for deg in edu.get("degrees", []):
            if deg.strip():
                _add_bullet(doc, deg.strip(), bold=True)

        _spacer(doc)

    # Certifications — bold bullets (no institution header)
    for cert in data.get("certifications", []):
        if cert.strip():
            _add_bullet(doc, cert.strip(), bold=True)

    if data.get("certifications"):
        _spacer(doc)

    # ── SKILLS ──────────────────────────────────────────────────────────────
    heading = doc.add_paragraph(style="Normal")
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = heading.add_run("Skills")
    r.bold = True

    _spacer(doc)

    for skill in data.get("skills", []):
        if skill.strip():
            _add_bullet(doc, skill.strip(), bold=False)

    _spacer(doc)
    _spacer(doc)

    # ── EXPERIENCE ──────────────────────────────────────────────────────────
    heading = doc.add_paragraph(style="Normal")
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = heading.add_run("Experience")
    r.bold = True

    for i, job in enumerate(data.get("experience", [])):
        company = job.get("company", "")
        dates   = job.get("dates", "")
        title   = job.get("title", "")
        bullets = job.get("bullets", [])

        # Company + dates (tab-separated, bold)
        if company or dates:
            _add_company_line(doc, company, dates)

        # Job title (bold italic)
        if title:
            p = doc.add_paragraph(style="Normal")
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(title)
            run.bold   = True
            run.italic = True

        # Bullet points (plain)
        for bullet in bullets:
            if bullet.strip():
                # Strip leading bullet characters
                clean = bullet.strip().lstrip("•·∙-").strip()
                _add_bullet(doc, clean, bold=False)

        if i < len(data.get("experience", [])) - 1:
            _spacer(doc)

    # ── Write to bytes ───────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
