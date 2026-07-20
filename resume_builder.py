"""
resume_builder.py  v2
Builds formatted DOCX matching the gold-standard manually formatted resume exactly.

Format (from gold standard analysis):
- Default font: Times New Roman, 11pt (via Normal style)
- Name: 16pt Bold Center
- Additional fields (location/commute/availability): 12pt Bold Center
- Section headings (Summary, Education, Skills, Experience, sub-headings): Bold Left, default size
- Sub-section headings (Research Experience, Teaching Experience etc): Bold Left
- Summary body: Justify
- Institution name: Plain Left
- Degree bullets: Bold, List Paragraph
- Certification bullets: Bold, List Paragraph
- Skill category labels (Laboratory:, Bioinformatics:): Bold Left
- Skill bullets: Plain, List Paragraph
- Company/org line: Bold Left with right-tab at 9360 twips for dates
- Job title line: Bold Italic Left
- Course line: Bold Left
- Bullet points under jobs: Plain, List Paragraph
- Empty spacer paragraphs between logical groups
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy
import io


# ── helpers ──────────────────────────────────────────────────────────────────

def _set_default_font(doc):
    """Set Times New Roman 11pt as the document default."""
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
    # Also set via XML for full compatibility
    from docx.oxml.ns import qn
    rPr = style.element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'),    'Times New Roman')
    rPr.insert(0, rFonts)


def _para(doc, text="", bold=False, italic=False, size_pt=None,
          align=WD_ALIGN_PARAGRAPH.LEFT, style="Normal"):
    p   = doc.add_paragraph(style=style)
    p.alignment = align
    if not text:
        return p
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    if size_pt:
        run.font.size = Pt(size_pt)
    return p


def _spacer(doc):
    doc.add_paragraph(style="Normal")


def _bullet(doc, text, bold=False):
    p   = doc.add_paragraph(style="List Paragraph")
    run = p.add_run(text)
    run.bold = bold
    return p


def _company_line(doc, institution, dates):
    """Bold left text with right-aligned dates via tab stop at 9360 twips (6.5")."""
    p   = doc.add_paragraph(style="Normal")
    pPr = p._p.get_or_add_pPr()

    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), '9360')
    tabs.append(tab)
    pPr.append(tabs)

    def bold_run(txt):
        r   = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:b'))
        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = txt
        if txt.startswith(' ') or txt.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        return r

    def tab_run():
        r   = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:b'))
        r.append(rPr)
        r.append(OxmlElement('w:tab'))
        return r

    p._p.append(bold_run(institution))
    if dates:
        p._p.append(tab_run())
        p._p.append(bold_run(dates))
    return p


# ── main builder ─────────────────────────────────────────────────────────────

def build_resume_docx(data):
    """
    Build formatted resume DOCX from extracted data dict.

    data keys:
        name                str
        contact             str   (email | phone | linkedin)
        location            str   (additional field 1)
        commute_info        str   (additional field 2)
        availability        str   (additional field 3)
        summary             str
        education           list of {institution, degrees:[str]}
        certifications      list of str
        skills              list of str   (flat skill list)
        experience_sections list of {section_heading, jobs:[{institution, dates, course, title, bullets}]}
    """
    doc = Document()
    _set_default_font(doc)

    # Page setup — 1" margins all around
    sec = doc.sections[0]
    sec.left_margin   = Inches(1)
    sec.right_margin  = Inches(1)
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)

    # ── NAME ────────────────────────────────────────────────────────────────
    _para(doc, data.get("name",""), bold=True, size_pt=16,
          align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── CONTACT LINE ────────────────────────────────────────────────────────
    contact = data.get("contact","").strip()
    if contact:
        _para(doc, contact, bold=False, size_pt=11,
              align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── ADDITIONAL FIELDS ───────────────────────────────────────────────────
    for field in [data.get("location",""), data.get("commute_info",""), data.get("availability","")]:
        if field and field.strip():
            _para(doc, field.strip(), bold=True, size_pt=12,
                  align=WD_ALIGN_PARAGRAPH.CENTER)

    _spacer(doc)

    # ── SUMMARY ─────────────────────────────────────────────────────────────
    summary = data.get("summary","").strip()
    if summary:
        _para(doc, "Summary", bold=True)
        _para(doc, summary, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        _spacer(doc)

    # ── EDUCATION AND CERTIFICATIONS ────────────────────────────────────────
    education    = data.get("education", [])
    certs        = data.get("certifications", [])
    has_edu_cert = bool(education or certs)

    if has_edu_cert:
        _para(doc, "Education and Certifications", bold=True)

        for edu in education:
            inst = edu.get("institution","").strip()
            if inst:
                _para(doc, inst)  # plain, left
            for deg in edu.get("degrees", []):
                if deg.strip():
                    _bullet(doc, deg.strip(), bold=True)
            _spacer(doc)

        for cert in certs:
            if cert.strip():
                _bullet(doc, cert.strip(), bold=True)

        if certs:
            _spacer(doc)

    # ── SKILLS ──────────────────────────────────────────────────────────────
    skills = data.get("skills", [])
    if skills:
        _para(doc, "Skills", bold=True)

        # Flat skill bullets — no categories
        for skill in skills:
            if isinstance(skill, dict):
                # Handle legacy grouped format just in case
                for item in skill.get("items", []):
                    if item.strip():
                        _bullet(doc, item.strip(), bold=False)
            elif isinstance(skill, str) and skill.strip():
                _bullet(doc, skill.strip(), bold=False)

        _spacer(doc)

    # ── EXPERIENCE (with sub-sections) ──────────────────────────────────────
    exp_sections = data.get("experience_sections", [])
    if exp_sections:
        _para(doc, "Experience", bold=True)

        for section in exp_sections:
            heading = section.get("section_heading","").strip()
            jobs    = section.get("jobs", [])

            if heading and heading.lower() != "experience":
                _para(doc, heading, bold=True)

            for job in jobs:
                institution = job.get("institution","").strip()
                dates       = job.get("dates","").strip()
                course      = job.get("course","").strip()
                title       = job.get("title","").strip()
                bullets     = job.get("bullets", [])

                # Institution + dates line (tab-separated, bold)
                if institution or dates:
                    _company_line(doc, institution, dates)

                # Course line if present
                if course:
                    _para(doc, course, bold=True)

                # Job title — bold italic
                if title:
                    p   = doc.add_paragraph(style="Normal")
                    run = p.add_run(title)
                    run.bold   = True
                    run.italic = True

                # Bullets — plain
                for b in bullets:
                    clean = b.strip().lstrip("•·∙-–").strip()
                    if clean:
                        _bullet(doc, clean, bold=False)

                _spacer(doc)

    # Write to bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
