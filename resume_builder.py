"""
resume_builder.py  v3
Matches the gold-standard manual formatted resume exactly.

Key format rules (verified from manual gold standard):
- Default: Times New Roman, 11pt
- Name: 16pt Bold Center
- Additional fields (location, commute, availability):
    Combined into ONE paragraph with line breaks (w:br) between them, 12pt Bold Center
    Contact email/phone NOT shown (recruiter-added fields only)
- Summary: first word bold, rest plain, Justify alignment
- Section headings (Summary, Education and Certifications, Skills, Experience): Bold Left
- Skills: one empty List Paragraph spacer after heading, then plain bullets
- Two spacers before Experience
- Company line: Bold Left with right-tab for dates
- Job title: Bold Italic Left
- Bullets: PLAIN (not bold), List Paragraph
- Employment gap entries: treated as jobs (company=gap label, title=description)
- No "Professional Experience" sub-heading
- No contact line in output
"""
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io


def _set_default_font(doc):
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
    rPr = style.element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'),    'Times New Roman')
    # Insert at beginning to not override existing
    existing = rPr.find(qn('w:rFonts'))
    if existing is None:
        rPr.insert(0, rFonts)


def _spacer(doc, style="Normal"):
    doc.add_paragraph(style=style)


def _plain_para(doc, text, bold=False, italic=False, size_pt=None,
                align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph(style="Normal")
    p.alignment = align
    if not text:
        return p
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    if size_pt:
        run.font.size = Pt(size_pt)
    return p


def _summary_para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Summary: first word bold, rest plain — matches gold standard."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = align
    if not text:
        return p
    words = text.split(' ', 1)
    r1 = p.add_run(words[0])
    r1.bold = True
    if len(words) > 1:
        r2 = p.add_run(' ' + words[1])
        r2.bold = False
    return p


def _bullet(doc, text, bold=False):
    """List Paragraph bullet — plain by default."""
    p   = doc.add_paragraph(style="List Paragraph")
    run = p.add_run(text)
    run.bold = bold
    return p


def _company_line(doc, company, dates):
    """Bold left with right-aligned dates via tab at 9360 twips."""
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
        if txt and (txt.startswith(' ') or txt.endswith(' ')):
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

    p._p.append(bold_run(company))
    if dates:
        p._p.append(tab_run())
        p._p.append(bold_run(dates))
    return p


def _additional_fields_para(doc, fields):
    """
    Combine non-empty additional fields into ONE paragraph with line breaks
    between them. 12pt Bold Center. Matches gold standard P001 structure.
    """
    non_empty = [f.strip() for f in fields if f and f.strip()]
    if not non_empty:
        return
    p   = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, field in enumerate(non_empty):
        run = p.add_run(field)
        run.bold = True
        run.font.size = Pt(12)
        if idx < len(non_empty) - 1:
            # Add a line break (w:br) between fields
            br  = OxmlElement('w:br')
            run._r.append(br)


def build_resume_docx(data):
    """
    Build formatted resume.

    data keys:
        name                str
        location            str   (additional field 1)
        commute_info        str   (additional field 2)
        availability        str   (additional field 3)
        summary             str
        education           list of {institution, degrees:[str]}
        certifications      list of str
        skills              list of str  (flat)
        experience_sections list of {section_heading, jobs:[{institution, dates, course, title, bullets}]}
    """
    doc = Document()
    _set_default_font(doc)

    sec = doc.sections[0]
    sec.left_margin   = Inches(1)
    sec.right_margin  = Inches(1)
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)

    # ── NAME ────────────────────────────────────────────────────────────────
    _plain_para(doc, data.get("name", ""),
                bold=True, size_pt=16, align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── ADDITIONAL FIELDS — combined single paragraph with line breaks ───────
    _additional_fields_para(doc, [
        data.get("location", ""),
        data.get("commute_info", ""),
        data.get("availability", ""),
    ])

    _spacer(doc)

    # ── SUMMARY ─────────────────────────────────────────────────────────────
    summary = data.get("summary", "").strip()
    if summary:
        _plain_para(doc, "Summary", bold=True)
        _summary_para(doc, summary)
        _spacer(doc)

    # ── EDUCATION AND CERTIFICATIONS ────────────────────────────────────────
    education = data.get("education", [])
    certs     = data.get("certifications", [])
    if education or certs:
        _plain_para(doc, "Education and Certifications", bold=True)
        for edu in education:
            inst = edu.get("institution", "").strip()
            if inst:
                _plain_para(doc, inst)
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
        _plain_para(doc, "Skills", bold=True)
        # Gold standard has one empty List Paragraph spacer after Skills heading
        _spacer(doc, style="List Paragraph")
        for skill in skills:
            s = skill.strip() if isinstance(skill, str) else ""
            if s:
                _bullet(doc, s, bold=False)  # plain bullets
        _spacer(doc)
        _spacer(doc)  # two spacers before Experience

    # ── EXPERIENCE ──────────────────────────────────────────────────────────
    exp_sections = data.get("experience_sections", [])
    if exp_sections:
        _plain_para(doc, "Experience", bold=True)
        # No sub-heading like "Professional Experience" — go straight to jobs

        for section in exp_sections:
            jobs = section.get("jobs", [])
            for job in jobs:
                institution = job.get("institution", "").strip()
                dates       = job.get("dates", "").strip()
                course      = job.get("course", "").strip()
                title       = job.get("title", "").strip()
                bullets     = job.get("bullets", [])

                if institution or dates:
                    _company_line(doc, institution, dates)

                if course:
                    _plain_para(doc, course, bold=True)

                if title:
                    p   = doc.add_paragraph(style="Normal")
                    run = p.add_run(title)
                    run.bold   = True
                    run.italic = True

                for b in bullets:
                    clean = b.strip().lstrip("•·∙-–").strip()
                    if clean:
                        _bullet(doc, clean, bold=False)  # plain bullets

                _spacer(doc)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
