"""
Resume Formatter — Single Client Tool
Formats any resume into the client's standard format.
Content is preserved verbatim. Only formatting changes.
"""
import streamlit as st
import os
import sys
import json

sys.path.insert(0, os.path.dirname(__file__))

from extractor import extract_text
from content_extractor import extract_content_verbatim
from resume_builder import build_resume_docx
from pdf_converter import docx_to_pdf

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def get_secret(key, fallback=""):
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, fallback)


GROQ_API_KEY = get_secret("GROQ_API_KEY")

st.set_page_config(
    page_title="Resume Formatter",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
#MainMenu, footer, .stDeployButton { display: none; visibility: hidden; }
.stApp { background: #f5f6f8; }

.app-header {
    background: #1a2744; color: white;
    padding: 1.1rem 1.8rem; border-radius: 6px; margin-bottom: 1.3rem;
}
.app-header h1 { margin: 0; font-size: 1.25rem; font-weight: 600; }
.app-header p  { margin: .15rem 0 0; font-size: .76rem; opacity: .6; }

.step-row { display: flex; gap: 3px; margin-bottom: 1.3rem; }
.step { flex: 1; text-align: center; padding: .4rem .2rem; border-radius: 3px;
        font-size: .69rem; font-weight: 500; }
.step-done    { background: #1a2744; color: white; }
.step-active  { background: #2d4a8a; color: white; }
.step-pending { background: #dde1ea; color: #6b7280; }

.ntc { padding: .55rem .85rem; border-radius: 4px; margin: .4rem 0;
       font-size: .83rem; line-height: 1.45; }
.ntc-info    { background: #eef2ff; border-left: 3px solid #2d4a8a; color: #1e2f5e; }
.ntc-warn    { background: #fffbeb; border-left: 3px solid #d97706; color: #78350f; }
.ntc-success { background: #f0fdf4; border-left: 3px solid #16a34a; color: #14532d; }
.ntc-error   { background: #fef2f2; border-left: 3px solid #dc2626; color: #7f1d1d; }

.section-label {
    font-size: .7rem; font-weight: 600; color: #6b7280;
    text-transform: uppercase; letter-spacing: .05em;
    margin: 1rem 0 .3rem;
}
.reorder-item {
    background: white; border: 1px solid #e5e7eb; border-radius: 4px;
    padding: .45rem .8rem; margin: .2rem 0; font-size: .81rem;
}

[data-testid="stSidebar"] > div:first-child {
    background: #fff; border-right: 1px solid #e5e7eb;
}
.stButton > button { border-radius: 4px; font-weight: 500; font-size: .82rem; }
hr { margin: .85rem 0; border-color: #e5e7eb; }
</style>
""", unsafe_allow_html=True)

# ── Defaults ─────────────────────────────────────────────────────────────────
DEFAULTS = {
    "step": 1,
    "raw_text": None,
    "candidate": None,
    "location": "",
    "commute_info": "",
    "availability": "",
    "output_docx": None,
    "output_pdf": None,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


def ntc(text, kind="info"):
    st.markdown(f'<div class="ntc ntc-{kind}">{text}</div>',
                unsafe_allow_html=True)


def step_bar(steps, cur):
    h = '<div class="step-row">'
    for i, s in enumerate(steps, 1):
        if i < cur:    c, t = "step-done",   f"&#10003; {s}"
        elif i == cur: c, t = "step-active",  s
        else:          c, t = "step-pending",  s
        h += f'<div class="step {c}">{t}</div>'
    return h + '</div>'


def reset():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("**Resume Formatter**")
    st.divider()
    st.caption(
        "Upload a candidate resume, fill in three fields, "
        "review the extracted data, and download the formatted output."
    )
    st.divider()
    st.caption("Format: Quest Diagnostics standard")
    st.caption("Version 1.0")
    st.divider()
    if GROQ_API_KEY:
        st.caption("API: Connected")
    else:
        st.error("GROQ_API_KEY missing. Add it in Streamlit Cloud secrets.")


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>Resume Formatter</h1>
  <p>Formats any resume into the client standard &nbsp;&middot;&nbsp;
     Content preserved verbatim &nbsp;&middot;&nbsp; Download Word + PDF</p>
</div>""", unsafe_allow_html=True)

step = st.session_state.step
st.markdown(
    step_bar(["Upload", "Additional Info", "Review & Edit", "Download"], step),
    unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 1 — Upload
# ─────────────────────────────────────────────────────────────────────────────
if step == 1:
    st.markdown("#### Step 1: Upload the candidate's resume")
    st.caption("Accepts PDF or Word (.docx / .doc). Content will be extracted verbatim — nothing will be rewritten.")

    upl = st.file_uploader(
        "Candidate resume",
        type=["pdf", "docx", "doc"],
        label_visibility="collapsed"
    )

    if upl:
        # Read file bytes ONCE and reuse — file pointer can only be read once
        file_bytes = upl.read()
        file_name  = upl.name

        with st.expander("Preview extracted text"):
            try:
                raw_preview = extract_text(file_bytes, file_name)
                st.text(raw_preview[:3000] if raw_preview.strip() else "(No text extracted)")
            except Exception as e:
                st.text(f"Preview error: {e}")

        if st.button("Extract and Continue", type="primary",
                     use_container_width=True):
            # Show API key status
            if not GROQ_API_KEY:
                st.error("GROQ_API_KEY is not configured. Add it to Streamlit secrets.")
                st.stop()

            with st.spinner("Reading resume..."):
                raw = extract_text(file_bytes, file_name)
                if not raw.strip():
                    st.error("Could not extract text from this file. Try a different format.")
                    st.stop()
                st.session_state.raw_text = raw

            with st.spinner("Extracting candidate data — this takes 10-20 seconds..."):
                try:
                    candidate = extract_content_verbatim(GROQ_API_KEY, raw)
                    st.session_state.candidate = candidate
                    st.session_state.step = 2
                    st.rerun()
                except Exception as e:
                    st.error(f"Extraction failed: {e}")
                    st.caption("Check that your GROQ_API_KEY is correctly set in Streamlit secrets.")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2 — Additional Information
# ─────────────────────────────────────────────────────────────────────────────
elif step == 2:
    st.markdown("#### Step 2: Additional Information")
    ntc(
        "These three fields are added by the recruiter and will appear directly below "
        "the candidate's name in the formatted resume. Leave any field blank to omit it.",
        "info"
    )

    candidate = st.session_state.candidate or {}
    st.session_state.location = st.text_input(
        "Candidate Location",
        value=st.session_state.location or candidate.get("location", ""),
        placeholder="e.g., Dallas, GA 30132",
        help="City, State ZIP — appears centered below the candidate's name"
    )

    st.session_state.commute_info = st.text_input(
        "Commute Information",
        value=st.session_state.commute_info,
        placeholder="e.g., (26 mins commute from quest site- 15 miles)",
        help="Leave blank to omit this line from the resume"
    )

    st.session_state.availability = st.text_input(
        "Interview Availability",
        value=st.session_state.availability,
        placeholder="e.g., Interview Availability: Any time (Monday to Friday need 24 hrs notice)",
        help="Leave blank to omit this line from the resume"
    )

    filled = sum(1 for f in [st.session_state.location,
                              st.session_state.commute_info,
                              st.session_state.availability] if f.strip())
    st.caption(f"{filled} of 3 field(s) filled.")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    with c2:
        if st.button("Continue to Review", type="primary", use_container_width=True):
            st.session_state.step = 3
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3 — Review & Edit
# ─────────────────────────────────────────────────────────────────────────────
elif step == 3:
    st.markdown("#### Step 3: Review and Edit")
    ntc(
        "All content below was extracted verbatim from the original resume. "
        "Correct any extraction errors before generating the formatted output.",
        "info"
    )

    cand = st.session_state.candidate or {}

    # Personal
    st.markdown('<div class="section-label">Personal Information</div>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        cand["name"] = st.text_input("Full Name", value=cand.get("name", ""))
    with c2:
        cand["location"] = st.text_input(
            "Location (from resume)",
            value=cand.get("location", ""),
            help="This is the location found in the resume. The recruiter-provided location from Step 2 takes priority."
        )

    # Additional fields reminder
    with st.expander("Additional fields (from Step 2)"):
        st.text(f"Location:     {st.session_state.location or '(blank)'}")
        st.text(f"Commute:      {st.session_state.commute_info or '(blank)'}")
        st.text(f"Availability: {st.session_state.availability or '(blank)'}")
        if st.button("Edit these fields", use_container_width=True, key="edit_step2"):
            st.session_state.step = 2
            st.rerun()

    # Summary
    st.markdown('<div class="section-label">Professional Summary</div>',
                unsafe_allow_html=True)
    st.caption("Appears verbatim in the resume. Edit to correct any extraction issues.")
    cand["summary"] = st.text_area(
        "Summary",
        value=cand.get("summary", ""),
        height=140,
        label_visibility="collapsed"
    )

    # Skills
    st.markdown('<div class="section-label">Skills</div>', unsafe_allow_html=True)
    skills_raw = "\n".join(cand.get("skills", []))
    skills_edited = st.text_area(
        "Skills (one per line)",
        value=skills_raw,
        height=200,
        label_visibility="collapsed",
        help="One skill per line. Order will be preserved."
    )
    cand["skills"] = [s.strip() for s in skills_edited.split("\n") if s.strip()]

    # Education
    st.markdown('<div class="section-label">Education</div>', unsafe_allow_html=True)
    for i, edu in enumerate(cand.get("education", [])):
        with st.expander(f"{edu.get('institution', 'Institution')} — "
                         f"{', '.join(edu.get('degrees', [])[:1])}"):
            edu["institution"] = st.text_input(
                "Institution", value=edu.get("institution", ""), key=f"edu_inst_{i}")
            degs_raw = "\n".join(edu.get("degrees", []))
            degs_edit = st.text_area(
                "Degrees (one per line)",
                value=degs_raw, height=80, key=f"edu_deg_{i}")
            edu["degrees"] = [d.strip() for d in degs_edit.split("\n") if d.strip()]

    # Certifications
    st.markdown('<div class="section-label">Certifications & Licenses</div>',
                unsafe_allow_html=True)
    certs_raw = "\n".join(cand.get("certifications", []))
    certs_edit = st.text_area(
        "Certifications (one per line)",
        value=certs_raw, height=140, label_visibility="collapsed")
    cand["certifications"] = [c.strip() for c in certs_edit.split("\n") if c.strip()]

    # Experience
    st.markdown('<div class="section-label">Work Experience</div>',
                unsafe_allow_html=True)
    exp = cand.get("experience", [])

    # Reorder controls
    if len(exp) > 1:
        st.caption("Drag to reorder — use Move Up / Move Down.")
        for i, job in enumerate(exp):
            st.markdown(
                f'<div class="reorder-item"><b>{job.get("title", "")}</b>'
                f' &mdash; {job.get("company", "")} ({job.get("dates", "")})</div>',
                unsafe_allow_html=True)
            r1, r2, _ = st.columns([1, 1, 5])
            with r1:
                if i > 0 and st.button("Move Up", key=f"exp_up_{i}"):
                    exp[i], exp[i-1] = exp[i-1], exp[i]
                    cand["experience"] = exp
                    st.session_state.candidate = cand
                    st.rerun()
            with r2:
                if i < len(exp)-1 and st.button("Move Down", key=f"exp_dn_{i}"):
                    exp[i], exp[i+1] = exp[i+1], exp[i]
                    cand["experience"] = exp
                    st.session_state.candidate = cand
                    st.rerun()
        st.divider()

    for i, job in enumerate(exp):
        with st.expander(f"Position {i+1}: {job.get('title', '')} at {job.get('company', '')}"):
            c1, c2, c3 = st.columns([2, 2, 2])
            job["company"] = c1.text_input("Company",    value=job.get("company", ""), key=f"jc{i}")
            job["title"]   = c2.text_input("Job Title",  value=job.get("title", ""),   key=f"jt{i}")
            job["dates"]   = c3.text_input("Dates",      value=job.get("dates", ""),   key=f"jd{i}")
            bullets_raw = "\n".join(job.get("bullets", []))
            bullets_edit = st.text_area(
                "Bullet points (one per line)",
                value=bullets_raw, height=180, key=f"jb{i}",
                help="Copy verbatim from the original resume")
            job["bullets"] = [b.strip().lstrip("•·∙-").strip()
                              for b in bullets_edit.split("\n") if b.strip()]

    cand["experience"] = exp
    st.session_state.candidate = cand

    st.divider()
    c1, c2 = st.columns([1, 2])
    with c1:
        if st.button("Back", use_container_width=True, key="rev_back"):
            st.session_state.step = 2
            st.rerun()
    with c2:
        if st.button("Generate Formatted Resume", type="primary",
                     use_container_width=True, key="rev_gen"):
            cand = st.session_state.candidate or {}

            # Assemble data dict
            data = {
                "name":        cand.get("name", ""),
                "location":    st.session_state.location.strip(),
                "commute_info":st.session_state.commute_info.strip(),
                "availability":st.session_state.availability.strip(),
                "summary":     cand.get("summary", ""),
                "education":   cand.get("education", []),
                "certifications": cand.get("certifications", []),
                "skills":      cand.get("skills", []),
                "experience":  cand.get("experience", []),
            }

            # Build DOCX
            with st.spinner("Building formatted resume..."):
                try:
                    docx_bytes = build_resume_docx(data)
                except Exception as e:
                    st.error(f"Build failed: {e}")
                    import traceback
                    st.text(traceback.format_exc())
                    pass  # error shown above

            # Save DOCX
            safe_name = "".join(
                c if c.isalnum() or c in "_-" else "_"
                for c in data["name"].replace(" ", "_")
            )
            docx_path = os.path.join(OUTPUT_DIR, f"{safe_name}_formatted.docx")
            pdf_path  = os.path.join(OUTPUT_DIR, f"{safe_name}_formatted.pdf")

            with open(docx_path, "wb") as f:
                f.write(docx_bytes)

            # Convert to PDF
            with st.spinner("Converting to PDF..."):
                pdf_ok, pdf_msg = docx_to_pdf(docx_path, pdf_path)

            st.session_state.output_docx = docx_path
            st.session_state.output_pdf  = pdf_path if pdf_ok else None
            st.session_state.step = 4
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 4 — Download
# ─────────────────────────────────────────────────────────────────────────────
elif step == 4:
    st.markdown("#### Step 4: Download")
    ntc("Resume formatted successfully. Download the files below.", "success")

    c1, c2 = st.columns(2)
    with c1:
        docx_path = st.session_state.output_docx
        if docx_path and os.path.exists(docx_path):
            with open(docx_path, "rb") as f:
                st.download_button(
                    "Download Word Document (.docx)",
                    data=f.read(),
                    file_name=os.path.basename(docx_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary",
                    key="dl_docx"
                )
    with c2:
        pdf_path = st.session_state.output_pdf
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "Download PDF",
                    data=f.read(),
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                    key="dl_pdf"
                )
        else:
            st.caption("PDF not available. Open the DOCX in Word and use File > Save As > PDF.")

    st.divider()
    if st.button("Format Another Resume", use_container_width=True, key="again"):
        reset()
        st.rerun()
