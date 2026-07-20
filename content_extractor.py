"""
content_extractor.py  v4
Extracts ALL content verbatim from any length resume/CV.
Uses multi-pass extraction for long CVs to avoid token limits.
"""
import json, re
from groq import Groq

MODELS = [
    "llama-3.3-70b-versatile",
    "llama-4-scout-instruct",
    "llama-3.1-8b-instant",
]

# ── Prompt for the STRUCTURE pass (finds all sections + basic info) ──────────
STRUCTURE_SYSTEM = """You are a resume/CV parser. Extract the complete structure and ALL content verbatim.

CRITICAL: Copy every word exactly as written. Do not skip, summarize, or rephrase anything.

Return ONLY valid JSON (no markdown fences):
{
  "name": "full name",
  "contact": "email | phone | linkedin all on one line",
  "location": "city, state zip if at top of resume",
  "summary": "full summary paragraph verbatim — empty string if no summary section",
  "education": [
    {
      "institution": "school name",
      "degrees": ["full degree line verbatim including all honors, minors, GPA, majors on one line"]
    }
  ],
  "certifications": ["each cert/license verbatim — look in Certifications, Professional Development, and any other section"],
  "skills": ["each skill item verbatim — flat list, no categories"],
  "experience_sections": [
    {
      "section_heading": "exact heading from CV e.g. Research Experience, Teaching Experience, Clinical Experience, Leadership Experience, Learning Assistant Experience, Professional Development, Other Research Experience",
      "jobs": [
        {
          "institution": "org/company name",
          "dates": "date range verbatim",
          "course": "Course: X if listed — empty string if none",
          "title": "job/role title verbatim",
          "bullets": ["each bullet verbatim without leading bullet character"]
        }
      ]
    }
  ]
}

Rules:
- experience_sections: use the EXACT section headings from the CV. Do NOT merge sections.
- For skills with sub-headings (Laboratory:, Bioinformatics:, Language:) — still put all as flat list, each item verbatim
- Certifications: look in ALL sections — Professional Development, Certifications section, any other
- Summary: this may be a manually added section not in the original resume — include if present
- course: prefix "Course: " exactly as in the document
- Include ALL jobs/entries in every section — do not truncate
"""


def _call_model(client, system, user_text, max_tokens=4000):
    last_err = None
    for model in MODELS:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role":"system","content":system},
                          {"role":"user","content":user_text}],
                max_tokens=max_tokens,
                temperature=0.0)
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            continue
    raise ValueError(f"All models failed: {last_err}")


def _parse_json(raw):
    """Strip markdown fences and parse JSON robustly."""
    cleaned = re.sub(r'^```json\s*', '', raw, flags=re.MULTILINE)
    cleaned = re.sub(r'^```\s*',     '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'\s*```$',     '', cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', cleaned, re.DOTALL)
        if m:
            return json.loads(m.group())
        raise


def _merge_results(base, extra):
    """Merge extra extraction results into base — append sections, extend lists."""
    # Merge experience_sections
    existing_headings = {s.get("section_heading","").strip().lower()
                         for s in base.get("experience_sections",[])}
    for section in extra.get("experience_sections", []):
        heading = section.get("section_heading","").strip()
        if heading.lower() not in existing_headings and section.get("jobs"):
            base.setdefault("experience_sections", []).append(section)
            existing_headings.add(heading.lower())
        else:
            # Merge jobs into existing section
            for s in base.get("experience_sections", []):
                if s.get("section_heading","").strip().lower() == heading.lower():
                    existing_titles = {j.get("title","") for j in s.get("jobs",[])}
                    for job in section.get("jobs", []):
                        if job.get("title","") not in existing_titles:
                            s.setdefault("jobs", []).append(job)
                    break

    # Merge skills
    existing_skills = set(base.get("skills", []))
    for skill in extra.get("skills", []):
        if skill not in existing_skills:
            base.setdefault("skills", []).append(skill)
            existing_skills.add(skill)

    # Merge certifications
    existing_certs = set(base.get("certifications", []))
    for cert in extra.get("certifications", []):
        if cert not in existing_certs:
            base.setdefault("certifications", []).append(cert)
            existing_certs.add(cert)

    # Fill missing top-level fields
    for field in ["summary", "location", "contact", "name"]:
        if not base.get(field) and extra.get(field):
            base[field] = extra[field]

    # Merge education
    if not base.get("education") and extra.get("education"):
        base["education"] = extra["education"]

    return base


def extract_content_verbatim(api_key, raw_text):
    """
    Extract ALL content verbatim from resume. Uses chunked extraction
    for long CVs so nothing is truncated.
    """
    if not api_key:
        raise ValueError("GROQ_API_KEY not set. Add to Streamlit secrets.")
    if not raw_text or not raw_text.strip():
        raise ValueError("Resume text is empty.")

    client = Groq(api_key=api_key)
    CHUNK = 6000   # chars per chunk — well within token limits
    OVERLAP = 500  # overlap to catch sections split across chunks

    text = raw_text.strip()
    total = len(text)

    if total <= CHUNK:
        # Short CV — single pass
        chunks = [text]
    else:
        # Long CV — split into overlapping chunks
        chunks = []
        pos = 0
        while pos < total:
            end = min(pos + CHUNK, total)
            chunks.append(text[pos:end])
            if end == total:
                break
            pos = end - OVERLAP

    result = None
    for i, chunk in enumerate(chunks):
        prompt = (f"Parse this {'complete' if len(chunks)==1 else f'part {i+1} of {len(chunks)}'} "
                  f"resume and return JSON:\n\n{chunk}")
        try:
            raw_resp = _call_model(client, STRUCTURE_SYSTEM, prompt)
            parsed   = _parse_json(raw_resp)
            if result is None:
                result = parsed
            else:
                result = _merge_results(result, parsed)
        except Exception as e:
            if result is None:
                raise ValueError(f"Extraction failed on chunk {i+1}: {e}")
            # Partial failure — continue with what we have

    if result is None:
        raise ValueError("Extraction returned no data.")

    # Ensure required keys
    for key, default in [('name',''),('contact',''),('location',''),('summary',''),
                         ('education',[]),('certifications',[]),('skills',[]),
                         ('experience_sections',[])]:
        if key not in result:
            result[key] = default

    return result
