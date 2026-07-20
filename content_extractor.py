"""
content_extractor.py
Extracts ALL structured content VERBATIM from any resume.
Handles complex academic/research CVs with many sub-sections.
"""
import json
import re
from groq import Groq

MODELS = [
    "llama-3.3-70b-versatile",
    "llama-4-scout-instruct",
    "llama-3.1-8b-instant",
]

EXTRACT_SYSTEM = """You are a resume/CV parser. Extract ALL content VERBATIM from the resume.

CRITICAL RULES:
1. Copy EVERY piece of text exactly as written — do NOT rephrase, summarize, or skip anything
2. Include ALL sections found: Education, Skills, Experience, Research, Teaching, Leadership, Publications, Clinical, Professional Development, Certifications, Awards — whatever is in the resume
3. Include ALL sub-headings within each section (e.g. "Research Experience", "Teaching Experience" inside Experience)
4. For skills with sub-categories (e.g. "Laboratory:", "Bioinformatics:"), preserve the category name and its items
5. For experience entries, include the course line if present (e.g. "Course: Make Your Mutant")
6. Preserve exact dates as written

Return ONLY valid JSON with NO markdown fences. Use this exact structure:

{
  "name": "Full name",
  "contact": "email | phone | linkedin — all on one line as written",
  "location": "City, State ZIP if shown",
  "summary": "Full summary paragraph verbatim — if no summary exists leave empty string",
  "education": [
    {
      "institution": "School name",
      "degrees": [
        "Full degree line verbatim",
        "Second degree or honor or minor verbatim"
      ]
    }
  ],
  "certifications": ["Each certification/license verbatim"],
  "skills": ["skill 1 verbatim", "skill 2 verbatim", "skill 3 verbatim"],
  "experience_sections": [
    {
      "section_heading": "Section heading (e.g. Research Experience, Teaching Experience, Clinical Experience)",
      "jobs": [
        {
          "institution": "Organization/company name",
          "dates": "Date range as written",
          "course": "Course name if listed (e.g. Course: Make Your Mutant) — empty string if none",
          "title": "Job/role title verbatim",
          "bullets": ["bullet 1 verbatim", "bullet 2 verbatim"]
        }
      ]
    }
  ]
}

IMPORTANT for skills:
- List ALL skills as a flat array — do NOT create sub-categories or groups
- Extract every single skill exactly as written

IMPORTANT for experience:
- Group experience entries under their section heading
- Common sections: Research Experience, Other Research Experience, Teaching Experience, Learning Assistant Experience, Leadership Experience, Clinical Experience, Professional Development
- If resume has no section groupings, put everything under "Experience"
- Include the institution/company on the job, not just the role

IMPORTANT for education:
- Put degrees, honors, GPA, minors, certificates under the institution as separate degree lines
- Put certifications/licenses in the certifications array, NOT education
"""


def extract_content_verbatim(api_key, raw_text):
    if not api_key:
        raise ValueError("GROQ_API_KEY is not set. Add it to your Streamlit secrets.")
    if not raw_text or not raw_text.strip():
        raise ValueError("Resume text is empty. The file could not be read.")

    client = Groq(api_key=api_key)
    text_to_send = raw_text[:7000] if len(raw_text) > 7000 else raw_text

    last_error = None
    raw_response = None
    for model in MODELS:
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": EXTRACT_SYSTEM},
                    {"role": "user",   "content": f"Parse this resume completely. Return only JSON:\n\n{text_to_send}"}
                ],
                max_tokens=4000,
                temperature=0.0
            )
            raw_response = response.choices[0].message.content.strip()
            break
        except Exception as e:
            last_error = e
            continue
    else:
        raise ValueError(f"All models failed. Last error: {last_error}\nCheck GROQ_API_KEY at console.groq.com")

    # Strip markdown fences
    cleaned = re.sub(r'^```json\s*', '', raw_response, flags=re.MULTILINE)
    cleaned = re.sub(r'^```\s*',     '', cleaned,      flags=re.MULTILINE)
    cleaned = re.sub(r'\s*```$',     '', cleaned,      flags=re.MULTILINE)
    cleaned = cleaned.strip()

    try:
        result = json.loads(cleaned)
        # Ensure required keys exist
        for key, default in [('name',''), ('contact',''), ('location',''),
                              ('summary',''), ('education',[]), ('certifications',[]),
                              ('skills',[]), ('experience_sections',[])]:
            if key not in result:
                result[key] = default
        return result
    except json.JSONDecodeError:
        match = re.search(r'\{.*\}', cleaned, re.DOTALL)
        if match:
            try:
                return json.loads(match.group())
            except json.JSONDecodeError:
                pass
        raise ValueError(f"Could not parse AI response as JSON.\nResponse:\n{raw_response[:800]}")
