"""
content_extractor.py
Extracts structured content VERBATIM from the raw resume.
No rewriting, no summarizing — only structural parsing.
The output preserves the candidate's exact words.
"""
import json
import re
from groq import Groq

MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"

EXTRACT_SYSTEM = """You are a resume parser. Extract structured data from the resume text.

CRITICAL RULES:
1. COPY text VERBATIM — do not rephrase, summarize, improve, or change ANY content
2. Do not add information that is not in the resume
3. Do not remove information from the resume
4. Preserve the exact wording of every bullet point, job title, skill, etc.
5. If a field is not present, use an empty string or empty list

Return ONLY this JSON (no markdown fences):
{
  "name": "Full name from resume",
  "location": "City, State ZIP if present at top of resume",
  "summary": "Full verbatim summary/objective text",
  "education": [
    {
      "institution": "School or college name exactly as written",
      "degrees": ["Degree name exactly as written", "second degree if any"]
    }
  ],
  "certifications": [
    "Certification name exactly as written (include expiry if shown)"
  ],
  "skills": [
    "Each skill exactly as written"
  ],
  "experience": [
    {
      "company": "Company name exactly as written",
      "location": "City, State if shown",
      "dates": "Date range exactly as written (e.g. January 2008 to December 2011)",
      "title": "Job title exactly as written",
      "bullets": [
        "Each bullet point verbatim, without the leading bullet character"
      ]
    }
  ]
}

IMPORTANT: 
- "dates" should be formatted as "Month Year - Month Year" or "Month Year - Present"
- For skills, include EVERY skill listed, no filtering
- For bullets, copy the exact text word-for-word
- Education and certifications may be listed together — separate them correctly
- Put certifications/licenses in the "certifications" array, not education
- Put degrees/diplomas in the "education" array under the institution
"""


def extract_content_verbatim(api_key, raw_text):
    """
    Extract structured content from raw resume text.
    Returns a dict with name, summary, education, certifications, skills, experience.
    Content is verbatim — not rewritten.
    """
    client = Groq(api_key=api_key)

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": EXTRACT_SYSTEM},
            {"role": "user", "content": f"Parse this resume and return JSON:\n\n{raw_text}"}
        ],
        max_tokens=4000,
        temperature=0.0
    )

    raw = response.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            return json.loads(m.group())
        raise ValueError(f"Could not parse extraction JSON:\n{raw[:500]}")
