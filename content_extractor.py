"""
content_extractor.py
Extracts structured content VERBATIM from the raw resume.
No rewriting, no summarizing — only structural parsing.
"""
import json
import re
from groq import Groq

MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"

EXTRACT_SYSTEM = """You are a resume parser. Extract ALL data from the resume text.

CRITICAL RULES:
1. COPY every piece of text VERBATIM — do not rephrase, summarize, or change anything
2. Do NOT add information that is not in the resume
3. Do NOT remove or skip any information
4. Preserve the exact wording of every bullet, job title, skill, certification, etc.
5. Extract ALL skills listed — every single one
6. Extract ALL jobs listed — every single one with all their bullet points
7. Extract ALL certifications and licenses

Return ONLY valid JSON with no markdown fences, no commentary, nothing else.
The JSON must have exactly these keys:

{
  "name": "Full name",
  "location": "City, State ZIP from top of resume",
  "summary": "Complete summary/objective paragraph verbatim",
  "education": [
    {
      "institution": "School name exactly as written",
      "degrees": ["Degree exactly as written"]
    }
  ],
  "certifications": [
    "Each certification/license exactly as written"
  ],
  "skills": [
    "Each skill exactly as written — include ALL of them"
  ],
  "experience": [
    {
      "company": "Company name exactly as written",
      "dates": "Date range as written, formatted Month Year - Month Year",
      "title": "Job title exactly as written",
      "bullets": [
        "Each bullet point verbatim without the leading bullet symbol"
      ]
    }
  ]
}

Rules for specific sections:
- Put school degrees/diplomas under "education" with the school as "institution"
- Put certifications, licenses, credentials under "certifications" (not education)
- "dates" format: use dash not "to", e.g. "January 2020 - Present"
- Include EVERY bullet under each job, even one-liners
- Skills: list them one per array element, include ALL
"""


def extract_content_verbatim(api_key, raw_text):
    """
    Extract structured content from raw resume text verbatim.
    Raises ValueError with details if extraction or parsing fails.
    """
    if not api_key:
        raise ValueError("GROQ_API_KEY is not set. Add it to your Streamlit secrets.")

    if not raw_text or not raw_text.strip():
        raise ValueError("Resume text is empty. The file could not be read.")

    client = Groq(api_key=api_key)

    # Truncate very long resumes to fit context (keep first 6000 chars)
    text_to_send = raw_text[:6000] if len(raw_text) > 6000 else raw_text

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": EXTRACT_SYSTEM},
            {"role": "user",   "content": f"Parse this resume. Return only JSON:\n\n{text_to_send}"}
        ],
        max_tokens=4000,
        temperature=0.0
    )

    raw_response = response.choices[0].message.content.strip()

    # Strip markdown code fences if present
    cleaned = raw_response
    cleaned = re.sub(r'^```json\s*', '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'^```\s*',     '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'\s*```$',     '', cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()

    # Try to parse as-is
    try:
        result = json.loads(cleaned)
        # Validate it has the expected keys
        for key in ['name', 'skills', 'experience']:
            if key not in result:
                result[key] = [] if key != 'name' else ''
        return result
    except json.JSONDecodeError:
        pass

    # Try to find JSON object within the response
    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # If all parsing fails, raise with the raw response for debugging
    raise ValueError(
        f"Could not parse the AI response as JSON.\n"
        f"Response was:\n{raw_response[:800]}"
    )
