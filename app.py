import io
import re
import json
import base64
from pathlib import Path

import streamlit as st
import fitz
import requests
from docx import Document

# ------------------------------------------------------------
# ResumePilot 3.1 - Cloud AI beta
# ------------------------------------------------------------
st.set_page_config(
    page_title="ResumePilot 3.1",
    page_icon="📄",
    layout="wide"
)

st.title("📄 ResumePilot 3.1")
st.caption("Job-specific resume tailoring with Gemini/Groq cloud AI")

st.warning(
    "Beta limitations: DOCX edits are applied to existing paragraphs and may "
    "still affect complex formatting or pagination. PDF files are currently "
    "previewed and retained unchanged; PDF text editing is not enabled."
)


# ------------------------------------------------------------
# Resume extraction and scoring
# ------------------------------------------------------------
def extract_pdf(data):
    with fitz.open(stream=data, filetype="pdf") as pdf:
        return "\n".join(page.get_text("text") for page in pdf)


def extract_docx(data):
    doc = Document(io.BytesIO(data))
    pieces = [p.text for p in doc.paragraphs]

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                pieces.append(cell.text)

    return "\n".join(pieces)


def token_set(text):
    return set(
        re.findall(r"[a-zA-Z][a-zA-Z0-9+#.-]{1,}", text.lower())
    )


def keyword_score(resume_text, jd_text):
    jd_words = token_set(jd_text)
    if not jd_words:
        return 0, []

    matched = sorted(token_set(resume_text) & jd_words)
    score = round(100 * len(matched) / len(jd_words))
    return score, matched


# ------------------------------------------------------------
# Shared prompt and AI providers
# ------------------------------------------------------------
def build_resume_prompt(resume_text, jd_text):
    return f"""
You are a careful professional resume editor.

Return valid JSON with exactly these keys:
{{
  "summary": "revised summary or empty string",
  "skills": ["skills already evidenced in the source resume"],
  "bullet_rewrites": [
    {{
      "exact_source": "an exact complete existing resume paragraph",
      "revised": "a clearer version preserving the same facts"
    }}
  ]
}}

STRICT FACTUAL RULES:
- Use only facts explicitly present in the source resume.
- Never invent employers, job titles, dates, skills, qualifications,
  certifications, metrics, achievements, tools or responsibilities.
- Do not claim a job-description skill unless the resume supports it.
- Preserve the candidate's real experience, seniority and meaning.
- Keep the original section order.
- Rewrite at most 5 existing experience paragraphs.
- Each exact_source must match an entire existing paragraph exactly.
- If no safe rewrite is possible, return an empty array.
- Do not add commentary outside the JSON object.

JOB DESCRIPTION:
{jd_text[:18000]}

SOURCE RESUME:
{resume_text[:24000]}
"""


def parse_json_response(raw):
    raw = raw.strip()
    raw = re.sub(
        r"^```(?:json)?\s*|\s*```$",
        "",
        raw,
        flags=re.IGNORECASE
    )
    result = json.loads(raw)

    if not isinstance(result, dict):
        raise ValueError("AI response was not a JSON object.")

    result.setdefault("summary", "")
    result.setdefault("skills", [])
    result.setdefault("bullet_rewrites", [])

    if not isinstance(result["bullet_rewrites"], list):
        result["bullet_rewrites"] = []
    if not isinstance(result["skills"], list):
        result["skills"] = []

    return result


def call_gemini(api_key, resume_text, jd_text, model):
    prompt = build_resume_prompt(resume_text, jd_text)
    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{model}:generateContent"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json"
        }
    }

    response = requests.post(
        url,
        params={"key": api_key},
        json=payload,
        timeout=90
    )
    response.raise_for_status()

    raw = response.json()["candidates"][0]["content"]["parts"][0]["text"]
    return parse_json_response(raw)


def call_groq(api_key, resume_text, jd_text):
    prompt = build_resume_prompt(resume_text, jd_text)
    url = "https://api.groq.com/openai/v1/chat/completions"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "openai/gpt-oss-120b",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
        "max_completion_tokens": 4000
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=90
    )
    response.raise_for_status()

    raw = response.json()["choices"][0]["message"]["content"]
    return parse_json_response(raw)


def call_ai(resume_text, jd_text, gemini_key, groq_key, gemini_model):
    errors = []

    if gemini_key:
        try:
            result = call_gemini(
                gemini_key,
                resume_text,
                jd_text,
                gemini_model
            )
            return result, "Gemini"
        except Exception as exc:
            errors.append(f"Gemini: {str(exc)[:300]}")

    if groq_key:
        try:
            result = call_groq(
                groq_key,
                resume_text,
                jd_text
            )
            return result, "Groq"
        except Exception as exc:
            errors.append(f"Groq: {str(exc)[:300]}")

    if not gemini_key and not groq_key:
        raise RuntimeError(
            "No AI provider key is configured in Streamlit Secrets."
        )

    raise RuntimeError(
        "All configured AI providers failed. " + " | ".join(errors)
    )


# ------------------------------------------------------------
# DOCX editing - retain existing paragraph style as far as possible
# ------------------------------------------------------------
def replace_paragraph_text_keep_first_run_style(paragraph, new_text):
    runs = paragraph.runs

    if runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def all_docx_paragraphs(doc):
    paragraphs = list(doc.paragraphs)

    def collect_table_paragraphs(table):
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
                for nested_table in cell.tables:
                    collect_table_paragraphs(nested_table)

    for table in doc.tables:
        collect_table_paragraphs(table)

    return paragraphs


def apply_docx_edits(data, edits):
    doc = Document(io.BytesIO(data))
    paragraphs = all_docx_paragraphs(doc)

    text_to_paragraphs = {}
    for paragraph in paragraphs:
        text_to_paragraphs.setdefault(
            paragraph.text.strip(), []
        ).append(paragraph)

    applied = []
    skipped = []

    summary = (edits.get("summary") or "").strip()

    if summary:
        headings = {
            "summary",
            "professional summary",
            "profile",
            "career objective",
            "professional profile"
        }

        target = None

        for index, paragraph in enumerate(paragraphs):
            heading = paragraph.text.strip().lower().rstrip(":")
            if heading in headings:
                for following in paragraphs[index + 1:index + 5]:
                    if following.text.strip():
                        target = following
                        break
                if target:
                    break

        if target:
            replace_paragraph_text_keep_first_run_style(target, summary)
            applied.append("Professional summary")
        else:
            skipped.append(
                "Summary not inserted: no clearly labelled summary section found."
            )

    for item in edits.get("bullet_rewrites", [])[:5]:
        source = (item.get("exact_source") or "").strip()
        revised = (item.get("revised") or "").strip()
        matches = text_to_paragraphs.get(source, [])

        if source and revised and matches:
            replace_paragraph_text_keep_first_run_style(
                matches[0], revised
            )
            applied.append("Experience bullet")
        elif source:
            skipped.append(
                "A proposed bullet was not an exact paragraph match; it was not changed."
            )

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), applied, skipped


# ------------------------------------------------------------
# PDF preview
# ------------------------------------------------------------
def pdf_preview_html(data):
    encoded = base64.b64encode(data).decode("ascii")
    return (
        '<iframe '
        f'src="data:application/pdf;base64,{encoded}" '
        'width="100%" height="760" type="application/pdf">'
        '</iframe>'
    )


# ------------------------------------------------------------
# Sidebar and secure Streamlit Secrets
# ------------------------------------------------------------
with st.sidebar:
    st.header("1. Master resume")
    resume_file = st.file_uploader(
        "Upload PDF or DOCX",
        type=["pdf", "docx"]
    )

    st.header("2. Job description")
    jd_file = st.file_uploader(
        "Optional TXT job description",
        type=["txt"]
    )
    jd_text = st.text_area(
        "Paste job description",
        height=230
    )

    if jd_file:
        uploaded_jd = jd_file.getvalue().decode(
            "utf-8", errors="ignore"
        )
        jd_text = (jd_text + "\n" + uploaded_jd).strip()

    st.header("3. AI configuration")

    # Keys are read from Streamlit Cloud Secrets, never from GitHub code.
    gemini_key = st.secrets.get("GEMINI_API_KEY", "")
    groq_key = st.secrets.get("GROQ_API_KEY", "")
    gemini_model = "gemini-2.5-flash"

    if gemini_key:
        st.success("Gemini configured")
    else:
        st.info("Gemini not configured")

    if groq_key:
        st.success("Groq configured")
    else:
        st.warning("Groq not configured")

    if not gemini_key and not groq_key:
        st.error("No AI provider configured. Add keys in Streamlit Secrets.")


# ------------------------------------------------------------
# Main application
# ------------------------------------------------------------
if resume_file:
    source = resume_file.getvalue()
    extension = Path(resume_file.name).suffix.lower()

    try:
        if extension == ".pdf":
            resume_text = extract_pdf(source)
        else:
            resume_text = extract_docx(source)
    except Exception as exc:
        st.error(f"Could not read the resume: {exc}")
        st.stop()

    left_col, right_col = st.columns(2)

    with left_col:
        st.subheader("Master resume")
        st.write(resume_file.name)
        st.caption(f"{len(source):,} bytes")

    with right_col:
        st.subheader("Job match indicator")

        if jd_text.strip():
            score, matched = keyword_score(resume_text, jd_text)
            st.metric("Keyword overlap", f"{score}%")
            st.caption(
                "Keyword overlap only. This is not an ATS score "
                "or a probability of being shortlisted."
            )
            st.write(
                "Shared terms:",
                ", ".join(matched[:35]) if matched else "None detected"
            )
        else:
            st.info("Paste a job description to calculate keyword overlap.")

    st.divider()

    if st.button(
        "Analyze and tailor resume",
        type="primary",
        disabled=not (
            jd_text.strip() and (gemini_key or groq_key)
        )
    ):
        try:
            with st.spinner(
                "Analyzing resume using Gemini/Groq cloud AI..."
            ):
                edits, provider = call_ai(
                    resume_text,
                    jd_text,
                    gemini_key,
                    groq_key,
                    gemini_model
                )

                st.session_state["rp_edits"] = edits
                st.session_state["rp_provider"] = provider
                st.session_state["rp_source_name"] = resume_file.name

                if extension == ".docx":
                    output, applied, skipped = apply_docx_edits(
                        source, edits
                    )
                    st.session_state["rp_output"] = output
                    st.session_state["rp_applied"] = applied
                    st.session_state["rp_skipped"] = skipped
                    st.session_state["rp_output_ext"] = "docx"
                else:
                    st.session_state["rp_output"] = source
                    st.session_state["rp_applied"] = []
                    st.session_state["rp_skipped"] = [
                        "PDF tailoring is not enabled in this beta. "
                        "The original PDF is retained unchanged."
                    ]
                    st.session_state["rp_output_ext"] = "pdf"

            st.success(
                f"Analysis completed using {provider}. "
                "Review the proposed edits below."
            )

        except Exception as exc:
            st.error(f"Resume tailoring failed: {str(exc)}")

    edits = st.session_state.get("rp_edits")

    if (
        edits
        and st.session_state.get("rp_source_name") == resume_file.name
    ):
        st.subheader("Proposed tailored content")

        if edits.get("summary"):
            st.markdown("**Revised professional summary**")
            st.write(edits["summary"])

        if edits.get("skills"):
            st.markdown("**Relevant existing skills identified by AI**")
            st.write(", ".join(edits["skills"]))

        if edits.get("bullet_rewrites"):
            st.markdown("**Proposed experience bullet changes**")
            for item in edits["bullet_rewrites"]:
                with st.expander(
                    item.get("exact_source", "Original bullet")
                ):
                    st.write("**Suggested wording**")
                    st.write(item.get("revised", ""))

        applied = st.session_state.get("rp_applied", [])
        if applied:
            st.success("Applied to DOCX: " + ", ".join(applied))

        for warning in st.session_state.get("rp_skipped", []):
            st.warning(warning)

        st.subheader("Modified resume preview")

        if extension == ".pdf":
            st.components.v1.html(
                pdf_preview_html(source),
                height=780,
                scrolling=True
            )
            st.caption(
                "This is the unchanged original PDF preview. "
                "PDF tailoring is not enabled in this beta."
            )
            st.download_button(
                "Download original PDF",
                source,
                file_name="ResumePilot_original.pdf",
                mime="application/pdf"
            )
        else:
            preview_doc = Document(
                io.BytesIO(st.session_state["rp_output"])
            )
            preview_text = "\n\n".join(
                paragraph.text for paragraph in preview_doc.paragraphs
            )
            st.text_area(
                "DOCX text preview",
                preview_text,
                height=380
            )
            st.caption(
                "This is a text preview, not a page-accurate rendering. "
                "Open the downloaded DOCX in Word to verify layout."
            )
            st.download_button(
                "Download tailored DOCX",
                st.session_state["rp_output"],
                file_name="ResumePilot_tailored.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )

else:
    st.markdown("Upload a PDF or DOCX resume in the sidebar to begin.")
    st.markdown(
        """
        **Workflow**
        1. Upload a master resume.
        2. Paste the job description.
        3. Configure Gemini and/or Groq in Streamlit Secrets.
        4. Analyze and review proposed edits.
        5. Download the tailored DOCX when using a DOCX source.
        """
    )

st.divider()
st.caption(
    "ResumePilot 3.1 beta. Gemini is tried first when configured; Groq is "
    "used if Gemini fails. DOCX editing preserves paragraph-level styling "
    "where possible but may affect complex formatting. PDF editing, "
    "bulk URL processing, authentication, persistent storage and "
    "production multi-user controls are not included in this build."
)
