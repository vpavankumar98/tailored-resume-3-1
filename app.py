import io, os, re, json, base64
from pathlib import Path
import streamlit as st
import fitz
import requests
from docx import Document
from docx.text.paragraph import Paragraph

st.set_page_config(page_title="ResumePilot 3.1", page_icon="📄", layout="wide")
st.title("📄 ResumePilot 3.1")
st.caption("Fact-checked job-specific resume tailoring with a preview")

st.warning(
    "This is a controlled beta. DOCX tailoring edits selected existing paragraphs while retaining their paragraph styles. "
    "PDFs are previewed and retained unchanged in this build: safe in-place PDF text replacement needs document-specific "
    "font and layout validation. Do not use this beta as a production multi-user service."
)

def extract_pdf(data):
    with fitz.open(stream=data, filetype="pdf") as doc:
        return "\n".join(page.get_text("text") for page in doc)

def extract_docx(data):
    doc = Document(io.BytesIO(data))
    out = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            out.extend(cell.text for cell in row.cells)
    return "\n".join(out)

def token_set(s):
    return set(re.findall(r"[a-zA-Z][a-zA-Z0-9+#.-]{1,}", s.lower()))

def keyword_score(resume, jd):
    jd_terms = token_set(jd)
    if not jd_terms:
        return 0, []
    matched = sorted(token_set(resume) & jd_terms)
    return round(100 * len(matched) / len(jd_terms)), matched

def call_gemini(api_key, resume_text, jd_text, model):
    prompt = f"""
You are a careful resume editor. Tailor the candidate's existing resume to the job description.
STRICT RULES:
- Use ONLY facts explicitly present in the source resume.
- Never invent or infer employers, dates, titles, degrees, certifications, tools, metrics, achievements or responsibilities.
- Do not claim a skill just because it appears in the job description.
- Keep the same meaning and level of seniority.
- Return JSON only, with keys: summary, skills, bullet_rewrites.
- summary: revised summary string, or empty string if there is no source summary.
- skills: array of strings. Only include skills already evidenced in the resume. Do not add new skills.
- bullet_rewrites: array of objects with exact_source and revised. exact_source must be an exact full existing resume bullet/line; revised must preserve its facts and meaning.
- Rewrite only the most relevant existing bullets, maximum 5.
- If a safe rewrite is not possible, leave it out.
JOB DESCRIPTION:
{jd_text[:18000]}
SOURCE RESUME:
{resume_text[:24000]}
"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    payload = {"contents":[{"parts":[{"text":prompt}]}],
               "generationConfig":{"temperature":0.1,"responseMimeType":"application/json"}}
    response = requests.post(url, json=payload, timeout=90)
    response.raise_for_status()
    raw = response.json()["candidates"][0]["content"]["parts"][0]["text"]
    # tolerate fenced JSON
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip(), flags=re.I)
    return json.loads(raw)

def replace_paragraph_text_keep_first_run_style(paragraph, new_text):
    # Retain paragraph properties and first run formatting; clear remaining runs.
    runs = paragraph.runs
    if runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)

def all_docx_paragraphs(doc):
    # Include top-level paragraphs and table-cell paragraphs.
    found = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                found.extend(cell.paragraphs)
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            found.extend(ncell.paragraphs)
    return found

def apply_docx_edits(data, edits):
    doc = Document(io.BytesIO(data))
    paragraphs = all_docx_paragraphs(doc)
    text_to_paragraph = {}
    for p in paragraphs:
        text_to_paragraph.setdefault(p.text.strip(), []).append(p)
    applied, skipped = [], []
    summary = (edits.get("summary") or "").strip()
    if summary:
        headings = {"summary","professional summary","profile","career objective","professional profile"}
        target = None
        for i,p in enumerate(paragraphs):
            if p.text.strip().lower().rstrip(":") in headings:
                for q in paragraphs[i+1:i+4]:
                    if q.text.strip():
                        target=q
                        break
                if target: break
        if target:
            replace_paragraph_text_keep_first_run_style(target, summary)
            applied.append("Professional summary")
        else:
            skipped.append("Summary: no clearly labelled summary section found")
    for item in edits.get("bullet_rewrites", [])[:5]:
        source = (item.get("exact_source") or "").strip()
        revised = (item.get("revised") or "").strip()
        matches = text_to_paragraph.get(source, [])
        if source and revised and matches:
            replace_paragraph_text_keep_first_run_style(matches[0], revised)
            applied.append("Experience bullet")
        elif source:
            skipped.append("A bullet was not found as an exact paragraph and was not changed")
    output=io.BytesIO()
    doc.save(output)
    return output.getvalue(), applied, skipped

def pdf_embed(data):
    b64=base64.b64encode(data).decode("ascii")
    return f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="760" type="application/pdf"></iframe>'

with st.sidebar:
    st.header("1. Master resume")
    resume_file=st.file_uploader("Upload PDF or DOCX", type=["pdf","docx"])
    st.header("2. Job description")
    jd_file=st.file_uploader("Optional TXT job description", type=["txt"])
    jd_text=st.text_area("Paste job description", height=230)
    if jd_file:
        jd_text = (jd_text + "\n" + jd_file.getvalue().decode("utf-8", errors="ignore")).strip()
    st.header("3. AI configuration")
    api_key=st.text_input("Gemini API key", type="password", help="Used only for this session. Do not commit API keys to GitHub.")
    model=st.text_input("Gemini model", value="gemini-2.5-flash")

if resume_file:
    source=resume_file.getvalue()
    ext=Path(resume_file.name).suffix.lower()
    try:
        resume_text=extract_pdf(source) if ext==".pdf" else extract_docx(source)
    except Exception as e:
        st.error(f"Could not read this resume: {e}")
        st.stop()

    c1,c2=st.columns(2)
    with c1:
        st.subheader("Master resume")
        st.write(resume_file.name)
        st.caption(f"{len(source):,} bytes")
    with c2:
        st.subheader("Job match indicator")
        if jd_text.strip():
            score,matched=keyword_score(resume_text,jd_text)
            st.metric("Keyword overlap", f"{score}%")
            st.caption("This is a simple keyword-overlap indicator, not an ATS score or hiring probability.")
            st.write("Shared terms:", ", ".join(matched[:35]) if matched else "None detected")
        else:
            st.info("Paste a job description to calculate keyword overlap.")

    st.divider()
    if st.button("Analyze and tailor resume", type="primary", disabled=not (jd_text.strip() and api_key.strip())):
        try:
            with st.spinner("Analyzing the job description and checking claims against your resume..."):
                edits=call_gemini(api_key.strip(),resume_text,jd_text,model.strip())
                st.session_state["rp_edits"]=edits
                st.session_state["rp_source_name"]=resume_file.name
                if ext==".docx":
                    output,applied,skipped=apply_docx_edits(source,edits)
                    st.session_state["rp_output"]=output
                    st.session_state["rp_applied"]=applied
                    st.session_state["rp_skipped"]=skipped
                    st.session_state["rp_output_ext"]="docx"
                else:
                    st.session_state["rp_output"]=source
                    st.session_state["rp_applied"]=[]
                    st.session_state["rp_skipped"]=["PDF tailoring is not enabled in this beta. The original PDF is retained unchanged to protect its design."]
                    st.session_state["rp_output_ext"]="pdf"
            st.success("Analysis completed. Review the proposed edits below.")
        except requests.HTTPError as e:
            st.error(f"Gemini API error: {e.response.status_code} – {e.response.text[:1000]}")
        except Exception as e:
            st.error(f"Could not complete tailoring: {e}")

    edits=st.session_state.get("rp_edits")
    if edits and st.session_state.get("rp_source_name")==resume_file.name:
        st.subheader("Proposed tailored content")
        if edits.get("summary"):
            st.markdown("**Revised professional summary**")
            st.write(edits["summary"])
        if edits.get("skills"):
            st.markdown("**Relevant existing skills**")
            st.write(", ".join(edits["skills"]))
        if edits.get("bullet_rewrites"):
            st.markdown("**Proposed experience bullet changes**")
            for item in edits["bullet_rewrites"]:
                with st.expander(item.get("exact_source","Original bullet")):
                    st.write("**Suggested wording**")
                    st.write(item.get("revised",""))
        if st.session_state.get("rp_applied"):
            st.success("Applied to DOCX: " + ", ".join(st.session_state["rp_applied"]))
        for warning in st.session_state.get("rp_skipped",[]):
            st.warning(warning)

        st.subheader("Modified resume preview")
        if ext==".pdf":
            st.components.v1.html(pdf_embed(source), height=780, scrolling=True)
            st.caption("This is the unchanged original PDF preview. Tailored PDF output is not yet enabled.")
            st.download_button("Download original PDF", source, file_name="ResumePilot_original.pdf", mime="application/pdf")
        else:
            preview_doc=Document(io.BytesIO(st.session_state["rp_output"]))
            preview_text="\n\n".join(p.text for p in preview_doc.paragraphs)
            st.text_area("DOCX text preview", preview_text, height=380)
            st.caption("This is a text preview. Download and open the DOCX in Word to inspect page-level formatting.")
            st.download_button("Download tailored DOCX", st.session_state["rp_output"],
                file_name="ResumePilot_tailored.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.markdown("Upload a PDF or DOCX resume in the sidebar to begin.")
    st.markdown("""
    **Workflow**
    1. Upload a master resume.
    2. Paste the job description.
    3. Add your Gemini API key in the sidebar.
    4. Analyze and review proposed changes.
    5. Download the tailored DOCX.
    """)

st.divider()
st.caption("Beta limitations: DOCX formatting is retained at paragraph/style level, but complex text runs, headers/footers, floating objects and page breaks require validation. PDF tailoring and true page-rendered DOCX preview are not included yet.")
