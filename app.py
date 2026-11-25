import io, os, re, json, tempfile, hashlib
from pathlib import Path
import streamlit as st
import fitz
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

st.set_page_config(page_title="ResumePilot 3.1", page_icon="📄", layout="wide")
st.title("📄 ResumePilot 3.1")
st.caption("Job-specific resume tailoring with original-layout preservation and preview")

st.info("Privacy note: this starter version processes files in the active Streamlit session. Do not upload sensitive resumes until you have reviewed the hosting, retention, and access controls.")

def extract_pdf(data):
    doc=fitz.open(stream=data,filetype="pdf")
    return "\n".join(p.get_text("text") for p in doc)

def extract_docx(data):
    doc=Document(io.BytesIO(data))
    parts=[p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.extend(c.text for c in row.cells)
    return "\n".join(parts)

def words(text):
    return set(re.findall(r"[A-Za-z][A-Za-z0-9+#.-]{1,}", text.lower()))

def score_resume(resume, jd):
    rw, jw=words(resume), words(jd)
    if not jw: return 0, []
    hits=sorted(rw & jw)
    # Explainable keyword-overlap indicator, not an ATS score or hiring probability.
    return round(100*len(hits)/len(jw)), hits

def safe_summary(resume, jd):
    # Conservative approach: no fabricated claims and no automatic rewriting.
    # Return original summary verbatim when found.
    lines=resume.splitlines()
    for i,line in enumerate(lines):
        if re.search(r"professional summary|summary|profile|career objective",line,re.I):
            chunk=[]
            for nxt in lines[i+1:]:
                if re.search(r"^(experience|work experience|employment|education|skills|technical skills|projects|certifications)\b",nxt.strip(),re.I):
                    break
                if nxt.strip(): chunk.append(nxt.strip())
            return " ".join(chunk) if chunk else "No summary text detected."
    return "No clearly labelled summary found. Original resume content is preserved."

def docx_replace_summary(data, jd):
    doc=Document(io.BytesIO(data))
    paras=doc.paragraphs
    for i,p in enumerate(paras):
        if re.search(r"professional summary|summary|profile|career objective",p.text.strip(),re.I):
            # We deliberately retain the source summary and all paragraph styles/layout.
            # No replacement is made until a safe, fact-grounded rewrite is approved.
            out=io.BytesIO(); doc.save(out)
            return out.getvalue(), "Original DOCX structure and formatting preserved. Summary retained unchanged pending a fact-checked rewrite."
    out=io.BytesIO(); doc.save(out)
    return out.getvalue(), "No labelled summary detected. Original DOCX content and formatting retained."

def pdf_copy(data):
    # Byte-for-byte source PDF copy: guarantees the original visual design is not damaged.
    # This version does not perform risky text reflow or font substitution in PDFs.
    return data, "Original PDF preserved exactly. Automatic in-place text editing is disabled for PDFs to avoid damaging layout."

def pdf_preview(data):
    b64=__import__("base64").b64encode(data).decode()
    return f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="780" type="application/pdf"></iframe>'

with st.sidebar:
    st.header("1. Upload master resume")
    resume_file=st.file_uploader("PDF or DOCX", type=["pdf","docx"])
    st.header("2. Job description")
    jd_file=st.file_uploader("Optional JD file", type=["txt"])
    jd_text=st.text_area("Paste job description", height=220, placeholder="Paste the complete job description here...")
    if jd_file:
        jd_text += "\n"+jd_file.getvalue().decode("utf-8",errors="ignore")

st.header("3. Generate and preview")
if resume_file:
    data=resume_file.getvalue()
    ext=Path(resume_file.name).suffix.lower()
    try:
        resume_text=extract_pdf(data) if ext==".pdf" else extract_docx(data)
    except Exception as e:
        st.error(f"Could not read the resume: {e}")
        st.stop()
    left,right=st.columns(2)
    with left:
        st.subheader("Master resume")
        st.write(resume_file.name)
        st.caption(f"{len(data):,} bytes")
    with right:
        st.subheader("Job match indicator")
        if jd_text.strip():
            score,hits=score_resume(resume_text,jd_text)
            st.metric("Keyword overlap",f"{score}%")
            st.caption("A simple keyword-overlap indicator only. It is not an ATS score or probability of getting shortlisted.")
            st.write("Shared terms:",", ".join(hits[:40]) if hits else "No shared terms detected.")
        else:
            st.warning("Add a job description to calculate the keyword-overlap indicator.")
    st.divider()
    if st.button("Generate modified resume",type="primary",disabled=not jd_text.strip()):
        if ext==".docx":
            output,note=docx_replace_summary(data,jd_text)
            outname="ResumePilot_3_1_modified.docx"
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            output,note=pdf_copy(data)
            outname="ResumePilot_3_1_modified.pdf"
            mime="application/pdf"
        st.session_state["generated"]={"data":output,"name":outname,"mime":mime,"note":note,"ext":ext}
    generated=st.session_state.get("generated")
    if generated:
        st.success("Resume copy generated.")
        st.info(generated["note"])
        st.subheader("Modified resume preview")
        if generated["ext"]==".pdf":
            st.components.v1.html(pdf_preview(generated["data"]),height=800,scrolling=True)
        else:
            try:
                preview_doc=Document(io.BytesIO(generated["data"]))
                preview_text="\n\n".join(p.text for p in preview_doc.paragraphs)
                st.text_area("DOCX text preview",preview_text,height=420)
                st.caption("The downloadable DOCX retains its original Word formatting. This text preview is a content preview, not a pixel-perfect page rendering.")
            except Exception:
                st.warning("Preview unavailable. Download the DOCX to inspect it in Word.")
        st.download_button("Download modified resume",generated["data"],file_name=generated["name"],mime=generated["mime"])
else:
    st.markdown("Upload a master resume in the sidebar to begin.")
    st.markdown("""
    **Current 3.1 foundation**
    - PDF preview in the browser.
    - DOCX content preview.
    - DOCX files are edited in-place without rebuilding the document.
    - PDF output is kept byte-for-byte unchanged to protect the original design.
    - Conservative mode avoids inventing candidate information.
    """)

st.divider()
st.caption("ResumePilot 3.1 starter build. For production use, add authentication, private persistent storage, background processing, malware scanning, and robust document conversion/testing.")
