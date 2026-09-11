import io, os, re, pdfplumber
from docx import Document
import streamlit as st
from groq import Groq

def extract_text(uploaded):
    data = uploaded.getvalue()
    if uploaded.name.lower().endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return "\n".join([p.extract_text() or "" for p in pdf.pages])

def sanitize_text_for_blind_review(text):
    # Strip emails
    text = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '[AUTHOR EMAIL REDACTED]', text)
    # Strip ORCID identifiers
    text = re.sub(r'https?://orcid\.org/\d{4}-\d{4}-\d{4}-\d{3}[\dX]', '[ORCID REDACTED]', text)
    
    clean_lines = []
    for line in text.split('\n'):
        line_lower = line.strip().lower()
        # Redact typical author affiliation metadata blocks
        if any(kw in line_lower for kw in ["university", "department of", "faculty of", "correspondence to:", "affiliated with"]):
            clean_lines.append("[AFFILIATION REDACTED]")
        else:
            clean_lines.append(line)
            
    return "\n".join(clean_lines)

def run_ai_analysis(text):
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
    if not api_key:
        return {"status": "ERROR", "message": "API Key missing. Please set GROQ_API_KEY in Secrets."}
    
    try:
        client = Groq(api_key=api_key)
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are a professional academic journal editor assistant. Evaluate manuscript structure, methodology, and compliance disclosures."},
                {"role": "user", "content": f"Analyze this manuscript text:\n\n{text[:15000]}"}
            ],
            model="openai/gpt-oss-120b",  # Active replacement for deprecated Llama models
            temperature=0.2
        )
        return {"status": "OK", "result": response.choices[0].message.content}
    except Exception as e:
        return {"status": "ERROR", "message": f"Groq Client Error: {str(e)}"}

def generate_blind_copy(text):
    sanitized = sanitize_text_for_blind_review(text)
    doc = Document()
    doc.add_heading('Anonymized Manuscript (Blind Reviewer Copy)', 0)
    
    for line in sanitized.split('\n'):
        if line.strip():
            doc.add_paragraph(line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Streamlit App Layout
st.set_page_config(page_title="Journal AI Pre-Screening", layout="wide")
st.title("Journal AI Editorial Pre-Screening")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    
    if st.button("Run Pre-Screening Pipeline", type="primary"):
        with st.spinner("Analyzing manuscript via Groq AI..."):
            res = run_ai_analysis(raw_text)
            
            st.subheader("1. AI Analysis & Compliance Findings")
            if res["status"] == "OK":
                st.success("Analysis Complete!")
                st.write(res["result"])
            else:
                st.error(res["message"])

        # Complete Report Download
                    report_docx = generate_report_docx(res["result"], raw_text)
                    st.download_button(
                        label="Download Complete Pre-Screening Report (.docx)",
                        data=report_docx,
                        file_name="Editorial_PreScreening_Report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error(res["message"])

            st.subheader("2. Anonymized Blind Reviewer Copy")
            sanitized_text = sanitize_text_for_blind_review(raw_text, client)
            blind_docx = generate_blind_copy_docx(sanitized_text)
            st.download_button(
                label="Download Blind Reviewer Copy (.docx)",
                data=blind_docx,
                file_name="Blind_Reviewer_Copy.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
