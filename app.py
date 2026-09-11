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

def sanitize_text_for_blind_review(text, client=None):
    # Strip emails and ORCID identifiers
    text = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '[AUTHOR EMAIL REDACTED]', text)
    text = re.sub(r'https?://orcid\.org/\d{4}-\d{4}-\d{4}-\d{3}[\dX]', '[ORCID REDACTED]', text)
    
    clean_lines = []
    for line in text.split('\n'):
        line_lower = line.strip().lower()
        if any(kw in line_lower for kw in ["university", "department of", "faculty of", "correspondence to:", "affiliated with", "institute of"]):
            clean_lines.append("[AFFILIATION REDACTED]")
        else:
            clean_lines.append(line)
            
    scrubbed = "\n".join(clean_lines)

    # Use LLM to scrub human author names from the header if client is available
    if client:
        try:
            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": "You are a text anonymization tool. Replace all author names and co-author names with '[AUTHOR NAME REDACTED]'. Do not alter abstract or paper contents. Return ONLY sanitized text."},
                    {"role": "user", "content": scrubbed[:3000]}
                ],
                model="openai/gpt-oss-120b",
                temperature=0.0
            )
            return response.choices[0].message.content + "\n" + scrubbed[3000:]
        except Exception:
            return scrubbed
    return scrubbed

def run_ai_analysis(text, client):
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are a professional academic journal editor assistant. Evaluate manuscript structure, methodology, and compliance disclosures."},
                {"role": "user", "content": f"Analyze this manuscript text:\n\n{text[:15000]}"}
            ],
            model="openai/gpt-oss-120b",
            temperature=0.2
        )
        return {"status": "OK", "result": response.choices[0].message.content}
    except Exception as e:
        return {"status": "ERROR", "message": f"Groq Client Error: {str(e)}"}

def generate_report_docx(ai_result, raw_text):
    doc = Document()
    doc.add_heading('Editorial AI Pre-Screening Summary Report', 0)
    
    doc.add_heading('1. AI Analysis & Compliance Findings', level=1)
    doc.add_paragraph(ai_result)
    
    doc.add_heading('2. Original Manuscript Preview (First 5,000 Characters)', level=1)
    doc.add_paragraph(raw_text[:5000] + ("..." if len(raw_text) > 5000 else ""))
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def generate_blind_copy_docx(sanitized_text):
    doc = Document()
    doc.add_heading('Anonymized Manuscript (Blind Reviewer Copy)', 0)
    
    for line in sanitized_text.split('\n'):
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
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
    
    if not api_key:
        st.error("API Key missing. Please set GROQ_API_KEY in Secrets.")
    else:
        client = Groq(api_key=api_key)

        if st.button("Run Pre-Screening Pipeline", type="primary"):
            with st.spinner("Analyzing manuscript via Groq AI..."):
                res = run_ai_analysis(raw_text, client)
                
                st.subheader("1. AI Analysis & Compliance Findings")
                if res["status"] == "OK":
                    st.success("Analysis Complete!")
                    st.write(res["result"])
                    
                    # Full Pre-Screening Report Download
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
