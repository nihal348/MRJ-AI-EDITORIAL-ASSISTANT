import io, os, json, requests, pdfplumber
from docx import Document
import streamlit as st

def extract_text(uploaded):
    data = uploaded.getvalue()
    if uploaded.name.lower().endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return "\n".join([p.extract_text() or "" for p in pdf.pages])

def run_ai_analysis(text):
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
    if not api_key:
        return {"status": "ERROR", "message": "API Key missing. Please set GROQ_API_KEY in Secrets."}
    
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": "You are a journal editor assistant. Evaluate structural compliance, methodologies, and ethics disclosures."},
            {"role": "user", "content": f"Analyze manuscript for structure and compliance: {text[:15000]}"}
        ],
        "temperature": 0.2
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=30)
        r.raise_for_status()
        return {"status": "OK", "result": r.json()["choices"][0]["message"]["content"]}
    except Exception as e:
        return {"status": "ERROR", "message": f"API Error: {str(e)}"}

def generate_blind_copy(text):
    doc = Document()
    doc.add_heading('Anonymized Manuscript (Blind Reviewer Copy)', 0)
    for paragraph in text.split('\n'):
        if paragraph.strip():
            doc.add_paragraph(paragraph)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# App UI
st.set_page_config(page_title="Journal AI Pre-Screening", layout="wide")
st.title("Journal AI Editorial Pre-Screening")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file:
    text_content = extract_text(uploaded_file)
    
    if st.button("Run Pre-Screening Pipeline", type="primary"):
        with st.spinner("Analyzing manuscript..."):
            res = run_ai_analysis(text_content)
            
            st.subheader("1. AI Analysis & Compliance Findings")
            if res["status"] == "OK":
                st.success("Analysis Complete!")
                st.write(res["result"])
            else:
                st.error(res["message"])

        st.subheader("2. Anonymized Blind Reviewer Copy")
        blind_docx = generate_blind_copy(text_content)
        st.download_button(
            label="Download Blind Reviewer Copy (.docx)",
            data=blind_docx,
            file_name="Blind_Reviewer_Copy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
