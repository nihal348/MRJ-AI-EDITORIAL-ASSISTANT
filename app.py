import io, os, re, pdfplumber, requests
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

    if client:
        try:
            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": "You are a text anonymization tool. Replace all author names and co-author names with '[AUTHOR NAME REDACTED]'. Do not alter abstract or paper contents. Return ONLY sanitized text."},
                    {"role": "user", "content": scrubbed[:3000]}
                ],
                model="openai/gpt-oss-120b",
                temperature=0.0,
                seed=42
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
            temperature=0.0,
            seed=42
        )
        return {"status": "OK", "result": response.choices[0].message.content}
    except Exception as e:
        return {"status": "ERROR", "message": f"Groq Client Error: {str(e)}"}

def recommend_reviewers(text, client):
    prompt = f"""
    Based on the following manuscript text/abstract, suggest potential peer reviewers categorized by geographic affiliation:
    1. India (2 Top/Senior Researchers in India relevant to this field)
    2. Northeast India (2 Leading Researchers from Northeast Indian institutions like IIT Guwahati, Tezpur University, NIT Silchar, NEHU, etc.)
    3. Assam (2 to 3 Eminent Researchers/Professors specifically based in universities or research institutes in Assam)

    For each reviewer, include:
    - Name & Academic Title
    - Institution & Department
    - Primary Area of Expertise / Research Focus matching this manuscript

    Manuscript Text:
    {text[:4000]}
    """
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are an expert editorial board advisor with deep knowledge of Indian academic institutions and subject matter experts."},
                {"role": "user", "content": prompt}
            ],
            model="openai/gpt-oss-120b",
            temperature=0.0,
            seed=42
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Could not retrieve reviewer suggestions: {str(e)}"

def generate_report_docx(ai_result, reviewer_suggestions, raw_text):
    doc = Document()
    doc.add_heading('Editorial AI Pre-Screening Summary Report', 0)
    
    doc.add_heading('1. AI Analysis & Compliance Findings', level=1)
    doc.add_paragraph(ai_result)

    doc.add_heading('2. Recommended Peer Reviewers (Regional & National)', level=1)
    doc.add_paragraph(reviewer_suggestions)
    
    doc.add_heading('3. Original Manuscript Preview (First 5,000 Characters)', level=1)
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

# Session State Initializations
if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None
if "reviewer_suggestions" not in st.session_state:
    st.session_state.reviewer_suggestions = None
if "sanitized_text" not in st.session_state:
    st.session_state.sanitized_text = None
if "raw_text" not in st.session_state:
    st.session_state.raw_text = None

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file:
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
    
    if not api_key:
        st.error("API Key missing. Please set GROQ_API_KEY in Secrets.")
    else:
        client = Groq(api_key=api_key)

        if st.button("Run Pre-Screening Pipeline", type="primary"):
            with st.spinner("Analyzing manuscript, anonymizing text, & matching regional reviewers..."):
                raw_text = extract_text(uploaded_file)
                st.session_state.raw_text = raw_text
                st.session_state.analysis_result = run_ai_analysis(raw_text, client)
                st.session_state.reviewer_suggestions = recommend_reviewers(raw_text, client)
                st.session_state.sanitized_text = sanitize_text_for_blind_review(raw_text, client)

        if st.session_state.analysis_result:
            res = st.session_state.analysis_result
            
            st.subheader("1. AI Analysis & Compliance Findings")
            if res["status"] == "OK":
                st.success("Analysis Complete!")
                st.write(res["result"])
            else:
                st.error(res["message"])

            st.subheader("2. Recommended Potential Reviewers")
            st.markdown(st.session_state.reviewer_suggestions)

            # Combined Report Download
            report_docx = generate_report_docx(
                res["result"] if res["status"] == "OK" else "Analysis failed.",
                st.session_state.reviewer_suggestions,
                st.session_state.raw_text
            )
            st.download_button(
                label="Download Complete Pre-Screening & Reviewer Report (.docx)",
                data=report_docx,
                file_name="Editorial_PreScreening_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            st.subheader("3. Anonymized Blind Reviewer Copy")
            blind_docx = generate_blind_copy_docx(st.session_state.sanitized_text)
            st.download_button(
                label="Download Blind Reviewer Copy (.docx)",
                data=blind_docx,
                file_name="Blind_Reviewer_Copy.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
