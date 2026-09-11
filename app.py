import io
import os
import re
import json
import requests
import pdfplumber
from docx import Document
from docx.shared import Pt
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# 1. DOCUMENT EXTRACTION ENGINE
# ============================================================
def extract_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for table in doc.tables:
        for row in table.rows:
            tables.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(paragraphs + tables)

def extract_pdf(file_bytes):
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    return text

def extract_text(uploaded):
    data = uploaded.getvalue()
    if uploaded.name.lower().endswith(".docx"):
        return extract_docx(data)
    elif uploaded.name.lower().endswith(".pdf"):
        return extract_pdf(data)
    raise ValueError("Only .docx and .pdf files are supported.")

# ============================================================
# 2. ANONYMIZATION ENGINE (BLIND REVIEW COPY)
# ============================================================
def anonymize_text(text):
    out = text
    out = re.sub(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", "[REMOVED EMAIL]", out, flags=re.I)
    out = re.sub(r"\bhttps?://orcid\.org/\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b", "[REMOVED ORCID]", out, flags=re.I)
    
    patterns = [
        r"(?im)^\s*(?:corresponding author|correspondence|email|e-mail)\s*:.*$",
        r"(?im)^\s*(?:authors?|author names?)\s*:.*$",
        r"(?im)^\s*(?:affiliation|affiliations|department|institution|university)\s*:.*$",
    ]
    for p in patterns:
        out = re.sub(p, "[REMOVED FOR BLIND REVIEW]", out)
    return out

def create_blind_docx(text):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    p = doc.add_paragraph()
    p.add_run("BLIND REVIEW COPY\n").bold = True
    p.add_run("Anonymized automatically — Editor must verify before distribution.\n\n")

    for block in text.split("\n"):
        if block.strip():
            doc.add_paragraph(block.strip())

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()

# ============================================================
# 3. OPENALEX LITERATURE SEARCH
# ============================================================
def literature_search(text):
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    query = " ".join(lines[:3])[:250]
    if not query:
        return []
    
    url = "https://api.openalex.org/works"
    params = {"search": query, "per-page": 5, "select": "display_name,publication_year,doi"}
    try:
        r = requests.get(url, params=params, timeout=10)
        return r.json().get("results", [])
    except Exception:
        return []

# ============================================================
# 4. FREE LLM AI ANALYSIS LAYER (Groq API)
# ============================================================
def run_ai_analysis(text, literature):
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY")
    if not api_key:
        return {"status": "NOT_RUN", "message": "API Key missing. Please set GROQ_API_KEY in Secrets."}

    url = "https://api.groq.com/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    
    prompt = f"""
    Analyze this academic manuscript as an editorial pre-screening assistant.
    Check for:
    1. Structural completeness (Abstract, Methods, Results, Funding, Conflicts).
    2. Internal consistency (Sample sizes, causal statements in observational studies).
    3. Ethics statements.

    MANUSCRIPT EXTRACT:
    {text[:25000]}

    Return JSON strictly in this structure:
    {{
        "overall_status": "GREEN/AMBER/RED",
        "findings": [
            {{"category": "category_name", "severity": "BLOCKING/EDITOR_REVIEW/SUGGESTION", "finding": "description"}}
        ]
    }}
    """
    
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": "You are a journal editor assistant. Always return valid JSON only."},
            {"role": "user", "content": prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.1
    }

    try:
        r = requests.post(url, headers=headers, json=payload, timeout=30)
        r.raise_for_status()
        return {"status": "OK", "result": r.json()["choices"][0]["message"]["content"]}
    except Exception as e:
        return {"status": "ERROR", "message": str(e)}

# ============================================================
# 5. USER INTERFACE (Streamlit)
# ============================================================
st.set_page_config(page_title="Journal AI Pre-Screening", layout="wide")
st.title("Journal AI Editorial Pre-Screening & Anonymizer")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file:
    if st.button("Run Pre-Screening Pipeline", type="primary"):
        with st.spinner("Extracting manuscript text..."):
            text = extract_text(uploaded_file)
            st.session_state["text"] = text

        with st.spinner("Searching literature database (OpenAlex)..."):
            st.session_state["lit"] = literature_search(text)

        with st.spinner("Running AI Editorial Check..."):
            st.session_state["ai"] = run_ai_analysis(text, st.session_state["lit"])

if "text" in st.session_state:
    text = st.session_state["text"]

    st.subheader("1. AI Analysis & Compliance Findings")
    ai = st.session_state["ai"]
    if ai.get("status") == "OK":
        res = json.loads(ai["result"])
        st.write(f"**Status:** {res.get('overall_status')}")
        for item in res.get('findings', []):
            st.warning(f"**[{item.get('severity')}]** {item.get('category')}: {item.get('finding')}")
    else:
        st.info(ai.get("message"))

    st.subheader("2. Prior Literature Found")
    for paper in st.session_state.get("lit", []):
        st.write(f"- **{paper.get('display_name')}** ({paper.get('publication_year')}) — {paper.get('doi')}")

    st.divider()
    st.subheader("3. Generate Blind Copy for Peer Reviewers")
    if st.checkbox("I approve generating anonymized reviewer files."):
        blind_text = anonymize_text(text)
        docx_bytes = create_blind_docx(blind_text)
        
        st.download_button(
            label="Download Blind Reviewer Copy (.docx)",
            data=docx_bytes,
            file_name=f"BLINDED_{uploaded_file.name.rsplit('.', 1)[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
