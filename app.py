import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber
import requests
import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph
from groq import Groq, BadRequestError


# ============================================================
# CONFIGURATION & CONSTANTS
# ============================================================

PRIMARY_MODEL = "openai/gpt-oss-120b"
FALLBACK_MODEL = "llama-3.3-70b-versatile"

OPENALEX_URL = "https://api.openalex.org/works"
OPENALEX_AUTHOR_URL = "https://api.openalex.org/authors"

REQUIRED_TEMPLATE_SECTIONS = [
    "Abstract",
    "Introduction",
    "Materials and Methods",
    "Results and Discussion",
    "Conclusions",
    "Multidisciplinary Domains",
    "Funding",
    "Acknowledgments",
    "Conflicts of Interest",
    "AI Usage",
    "References",
    "Ethics Statement",
]

SECTION_PATTERNS = {
    "Abstract": r"(?i)^\s*(?:abstract|executive\s+summary)\s*:?$",
    "Introduction": r"(?i)^\s*(?:(?:section\s+)?1(?:\.0?)?|[ivx]+\.?)?\s*(?:introduction|background)\s*:?$",
    "Materials and Methods": r"(?i)^\s*(?:(?:section\s+)?2(?:\.0?)?|[ivx]+\.?)?\s*(?:materials\s+and\s+methods|materials\s+&\s+methods|methodology|methods|experimental\s+procedures)\s*:?$",
    "Results and Discussion": r"(?i)^\s*(?:(?:section\s+)?3(?:\.0?)?|[ivx]+\.?)?\s*(?:results\s+and\s+discussion|results\s+&\s+discussion|results|findings\s+and\s+discussion)\s*:?$",
    "Conclusions": r"(?i)^\s*(?:(?:section\s+)?4(?:\.0?)?|[ivx]+\.?)?\s*(?:conclusions?|summary\s+and\s+conclusions?|concluding\s+remarks)\s*:?$",
    "Multidisciplinary Domains": r"(?i)^\s*(?:multidisciplinary\s+domains?|research\s+domains?)\s*:?$",
    "Funding": r"(?i)^\s*(?:funding(?:\s+information)?|financial\s+support|grant\s+support)\s*:?$",
    "Acknowledgments": r"(?i)^\s*(?:acknowledgments?|acknowledgements?)\s*:?$",
    "Conflicts of Interest": r"(?i)^\s*(?:conflicts?\s+of\s+interest|competing\s+interests?|disclosure\s+statement)\s*:?$",
    "AI Usage": r"(?i)^\s*(?:declaration\s+on\s+ai(?:\s+usage)?|generative\s+ai\s+statement|ai\s+usage|declaration\s+on\s+artificial\s+intelligence)\s*:?$",
    "References": r"(?i)^\s*(?:references?|bibliography|literature\s+cited)\s*:?$",
    "Ethics Statement": r"(?i)^\s*(?:ethics\s+statement|ethical\s+approval|ethics\s+approval|institutional\s+review\s+board|irb\s+statement)\s*:?$",
}

NORTHEAST_STATES = {
    "assam", "arunachal pradesh", "manipur", "meghalaya", "mizoram", "nagaland", "sikkim", "tripura"
}

NORTHEAST_INSTITUTION_TERMS = [
    "iit guwahati", "tezu university", "tezpur university", "nit silchar", "assam university",
    "gauhati university", "cotton university", "dibrugarh university", "nehu",
    "north-eastern hill university", "niser", "nit agartala", "manipur university",
    "mizoram university", "nagaland university", "tripura university", "rajiv gandhi university",
    "sikkim university"
]

ASSAM_TERMS = [
    "assam", "iit guwahati", "gauhati university", "cotton university", "dibrugarh university",
    "assam university", "tezpur university", "nit silchar", "indian institute of technology guwahati"
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_secret(name: str, default: str = "") -> str:
    try:
        val = st.secrets.get(name)
        if val:
            return str(val)
    except Exception:
        pass
    return os.getenv(name, default)


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def word_count(text: str) -> int:
    return len(re.findall(r"\b[\w'-]+\b", text or ""))


def clean_json_response(raw_resp: str) -> Dict[str, Any]:
    """Strip markdown wrapping and parse strictly valid JSON."""
    clean = re.sub(r"^```(?:json)?\s*", "", raw_resp.strip(), flags=re.MULTILINE)
    clean = re.sub(r"```\s*$", "", clean.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(clean)
    except Exception:
        match = re.search(r"(\{.*\})", clean, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        raise ValueError("Could not parse valid JSON from AI response.")


# ============================================================
# ASSET & SECTION EXTRACTION (RULE 4: DEDUPLICATION & FORMAL CAPTIONS)
# ============================================================

def extract_text_and_assets(uploaded_file) -> Tuple[str, Dict[str, Any]]:
    """
    Extract document text and deduplicate formal captions.
    Ignores informal narrative mentions (e.g. 'as shown in Figure 1').
    """
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    asset_meta = {
        "file_type": "docx" if name.endswith(".docx") else "pdf",
        "total_images": 0,
        "total_tables": 0,
        "detected_captions": [],
    }

    seen_labels = set()
    caption_regex = re.compile(r"^\s*((?:Figure|Fig\.?|Table)\s*\d+)[\s.:-]+([^\n]+)", re.IGNORECASE)

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = []
        asset_meta["total_tables"] = len(doc.tables)
        xml_text = doc._element.xml
        asset_meta["total_images"] = len(re.findall(r"<a:blip|<w:drawing", xml_text))

        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                parts.append(txt)
                m = caption_regex.match(txt)
                if m:
                    label = normalize(m.group(1)).title()
                    if label not in seen_labels:
                        seen_labels.add(label)
                        asset_meta["detected_captions"].append({
                            "label": label,
                            "caption": txt,
                            "placement": "In-line document body",
                            "has_image": asset_meta["total_images"] > 0,
                            "has_table": asset_meta["total_tables"] > 0,
                        })

        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))

        return "\n".join(parts), asset_meta

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                page_text = page.extract_text() or ""
                pages.append(page_text)

                images_on_page = len(page.images)
                tables_on_page = len(page.extract_tables() or [])
                asset_meta["total_images"] += images_on_page
                asset_meta["total_tables"] += tables_on_page

                for line in page_text.splitlines():
                    sline = line.strip()
                    m = caption_regex.match(sline)
                    if m:
                        label = normalize(m.group(1)).title()
                        if label not in seen_labels:
                            seen_labels.add(label)
                            asset_meta["detected_captions"].append({
                                "label": label,
                                "caption": sline,
                                "placement": "In-line document body",
                                "has_image": images_on_page > 0,
                                "has_table": tables_on_page > 0,
                            })

        return "\n".join(pages), asset_meta

    raise ValueError("Unsupported format. Please upload a .docx or .pdf file.")


# ============================================================
# AUDIT JSON SCHEMA SPECIFICATION
# ============================================================

AUDIT_STRICT_SCHEMA = {
    "name": "manuscript_pre_screening_audit",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "manuscript_title": {"type": "string"},
            "editorial_verdict": {
                "type": "object",
                "properties": {
                    "decision": {
                        "type": "string",
                        "enum": ["Accept with Minor Revisions", "Major Revisions", "Reject"],
                    },
                    "summary_notes": {"type": "string"},
                },
                "required": ["decision", "summary_notes"],
                "additionalProperties": False,
            },
            "structural_section_checks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "section_name": {"type": "string"},
                        "detected_heading": {"type": "string"},
                        "first_line_quote": {"type": "string"},
                        "status": {
                            "type": "string",
                            "enum": ["PASS", "WARN", "FAIL", "NOT EVALUATED (EXCERPT PROVIDED)"],
                        },
                    },
                    "required": ["section_name", "detected_heading", "first_line_quote", "status"],
                    "additionalProperties": False,
                },
            },
            "visual_asset_audit": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "placement": {"type": "string"},
                        "visual_present": {"type": "boolean"},
                        "status": {
                            "type": "string",
                            "enum": ["PASS", "FAIL", "NOT EVALUATED (EXCERPT PROVIDED)"],
                        },
                    },
                    "required": ["label", "placement", "visual_present", "status"],
                    "additionalProperties": False,
                },
            },
            "methodology_and_math_logic": {
                "type": "object",
                "properties": {
                    "sample_size_check": {
                        "type": "string",
                        "enum": ["PASS", "WARN", "FAIL"],
                    },
                    "missing_domain_metrics": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "software_reproducibility_notes": {"type": "string"},
                },
                "required": ["sample_size_check", "missing_domain_metrics", "software_reproducibility_notes"],
                "additionalProperties": False,
            },
        },
        "required": [
            "manuscript_title",
            "editorial_verdict",
            "structural_section_checks",
            "visual_asset_audit",
            "methodology_and_math_logic",
        ],
        "additionalProperties": False,
    },
}


# ============================================================
# AUDIT ENGINE (EXACT INSTRUCTIONS APPLIED)
# ============================================================

def run_editorial_audit(raw_text: str, asset_meta: Dict[str, Any], client: Groq) -> Dict[str, Any]:
    """Execute evidence-based manuscript audit enforcing all 5 critical audit rules."""
    total_words = word_count(raw_text)
    has_references = bool(re.search(r"(?i)\b(?:references|bibliography)\b", raw_text))
    is_likely_excerpt = (total_words < 2500) and not has_references

    truncation_context = (
        "DOCUMENT COMPLETION STATUS: FULL PAPER LIKELY."
        if not is_likely_excerpt
        else "DOCUMENT COMPLETION STATUS: FRAGMENT / EXCERPT DETECTED (Document length is brief and lacks References)."
    )

    prompt = f"""You are an expert academic manuscript pre-screener and editorial auditor. Conduct a thorough, evidence-based audit of the provided manuscript text.

CRITICAL AUDIT CONSTRAINTS & RULES:

1. EXCERPT & TRUNCATION HANDLING (STRICT):
   - NEVER declare a section "FAIL" or "Not Found" if you are evaluating an incomplete excerpt or fragment of a document.
   - If a standard section is absent due to text truncation or partial file upload, mark its status strictly as "NOT EVALUATED (EXCERPT PROVIDED)".
   - Mark a section as "FAIL" ONLY if the manuscript is complete and the section is explicitly missing.
   - For optional sections (e.g., Acknowledgments, Funding, Ethics Statement, AI Usage), mark as "WARN" rather than "FAIL" if omitted in full papers.

2. MANUSCRIPT TITLE EXTRACTION:
   - Extract the actual academic paper title (e.g., "Consumer Perception, Food Waste and Food Packaging Research...").
   - NEVER output a DOI link, URL, header string, or journal name in the "manuscript_title" field.

3. EVIDENCE-BASED SECTION VERIFICATION:
   - For every detected section, extract the exact first-line quote from the document text to prove its presence.

4. VISUAL ASSET AUDIT & DEDUPLICATION:
   - Audit formal captions only (e.g., "Figure 1: ...", "Table 2: ...").
   - Do NOT create separate entries for informal text mentions within narrative paragraphs (e.g., ignore sentences like "as shown in Figure 1").

5. METHODOLOGICAL & SCIENTOMETRIC TRANSPARENCY:
   - For bibliometric/scientometric papers, explicitly check for missing standard metrics: h-index, g-index, m-index.
   - Audit software reproducibility: check if version numbers, parameter settings, or normalization techniques (e.g., VOSviewer, Biblioshiny, CiteSpace) are explicitly reported.

{truncation_context}

DETECTED ASSET CAPTIONS IN EXTRACTION:
{json.dumps(asset_meta['detected_captions'], indent=2)}

SECTIONS TO EVALUATE:
{json.dumps(REQUIRED_TEMPLATE_SECTIONS)}

MANUSCRIPT TEXT:
--- START OF TEXT ---
{raw_text[:14000]}
--- END OF TEXT ---

Return strictly valid, unformatted JSON matching the schema."""

    messages = [
        {
            "role": "system",
            "content": "You are a senior academic peer reviewer and editorial auditor. Output strictly valid JSON matching the requested schema.",
        },
        {"role": "user", "content": prompt},
    ]

    # Attempt 1: Strict JSON schema execution with primary model
    try:
        resp = client.chat.completions.create(
            model=PRIMARY_MODEL,
            messages=messages,
            temperature=0.0,
            reasoning_effort="low",
            include_reasoning=False,
            max_completion_tokens=4500,
            response_format={"type": "json_schema", "json_schema": AUDIT_STRICT_SCHEMA},
        )
        return clean_json_response(resp.choices[0].message.content or "{}")
    except (BadRequestError, Exception):
        pass

    # Attempt 2: Primary model with json_object enforcement
    try:
        resp = client.chat.completions.create(
            model=PRIMARY_MODEL,
            messages=messages,
            temperature=0.0,
            reasoning_effort="low",
            include_reasoning=False,
            max_completion_tokens=4500,
            response_format={"type": "json_object"},
        )
        return clean_json_response(resp.choices[0].message.content or "{}")
    except (BadRequestError, Exception):
        pass

    # Attempt 3: High-reliability fallback model
    resp = client.chat.completions.create(
        model=FALLBACK_MODEL,
        messages=messages,
        temperature=0.0,
        max_completion_tokens=4000,
        response_format={"type": "json_object"},
    )
    return clean_json_response(resp.choices[0].message.content or "{}")


# ============================================================
# REVIEWER DISCOVERY (OPENALEX)
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    res = requests.get(url, params=params, timeout=30)
    res.raise_for_status()
    return res.json()


def candidate_region_match(candidate: Dict[str, Any], region: str) -> bool:
    text = " ".join([
        candidate.get("institution", ""),
        candidate.get("city", ""),
        candidate.get("region", ""),
        candidate.get("country", ""),
    ]).lower()

    if region == "India":
        return candidate.get("country", "").lower() == "in" or "india" in text
    if region == "Northeast India":
        return any(term in text for term in (set(NORTHEAST_STATES) | set(NORTHEAST_INSTITUTION_TERMS)))
    if region == "Assam":
        return any(term in text for term in ASSAM_TERMS)
    return False


def search_openalex_reviewers(query_terms: List[str], region: str, mailto: str = "", max_candidates: int = 5) -> List[Dict[str, Any]]:
    query = " ".join(query_terms[:4]).strip()
    if not query:
        return []

    data = openalex_get(
        OPENALEX_URL,
        {"search": query, "per-page": 30, "sort": "publication_year:desc"},
        mailto=mailto,
    )

    candidates: Dict[str, Dict[str, Any]] = {}
    for work in data.get("results", []):
        year = work.get("publication_year") or 0
        title = work.get("display_name") or ""
        doi = work.get("doi") or ""

        for authorship in work.get("authorships", []):
            author = authorship.get("author") or {}
            author_id = author.get("id")
            author_name = author.get("display_name")
            if not author_id or not author_name:
                continue

            institutions = authorship.get("institutions") or []
            for inst in institutions:
                candidate = {
                    "author_id": author_id,
                    "name": author_name,
                    "institution": inst.get("display_name") or "Unspecified Institution",
                    "country": inst.get("country_code") or "",
                    "city": (inst.get("geo") or {}).get("city") or "",
                    "region": (inst.get("geo") or {}).get("region") or "",
                    "recent_pubs": [],
                }
                if not candidate_region_match(candidate, region):
                    continue

                if author_id not in candidates:
                    candidates[author_id] = candidate
                candidates[author_id]["recent_pubs"].append({"year": year, "title": title, "doi": doi})

    results = list(candidates.values())
    results.sort(key=lambda x: len(x["recent_pubs"]), reverse=True)
    return results[:max_candidates]


def reviewer_discovery_report(title: str, mailto: str = "") -> Dict[str, List[Dict[str, Any]]]:
    # Extract salient search terms from the identified manuscript title
    stopwords = {"a", "an", "the", "and", "or", "in", "on", "at", "to", "for", "with", "of", "by", "from", "using", "study", "analysis"}
    terms = [w for w in re.findall(r"\b[A-Za-z]{3,}\b", title) if w.lower() not in stopwords]
    output = {}
    for r in ["India", "Northeast India", "Assam"]:
        try:
            output[r] = search_openalex_reviewers(terms, r, mailto=mailto)
        except Exception as exc:
            output[r] = [{"error": f"OpenAlex query failed: {exc}"}]
    return output


# ============================================================
# ANONYMIZED BLIND COPY BUILDER
# ============================================================

def blind_copy_docx(original_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(original_bytes))
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""

    paras = list(doc.paragraphs)
    abstract_idx = None
    for i, p in enumerate(paras):
        if re.match(r"(?i)^\s*abstract\s*:?", p.text.strip()):
            abstract_idx = i
            break

    if abstract_idx is not None and abstract_idx > 1:
        first_kept = False
        for i in range(abstract_idx):
            txt = paras[i].text.strip()
            if not txt:
                continue
            if not first_kept:
                first_kept = True
                continue
            p_elem = paras[i]._element
            p_elem.getparent().remove(p_elem)

    for p in doc.paragraphs:
        for r in p.runs:
            r.text = re.sub(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", "[REDACTED EMAIL]", r.text, flags=re.I)
            r.text = re.sub(r"\b(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}\b", "[REDACTED ORCID]", r.text, flags=re.I)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# REPORT BUILDER (.DOCX)
# ============================================================

def generate_docx_report(audit: Dict[str, Any], reviewers: Dict[str, List[Dict[str, Any]]]) -> bytes:
    doc = Document()
    doc.add_heading("Manuscript Pre-Screening & Editorial Audit Report", 0)
    doc.add_paragraph(f"Manuscript Title: {audit.get('manuscript_title', 'Not specified')}")
    
    verd = audit.get("editorial_verdict", {})
    doc.add_paragraph(f"Verdict: {verd.get('decision', 'Under Review')}")
    doc.add_paragraph(f"Summary Notes: {verd.get('summary_notes', '')}")

    doc.add_heading("1. Structural Section Verification", level=1)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    h = tbl.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Section", "Detected Heading", "First-Line Quote", "Status"
    for s in audit.get("structural_section_checks", []):
        row = tbl.add_row().cells
        row[0].text = s.get("section_name", "")
        row[1].text = s.get("detected_heading", "")
        row[2].text = s.get("first_line_quote", "")
        row[3].text = s.get("status", "")

    doc.add_heading("2. Visual Asset Audit", level=1)
    v_tbl = doc.add_table(rows=1, cols=4)
    v_tbl.style = "Table Grid"
    vh = v_tbl.rows[0].cells
    vh[0].text, vh[1].text, vh[2].text, vh[3].text = "Label", "Placement", "Visual Present", "Status"
    for v in audit.get("visual_asset_audit", []):
        row = v_tbl.add_row().cells
        row[0].text = v.get("label", "")
        row[1].text = v.get("placement", "")
        row[2].text = "Yes" if v.get("visual_present") else "No"
        row[3].text = v.get("status", "")

    doc.add_heading("3. Methodology & Math Logic", level=1)
    meth = audit.get("methodology_and_math_logic", {})
    doc.add_paragraph(f"Sample Size Accounting: {meth.get('sample_size_check', 'N/A')}")
    missing_metrics = ", ".join(meth.get("missing_domain_metrics", [])) or "None identified"
    doc.add_paragraph(f"Missing Domain Metrics: {missing_metrics}")
    doc.add_paragraph(f"Software Reproducibility Notes: {meth.get('software_reproducibility_notes', '')}")

    doc.add_heading("4. Potential Reviewer Candidates (OpenAlex)", level=1)
    for reg, cands in reviewers.items():
        doc.add_heading(reg, level=2)
        for c in cands:
            if "name" in c:
                doc.add_paragraph(f"- {c['name']} ({c.get('institution', 'N/A')}): {len(c.get('recent_pubs', []))} verified publications")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI (CLEAN EDITORIAL DASHBOARD, NO RAW JSON)
# ============================================================

st.set_page_config(page_title="Academic Manuscript Audit", page_icon="📑", layout="wide")
st.title("📑 Academic Manuscript Pre-Screening & Editorial Auditor")
st.caption("Evidence-based structural checks, caption deduplication, and scientometric reproducibility analysis.")

with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key = st.text_input("Groq API Key", value=get_secret("GROQ_API_KEY"), type="password")
    openalex_mailto = st.text_input("OpenAlex Mailto Email", value=get_secret("OPENALEX_MAILTO", "editor@academicprescreen.org"))
    st.markdown("---")
    st.markdown("**Critical Pre-Screening Directives:**")
    st.markdown("1. **Truncation Handling**: Missing sections in excerpts marked as `NOT EVALUATED (EXCERPT PROVIDED)`.")
    st.markdown("2. **Title Extraction**: Real paper title extracted; no DOIs/URLs.")
    st.markdown("3. **Evidence-Based Quotes**: Verifies first-line quotes for detected headings.")
    st.markdown("4. **Asset Deduplication**: Audits formal captions only.")
    st.markdown("5. **Scientometrics**: Checks h/g/m indices & software parameters.")

uploaded_file = st.file_uploader("Upload Manuscript (.pdf or .docx)", type=["pdf", "docx"])

if uploaded_file and st.button("🚀 Conduct Evidence-Based Audit", type="primary"):
    if not groq_api_key:
        st.error("Please provide a valid Groq API Key.")
        st.stop()

    try:
        with st.spinner("Extracting document text and deduplicating formal captions..."):
            raw_text, asset_meta = extract_text_and_assets(uploaded_file)
            if not raw_text.strip():
                st.error("Could not extract readable text from the uploaded document.")
                st.stop()

        with st.spinner("Executing rigorous pre-screening audit with Groq AI..."):
            client = Groq(api_key=groq_api_key)
            audit_result = run_editorial_audit(raw_text, asset_meta, client)

        with st.spinner("Querying OpenAlex for verified regional reviewers..."):
            detected_title = audit_result.get("manuscript_title", "")
            reviewers = reviewer_discovery_report(detected_title, mailto=openalex_mailto)

        with st.spinner("Compiling Word report and blind reviewer copy..."):
            orig_bytes = uploaded_file.getvalue()
            blind_bytes = blind_copy_docx(orig_bytes) if uploaded_file.name.endswith(".docx") else orig_bytes
            docx_report = generate_docx_report(audit_result, reviewers)

        st.session_state["audit"] = audit_result
        st.session_state["reviewers"] = reviewers
        st.session_state["blind_bytes"] = blind_bytes
        st.session_state["docx_report"] = docx_report
        st.success("Audit complete.")

    except Exception as e:
        st.error(f"Audit processing error: {e}")

# Render results in structured dashboard (NO RAW JSON DUMP)
if "audit" in st.session_state:
    audit = st.session_state["audit"]
    reviewers = st.session_state["reviewers"]
    verd = audit.get("editorial_verdict", {})
    decision = verd.get("decision", "Under Review")

    st.markdown("---")
    c1, c2 = st.columns([1, 3])
    with c1:
        if decision == "Accept with Minor Revisions":
            st.success(f"### Verdict:\n**{decision}**")
        elif decision == "Major Revisions":
            st.warning(f"### Verdict:\n**{decision}**")
        else:
            st.error(f"### Verdict:\n**{decision}**")
    with c2:
        st.subheader(audit.get("manuscript_title", "Untitled Manuscript"))
        st.write(f"**Editorial Summary:** {verd.get('summary_notes', '')}")

    tab1, tab2, tab3, tab4 = st.tabs([
        "🏛️ Structural Section Verification",
        "🖼️ Visual Asset Audit",
        "🔬 Methodology & Math Logic",
        "👥 Verified Reviewer Candidates",
    ])

    # Tab 1: Structural Section Checks
    with tab1:
        st.subheader("Structural Section Checks")
        struct_data = audit.get("structural_section_checks", [])
        rows = []
        for s in struct_data:
            stat = s.get("status", "")
            badge = "🟢 PASS" if stat == "PASS" else ("🟡 WARN" if stat == "WARN" else ("⚪ NOT EVALUATED" if "NOT EVALUATED" in stat else "🔴 FAIL"))
            rows.append({
                "Section": s.get("section_name"),
                "Status": badge,
                "Detected Heading": s.get("detected_heading"),
                "First-Line Quote": s.get("first_line_quote"),
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)

    # Tab 2: Visual Asset Audit
    with tab2:
        st.subheader("Formal Visual Asset Audit")
        v_data = audit.get("visual_asset_audit", [])
        if not v_data:
            st.info("No formal figure or table captions detected.")
        else:
            v_rows = []
            for v in v_data:
                stat = v.get("status", "")
                badge = "🟢 PASS" if stat == "PASS" else ("⚪ NOT EVALUATED" if "NOT EVALUATED" in stat else "🔴 FAIL")
                v_rows.append({
                    "Asset Label": v.get("label"),
                    "Placement": v.get("placement"),
                    "Visual Present": "✅ Yes" if v.get("visual_present") else "❌ No",
                    "Status": badge,
                })
            st.dataframe(v_rows, use_container_width=True, hide_index=True)

    # Tab 3: Methodology & Math Logic
    with tab3:
        st.subheader("Methodology, Scientometrics & Software Reproducibility")
        meth = audit.get("methodology_and_math_logic", {})
        mc1, mc2 = st.columns(2)
        mc1.metric("Sample Size Accounting", meth.get("sample_size_check", "N/A"))
        missing = meth.get("missing_domain_metrics", [])
        mc2.write("**Missing Domain-Specific Metrics (Scientometrics):**")
        if missing:
            for m in missing:
                mc2.markdown(f"- ⚠️ `{m}`")
        else:
            mc2.write("✅ None identified as missing.")
        st.info(f"**Software Reproducibility & Parameter Notes:**\n\n{meth.get('software_reproducibility_notes', 'None recorded.')}")

    # Tab 4: OpenAlex Reviewers
    with tab4:
        st.subheader("OpenAlex Peer Reviewer Discovery")
        st.caption("Matched against candidate profiles in verified regional institutions.")
        for region, cands in reviewers.items():
            st.markdown(f"### Region: {region}")
            if not cands or "error" in cands[0]:
                st.write("No matching profiles found.")
                continue
            r_rows = []
            for c in cands:
                pubs = c.get("recent_pubs", [])
                latest_title = pubs[0]["title"] if pubs else "N/A"
                r_rows.append({
                    "Candidate Name": c.get("name"),
                    "Institution": c.get("institution"),
                    "Verified Works": len(pubs),
                    "Recent Representative Publication": latest_title,
                })
            st.dataframe(r_rows, use_container_width=True, hide_index=True)

    # Editorial Exports
    st.markdown("---")
    st.subheader("📥 Editorial Exports")
    d1, d2, d3 = st.columns(3)
    d1.download_button(
        "📄 Download Editorial Report (.docx)",
        data=st.session_state["docx_report"],
        file_name="Editorial_Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    if uploaded_file.name.endswith(".docx"):
        d2.download_button(
            "🙈 Download Anonymized Blind Copy (.docx)",
            data=st.session_state["blind_bytes"],
            file_name="Anonymized_Reviewer_Copy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    d3.download_button(
        "💾 Download Audit Data (.json)",
        data=json.dumps(audit, indent=2),
        file_name="audit_data.json",
        mime="application/json",
        use_container_width=True,
    )
