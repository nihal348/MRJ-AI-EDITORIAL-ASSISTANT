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
from groq import Groq, BadRequestError, RateLimitError


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
    """Parse JSON text, stripping potential markdown fences and cleaning invalid control characters."""
    clean = re.sub(r"^```(?:json)?\s*", "", raw_resp.strip(), flags=re.MULTILINE)
    clean = re.sub(r"```\s*$", "", clean.strip(), flags=re.MULTILINE)
    clean = clean.strip()
    try:
        return json.loads(clean)
    except Exception:
        # Match the outermost JSON object if surrounding commentary exists
        match = re.search(r"(\{.*\})", clean, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        raise ValueError("Could not parse valid JSON from AI completion.")


# ============================================================
# ASSET & SECTION EXTRACTION
# ============================================================

def extract_text_and_assets(uploaded_file) -> Tuple[str, Dict[str, Any]]:
    """Extract full manuscript text while auditing visual assets (images and tables)."""
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    asset_meta = {
        "file_type": "docx" if name.endswith(".docx") else "pdf",
        "total_images": 0,
        "total_tables": 0,
        "detected_captions": [],
    }

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
                if re.match(r"(?i)^(figure|fig\.?|table)\s+\d+", txt):
                    asset_meta["detected_captions"].append({
                        "caption": txt,
                        "location": "Inline Body",
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
                    if re.match(r"(?i)^(figure|fig\.?|table)\s+\d+", sline):
                        asset_meta["detected_captions"].append({
                            "caption": sline,
                            "location": f"Page {page_idx + 1}",
                            "has_image": images_on_page > 0,
                            "has_table": tables_on_page > 0,
                        })

        return "\n".join(pages), asset_meta

    raise ValueError("Unsupported format. Please upload a .docx or .pdf file.")


# ============================================================
# AUDIT JSON SCHEMA SPECIFICATION
# ============================================================

AUDIT_STRICT_SCHEMA = {
    "name": "manuscript_peer_review_audit",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "manuscript_meta": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "keywords": {"type": "array", "items": {"type": "string"}},
                    "editorial_verdict": {
                        "type": "string",
                        "enum": ["Accept with Minor Revisions", "Major Revisions", "Reject"],
                    },
                    "verdict_rationale": {"type": "string"},
                },
                "required": ["title", "keywords", "editorial_verdict", "verdict_rationale"],
                "additionalProperties": False,
            },
            "structural_template_audit": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "section_name": {"type": "string"},
                        "detected": {"type": "boolean"},
                        "heading_evidence": {"type": "string"},
                        "status": {"type": "string", "enum": ["PASS", "WARN", "FAIL"]},
                        "critique": {"type": "string"},
                    },
                    "required": ["section_name", "detected", "heading_evidence", "status", "critique"],
                    "additionalProperties": False,
                },
            },
            "visual_asset_audit": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "asset_type": {"type": "string"},
                        "label": {"type": "string"},
                        "caption": {"type": "string"},
                        "placement": {"type": "string"},
                        "visual_present": {"type": "boolean"},
                        "status": {"type": "string", "enum": ["PASS", "FAIL"]},
                    },
                    "required": ["asset_type", "label", "caption", "placement", "visual_present", "status"],
                    "additionalProperties": False,
                },
            },
            "detailed_methodology_analysis": {
                "type": "object",
                "properties": {
                    "sample_size_check": {
                        "type": "object",
                        "properties": {
                            "status": {"type": "string", "enum": ["PASS", "WARN", "FAIL"]},
                            "reported_n": {"type": "string"},
                            "explanation": {"type": "string"},
                        },
                        "required": ["status", "reported_n", "explanation"],
                        "additionalProperties": False,
                    },
                    "software_and_reproducibility": {
                        "type": "object",
                        "properties": {
                            "status": {"type": "string", "enum": ["PASS", "WARN", "FAIL"]},
                            "tools_identified": {"type": "string"},
                            "missing_parameters": {"type": "string"},
                        },
                        "required": ["status", "tools_identified", "missing_parameters"],
                        "additionalProperties": False,
                    },
                    "domain_specific_metrics": {
                        "type": "object",
                        "properties": {
                            "status": {"type": "string", "enum": ["PASS", "WARN", "FAIL"]},
                            "h_index_present": {"type": "boolean"},
                            "g_index_present": {"type": "boolean"},
                            "m_index_present": {"type": "boolean"},
                            "counting_method": {"type": "string"},
                            "analysis_notes": {"type": "string"},
                        },
                        "required": ["status", "h_index_present", "g_index_present", "m_index_present", "counting_method", "analysis_notes"],
                        "additionalProperties": False,
                    },
                },
                "required": ["sample_size_check", "software_and_reproducibility", "domain_specific_metrics"],
                "additionalProperties": False,
            },
            "facet_narrative_evaluations": {
                "type": "object",
                "properties": {
                    "research_gap_novelty": {"type": "string"},
                    "methodological_rigor": {"type": "string"},
                    "data_discussion_alignment": {"type": "string"},
                },
                "required": ["research_gap_novelty", "methodological_rigor", "data_discussion_alignment"],
                "additionalProperties": False,
            },
        },
        "required": [
            "manuscript_meta",
            "structural_template_audit",
            "visual_asset_audit",
            "detailed_methodology_analysis",
            "facet_narrative_evaluations",
        ],
        "additionalProperties": False,
    },
}


# ============================================================
# RESILIENT GROQ AUDIT ENGINE (PREVENTS 400 BAD REQUEST)
# ============================================================

def run_editorial_audit(raw_text: str, asset_meta: Dict[str, Any], client: Groq) -> Dict[str, Any]:
    """Execute the editorial audit with multiple fallback layers against HTTP 400 Bad Request."""
    prompt = f"""You are a senior academic peer reviewer and editorial prescreening manager. Conduct a rigorous, line-by-line audit of this manuscript text.

ASSET METADATA (Cross-modal visual detection):
- Total embedded image elements found: {asset_meta['total_images']}
- Total data tables found: {asset_meta['total_tables']}
- Captions detected: {json.dumps(asset_meta['detected_captions'])}

--- SECTION 1: UNIVERSAL ACADEMIC TEMPLATE AUDIT ---
Evaluate each required section:
{json.dumps(REQUIRED_TEMPLATE_SECTIONS)}
1. Title & Abstract: Check if Abstract exceeds 200–250 words and if structured. Extract 3–5 key keywords.
2. Introduction: Check Research Problem, Background, and Research Gap/Novelty.
3. Materials and Methods: Check software versions, parameter settings, algorithm choices, normalization methods, and search strings.
4. Results and Discussion: Check if findings are backed by data.
5. Conclusions: Summarize main findings, implications, and limitations.
6. Declarations: Verify Funding, Conflicts of Interest, AI Usage, Acknowledgments, Ethics Approval.

--- SECTION 2: DETAILED METHODOLOGY & SCIENTIFIC AUDIT ---
1. Data & Sample Size Accounting: Mathematical consistency between raw data, filtering steps, and final sample size (N).
2. Bibliometric indicators: Check h-index, g-index, m-index, software versions (VOSviewer, Biblioshiny), normalization, and counting logic (Full vs Fractional).
3. Non-Bibliometric studies: Sample size justification, statistical assumptions, power analysis, PRISMA flow.

--- SECTION 3: VISUAL ASSET & CAPTION INTEGRITY ---
Audit all Figures and Tables. Confirm whether inline embedded assets exist for every caption found.

MANUSCRIPT EXCERPTS:
--- START OF TEXT ---
{raw_text[:12000]}
--- END OF TEXT ---

Return strictly a valid JSON object matching the requested schema."""

    messages = [
        {
            "role": "system",
            "content": "You are a senior academic peer reviewer. Audit the manuscript thoroughly with zero hallucinations. Output strictly valid JSON.",
        },
        {"role": "user", "content": prompt},
    ]

    # ATTEMPT 1: Primary model with json_schema and explicit token budget
    try:
        resp = client.chat.completions.create(
            model=PRIMARY_MODEL,
            messages=messages,
            temperature=0.1,
            reasoning_effort="low",
            include_reasoning=False,
            max_completion_tokens=4500,
            response_format={"type": "json_schema", "json_schema": AUDIT_STRICT_SCHEMA},
        )
        content = resp.choices[0].message.content or "{}"
        return clean_json_response(content)
    except (BadRequestError, Exception) as e1:
        st.warning(f"Note: Primary structured call adjusted due to provider constraints ({type(e1).__name__}). Switching to JSON mode fallback.")

    # ATTEMPT 2: Primary model with json_object mode (universal JSON enforcement)
    try:
        resp = client.chat.completions.create(
            model=PRIMARY_MODEL,
            messages=messages,
            temperature=0.1,
            reasoning_effort="low",
            include_reasoning=False,
            max_completion_tokens=4500,
            response_format={"type": "json_object"},
        )
        content = resp.choices[0].message.content or "{}"
        return clean_json_response(content)
    except (BadRequestError, Exception) as e2:
        st.warning(f"Note: Secondary fallback invoked on stable model {FALLBACK_MODEL}.")

    # ATTEMPT 3: Fallback model (llama-3.3-70b-versatile) with json_object mode
    resp = client.chat.completions.create(
        model=FALLBACK_MODEL,
        messages=messages,
        temperature=0.1,
        max_completion_tokens=4000,
        response_format={"type": "json_object"},
    )
    content = resp.choices[0].message.content or "{}"
    return clean_json_response(content)


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


def search_openalex_reviewers(keywords: List[str], region: str, mailto: str = "", max_candidates: int = 5) -> List[Dict[str, Any]]:
    query = " ".join(keywords[:5]).strip()
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


def reviewer_discovery_report(keywords: List[str], mailto: str = "") -> Dict[str, List[Dict[str, Any]]]:
    output = {}
    for r in ["India", "Northeast India", "Assam"]:
        try:
            output[r] = search_openalex_reviewers(keywords, r, mailto=mailto)
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
    meta = audit.get("manuscript_meta", {})
    doc.add_heading("Academic Editorial Pre-Screening Audit Report", 0)
    doc.add_paragraph(f"Manuscript Title: {meta.get('title', 'Not specified')}")
    doc.add_paragraph(f"Editorial Verdict: {meta.get('editorial_verdict', 'Under Review')}")
    doc.add_paragraph(f"Verdict Rationale: {meta.get('verdict_rationale', '')}")

    doc.add_heading("1. Universal Academic Template Audit", level=1)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    h = tbl.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Section", "Detected", "Status", "Critique"
    for item in audit.get("structural_template_audit", []):
        row = tbl.add_row().cells
        row[0].text = item.get("section_name", "")
        row[1].text = "Yes" if item.get("detected") else "No"
        row[2].text = item.get("status", "")
        row[3].text = item.get("critique", "")

    doc.add_heading("2. Detailed Methodology & Scientific Audit", level=1)
    d_meth = audit.get("detailed_methodology_analysis", {})
    s_check = d_meth.get("sample_size_check", {})
    doc.add_paragraph(f"Sample Size Accounting: [{s_check.get('status', 'N/A')}] Reported N = {s_check.get('reported_n', 'N/A')}")
    doc.add_paragraph(f"Filtering Logic: {s_check.get('explanation', '')}")

    s_rep = d_meth.get("software_and_reproducibility", {})
    doc.add_paragraph(f"Software Tools: {s_rep.get('tools_identified', 'None')}")
    doc.add_paragraph(f"Missing Parameters: {s_rep.get('missing_parameters', 'None')}")

    d_metr = d_meth.get("domain_specific_metrics", {})
    doc.add_paragraph(f"Counting Method: {d_metr.get('counting_method', 'Unspecified')}")
    doc.add_paragraph(f"Domain Metrics Notes: {d_metr.get('analysis_notes', '')}")

    doc.add_heading("3. Visual Asset & Caption Integrity", level=1)
    for v in audit.get("visual_asset_audit", []):
        doc.add_paragraph(f"• {v.get('asset_type')} {v.get('label')} [{v.get('status')}]: {v.get('caption')} (Placement: {v.get('placement')}, Graphic Present: {v.get('visual_present')})")

    doc.add_heading("4. Reviewer Candidates (OpenAlex)", level=1)
    for reg, cands in reviewers.items():
        doc.add_heading(reg, level=2)
        for c in cands:
            if "name" in c:
                doc.add_paragraph(f"- {c['name']} ({c.get('institution', 'N/A')}): {len(c.get('recent_pubs', []))} verified publications")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI (STRUCTURED DASHBOARD - NO RAW JSON)
# ============================================================

st.set_page_config(page_title="Academic Manuscript Audit", page_icon="🎓", layout="wide")
st.title("🎓 Academic Editorial Pre-Screening & Scientific Audit")
st.caption("Universal template compliance, methodological logic & sample size validation, and visual asset integrity.")

with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key = st.text_input("Groq API Key", value=get_secret("GROQ_API_KEY"), type="password")
    openalex_mailto = st.text_input("OpenAlex Mailto Email", value=get_secret("OPENALEX_MAILTO", "editor@academicprescreen.org"))
    st.markdown("---")
    st.markdown("**Strict Scientific Checkpoints:**")
    st.markdown("1. **Structural Audit**: Abstract 200–250w, IMRaD, Declarations & Governance.")
    st.markdown("2. **Methodology Audit**: Sample accounting ($N$), h/g/m indices, counting logic.")
    st.markdown("3. **Visual Integrity**: Inline graphic verification for every caption.")

uploaded_file = st.file_uploader("Upload Manuscript (.pdf or .docx)", type=["pdf", "docx"])

if uploaded_file and st.button("🚀 Conduct Line-by-Line Academic Audit", type="primary"):
    if not groq_api_key:
        st.error("Please provide a valid Groq API Key.")
        st.stop()

    try:
        with st.spinner("Extracting text and auditing inline visual elements..."):
            raw_text, asset_meta = extract_text_and_assets(uploaded_file)
            if not raw_text.strip():
                st.error("Could not extract readable text from the uploaded document.")
                st.stop()

        with st.spinner("Conducting line-by-line editorial and methodological review..."):
            client = Groq(api_key=groq_api_key)
            audit_result = run_editorial_audit(raw_text, asset_meta, client)

        with st.spinner("Discovering verified regional peer reviewers via OpenAlex..."):
            keywords = audit_result.get("manuscript_meta", {}).get("keywords", [])
            reviewers = reviewer_discovery_report(keywords, mailto=openalex_mailto)

        with st.spinner("Assembling editorial reports..."):
            orig_bytes = uploaded_file.getvalue()
            blind_bytes = blind_copy_docx(orig_bytes) if uploaded_file.name.endswith(".docx") else orig_bytes
            docx_report = generate_docx_report(audit_result, reviewers)

        st.session_state["audit"] = audit_result
        st.session_state["reviewers"] = reviewers
        st.session_state["blind_bytes"] = blind_bytes
        st.session_state["docx_report"] = docx_report
        st.success("Manuscript audit completed.")

    except Exception as e:
        st.error(f"Audit processing encountered an error: {e}")
        st.info("Tip: If you encounter token limits, try uploading a slightly shorter excerpt or verify your Groq API key quota.")

# Display results if audit completed
if "audit" in st.session_state:
    audit = st.session_state["audit"]
    reviewers = st.session_state["reviewers"]
    meta = audit.get("manuscript_meta", {})
    verdict = meta.get("editorial_verdict", "Under Review")

    st.markdown("---")
    v_col1, v_col2 = st.columns([1, 3])
    with v_col1:
        if verdict == "Accept with Minor Revisions":
            st.success(f"### Verdict:\n**{verdict}**")
        elif verdict == "Major Revisions":
            st.warning(f"### Verdict:\n**{verdict}**")
        else:
            st.error(f"### Verdict:\n**{verdict}**")
    with v_col2:
        st.subheader(meta.get("title", "Manuscript Title"))
        st.write(f"**Rationale:** {meta.get('verdict_rationale', '')}")
        st.write("**Extracted Keywords:** " + ", ".join([f"`{k}`" for k in meta.get("keywords", [])]))

    # DASHBOARD TABS (NO RAW JSON DISPLAY)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏛️ Section 1: Template Audit",
        "🔬 Section 2: Methodology & Data Logic",
        "🖼️ Section 3: Visual Asset Integrity",
        "📝 In-Depth Narrative Critique",
        "👥 Verified Reviewer Discovery",
    ])

    # TAB 1: Structural Template Audit
    with tab1:
        st.subheader("Universal Academic Template Compliance")
        struct_data = audit.get("structural_template_audit", [])
        table_rows = []
        for s in struct_data:
            badge = "🟢 PASS" if s.get("status") == "PASS" else ("🟡 WARN" if s.get("status") == "WARN" else "🔴 FAIL")
            table_rows.append({
                "Section": s.get("section_name"),
                "Detected": "✅ Yes" if s.get("detected") else "❌ No",
                "Status": badge,
                "Heading In Text": s.get("heading_evidence"),
                "Line-by-Line Critique": s.get("critique"),
            })
        st.dataframe(table_rows, use_container_width=True, hide_index=True)

    # TAB 2: Detailed Methodology & Scientific Audit
    with tab2:
        st.subheader("Detailed Methodology & Scientific Rigor")
        meth = audit.get("detailed_methodology_analysis", {})

        # Sample size
        sc = meth.get("sample_size_check", {})
        st.markdown("#### 1. Data & Sample Size Accounting")
        sc_col1, sc_col2 = st.columns([1, 3])
        sc_col1.metric("Sample Check Status", sc.get("status", "N/A"), f"N = {sc.get('reported_n', 'N/A')}")
        sc_col2.info(f"**Filtering & Arithmetic Evaluation:**\n{sc.get('explanation', '')}")

        # Software
        st.markdown("#### 2. Software Parameters & Reproducibility")
        sw = meth.get("software_and_reproducibility", {})
        sw_col1, sw_col2 = st.columns(2)
        sw_col1.write(f"**Tools & Packages Identified:**\n{sw.get('tools_identified', 'None')}")
        sw_col2.warning(f"**Missing Parameters & Version Details:**\n{sw.get('missing_parameters', 'None')}")

        # Domain metrics
        st.markdown("#### 3. Domain-Specific & Scientometric Metrics")
        dm = meth.get("domain_specific_metrics", {})
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("h-index Reported", "Yes" if dm.get("h_index_present") else "No")
        m2.metric("g-index Reported", "Yes" if dm.get("g_index_present") else "No")
        m3.metric("m-index Reported", "Yes" if dm.get("m_index_present") else "No")
        m4.metric("Counting Method", dm.get("counting_method", "Unspecified"))
        st.write(f"**Scientometric / Statistical Analysis Notes:**\n{dm.get('analysis_notes', '')}")

    # TAB 3: Visual Asset & Caption Integrity
    with tab3:
        st.subheader("Visual Asset & Caption Integrity")
        v_data = audit.get("visual_asset_audit", [])
        if not v_data:
            st.info("No figures or tables detected in the document.")
        else:
            v_rows = []
            for v in v_data:
                v_badge = "🟢 PASS" if v.get("status") == "PASS" else "🔴 FAIL"
                v_rows.append({
                    "Type": v.get("asset_type"),
                    "Label": v.get("label"),
                    "Caption Text": v.get("caption"),
                    "Placement": v.get("placement"),
                    "Inline Graphic Present": "✅ Yes" if v.get("visual_present") else "❌ Missing Graphic",
                    "Status": v_badge,
                })
            st.dataframe(v_rows, use_container_width=True, hide_index=True)

    # TAB 4: In-Depth Narrative Critique
    with tab4:
        st.subheader("In-Depth Scientific Narrative Evaluations")
        fn = audit.get("facet_narrative_evaluations", {})
        with st.expander("🔍 Research Problem, Gap & Novelty", expanded=True):
            st.write(fn.get("research_gap_novelty", "Not evaluated."))
        with st.expander("🧪 Methodological Rigor & Parameter Clarity", expanded=True):
            st.write(fn.get("methodological_rigor", "Not evaluated."))
        with st.expander("📊 Data-Discussion Alignment & Evidence", expanded=True):
            st.write(fn.get("data_discussion_alignment", "Not evaluated."))

    # TAB 5: Reviewer Candidates
    with tab5:
        st.subheader("Verified OpenAlex Peer Reviewer Candidates")
        st.caption("Retrieved from publication records based on extracted research keywords. Verified institutional associations.")
        for region, cands in reviewers.items():
            st.markdown(f"### Region: {region}")
            if not cands or "error" in cands[0]:
                st.write("No eligible candidates found in this region.")
                continue
            r_rows = []
            for c in cands:
                pubs = c.get("recent_pubs", [])
                latest_title = pubs[0]["title"] if pubs else "N/A"
                r_rows.append({
                    "Candidate Name": c.get("name"),
                    "Affiliated Institution": c.get("institution"),
                    "Verified Works": len(pubs),
                    "Recent Representative Publication": latest_title,
                })
            st.dataframe(r_rows, use_container_width=True, hide_index=True)

    # DOWNLOAD SECTION
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
