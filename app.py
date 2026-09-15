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
from groq import Groq


# ============================================================
# CONFIGURATION & CONSTANTS
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
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


def extract_section_excerpts(text: str) -> Dict[str, str]:
    """Segment text into core sections for focused model evaluation."""
    lines = text.splitlines()
    found: Dict[str, int] = {}

    for i, line in enumerate(lines):
        clean = normalize(line)
        if not clean or len(clean) > 85:
            continue
        for canonical, pattern in SECTION_PATTERNS.items():
            if canonical not in found and re.match(pattern, clean):
                found[canonical] = i

    sorted_sections = sorted(found.items(), key=lambda x: x[1])
    excerpts = {}
    for idx, (canonical, start_line) in enumerate(sorted_sections):
        end_line = sorted_sections[idx + 1][1] if idx + 1 < len(sorted_sections) else len(lines)
        excerpts[canonical] = "\n".join(lines[start_line:end_line]).strip()

    return excerpts


# ============================================================
# GROQ STRUCTURED AUDIT ENGINE (EXACT SCHEMA COMPLIANCE)
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
                        "asset_type": {"type": "string", "enum": ["Table", "Figure"]},
                        "label": {"type": "string"},
                        "caption": {"type": "string"},
                        "placement": {"type": "string", "enum": ["Inline Body", "Appendix", "Missing"]},
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
                            "counting_method": {
                                "type": "string",
                                "enum": ["Full-counting", "Fractional-counting", "Unspecified"],
                            },
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


def run_editorial_audit(raw_text: str, asset_meta: Dict[str, Any], client: Groq) -> Dict[str, Any]:
    """Execute the full editorial line-by-line audit across all 3 directive sections."""
    excerpts = extract_section_excerpts(raw_text)

    prompt = f"""You are a senior academic peer reviewer and editorial prescreening manager. Conduct a rigorous, line-by-line audit of this manuscript text.

ASSET METADATA (Cross-modal visual detection):
- Total embedded image elements found: {asset_meta['total_images']}
- Total data tables found: {asset_meta['total_tables']}
- Captions detected: {json.dumps(asset_meta['detected_captions'])}

--- SECTION 1: UNIVERSAL ACADEMIC TEMPLATE AUDIT ---
Evaluate each required section:
{json.dumps(REQUIRED_TEMPLATE_SECTIONS)}
1. Title & Abstract: Check if Abstract exceeds 200–250 words and if structured (Background, Methods, Results, Conclusion). Extract 3–5 key keywords.
2. Introduction: Must clearly state Research Problem, Background, and highlight Research Gap/Novelty.
3. Materials and Methods: Check for exact software versions, parameter settings, algorithm choices, normalization methods, and search strings.
4. Results and Discussion: Check if findings are backed by data.
5. Conclusions: Must summarize main findings, implications, and limitations.
6. Declarations & Governance: Verify presence of Funding, Conflicts of Interest, AI Usage, Acknowledgments, and Ethics Approval.

--- SECTION 2: DETAILED METHODOLOGY & SCIENTIFIC AUDIT ---
1. Data & Sample Size Accounting: Verify mathematical consistency between raw dataset, filters, and final sample size (N).
2. Bibliometric/Scientometric indicators (if applicable):
   - Check h-index, g-index, m-index.
   - Verify parameters for VOSviewer, Bibliometrix (Biblioshiny), CiteSpace, Pajek.
   - Verify network normalization methods (Association Strength, Fractionalization, Cosine Similarity).
   - Check counting logic (Full-counting vs. Fractional-counting).
3. Non-Bibliometric studies: Check sample size justification, statistical test assumptions, power analysis, PRISMA flow.

--- SECTION 3: VISUAL ASSET & CAPTION INTEGRITY ---
Audit all Figures and Tables. Confirm whether inline embedded assets exist for every caption found.

MANUSCRIPT EXCERPTS:
--- START OF TEXT ---
{raw_text[:14000]}
--- END OF TEXT ---

Return ONLY valid JSON matching the schema."""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a senior academic peer reviewer. Audit the manuscript thoroughly with zero hallucinations. Output strictly valid JSON.",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0,
        response_format={"type": "json_schema", "json_schema": AUDIT_STRICT_SCHEMA},
    )

    return json.loads(response.choices[0].message.content or "{}")


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
        # Keep title (first non-empty paragraph), delete author block
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

    # Redact email and ORCID
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

    # Section 1
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

    # Section 2
    doc.add_heading("2. Detailed Methodology & Scientific Audit", level=1)
    d_meth = audit.get("detailed_methodology_analysis", {})
    s_check = d_meth.get("sample_size_check", {})
    doc.add_paragraph(f"Sample Size Accounting: [{s_check.get('status', 'N/A')}] Reported N = {s_check.get('reported_n', 'N/A')}")
    doc.add_paragraph(f"
