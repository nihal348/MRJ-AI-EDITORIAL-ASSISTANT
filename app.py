import io
import json
import os
import re
from copy import deepcopy
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber
import requests
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from groq import Groq


# ============================================================
# CONFIGURATION & EVALUATION DIRECTIVES
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
OPENALEX_URL = "https://api.openalex.org/works"
OPENALEX_AUTHOR_URL = "https://api.openalex.org/authors"

MRJ_RULES = {
    "abstract_target_words": 200,
    "abstract_soft_limit_words": 220,  # Directive 5: Soften penalties for minor overflows
    "keywords_min": 3,
    "keywords_max": 10,
    "minimum_domains": 2,
    "required_sections": [
        "Introduction",
        "Materials and Methods",
        "Results and Discussion",
        "Conclusions",
        "Multidisciplinary Domains",
        "Funding",
        "Acknowledgments",
        "Conflicts of Interest",
        "Declaration on AI Usage",
        "References",
    ],
}

# Directive 1: Comprehensive section regex recognizing numerical prefixes, roman numerals, and synonyms
SECTION_HEADING_PATTERNS = {
    "abstract": r"(?i)^\s*(?:abstract|executive\s+summary)\s*:?$",
    "introduction": r"(?i)^\s*(?:(?:section\s+)?1(?:\.0?)?|[ivx]+\.?)?\s*(?:introduction|background)\s*:?$",
    "materials and methods": r"(?i)^\s*(?:(?:section\s+)?2(?:\.0?)?|[ivx]+\.?)?\s*(?:materials\s+and\s+methods|materials\s+&\s+methods|methodology|methods|experimental\s+procedures)\s*:?$",
    "results and discussion": r"(?i)^\s*(?:(?:section\s+)?3(?:\.0?)?|[ivx]+\.?)?\s*(?:results\s+and\s+discussion|results\s+&\s+discussion|results|findings\s+and\s+discussion)\s*:?$",
    "conclusions": r"(?i)^\s*(?:(?:section\s+)?4(?:\.0?)?|[ivx]+\.?)?\s*(?:conclusions?|summary\s+and\s+conclusions?|concluding\s+remarks)\s*:?$",
    "multidisciplinary domains": r"(?i)^\s*(?:multidisciplinary\s+domains?|research\s+domains?)\s*:?$",
    "funding": r"(?i)^\s*(?:funding(?:\s+information)?|financial\s+support|grant\s+support)\s*:?$",
    "acknowledgments": r"(?i)^\s*(?:acknowledgments?|acknowledgements?)\s*:?$",
    "conflicts of interest": r"(?i)^\s*(?:conflicts?\s+of\s+interest|competing\s+interests?|disclosure\s+statement)\s*:?$",
    "declaration on ai usage": r"(?i)^\s*(?:declaration\s+on\s+ai(?:\s+usage)?|generative\s+ai\s+statement|ai\s+usage)\s*:?$",
    "references": r"(?i)^\s*(?:references?|bibliography|literature\s+cited)\s*:?$",
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
# GENERAL HELPERS & TEXT NORMALIZATION
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
# DIRECTIVE 1 & 2: MULTI-MODAL ASSET & SECTION EXTRACTION
# ============================================================

def extract_text_and_assets(uploaded_file) -> Tuple[str, Dict[str, Any]]:
    """Extract document text alongside visual assets, image counts, and tables."""
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    asset_meta = {
        "file_type": "docx" if name.endswith(".docx") else "pdf",
        "total_images": 0,
        "total_tables": 0,
        "pages_or_paragraphs": 0,
        "table_locations": [],
        "detected_captions": [],
    }

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = []

        # Count tables
        asset_meta["total_tables"] = len(doc.tables)

        # Count inline images via xml drawing tags
        xml_text = doc._element.xml
        asset_meta["total_images"] = len(re.findall(r"<a:blip|<w:drawing", xml_text))

        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                parts.append(txt)
                # Check for figure/table caption markers
                if re.match(r"(?i)^(figure|fig\.?|table)\s+\d+", txt):
                    asset_meta["detected_captions"].append(txt)

        for t_idx, table in enumerate(doc.tables):
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))
            asset_meta["table_locations"].append(f"Embedded Table #{t_idx + 1}")

        asset_meta["pages_or_paragraphs"] = len(doc.paragraphs)
        return "\n".join(parts), asset_meta

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            asset_meta["pages_or_paragraphs"] = len(pdf.pages)
            for page_idx, page in enumerate(pdf.pages):
                page_text = page.extract_text() or ""
                pages.append(page_text)

                # Count visual image objects per page
                images_on_page = len(page.images)
                asset_meta["total_images"] += images_on_page

                # Extract and count tables
                tables_on_page = page.extract_tables() or []
                asset_meta["total_tables"] += len(tables_on_page)

                for line in page_text.splitlines():
                    sline = line.strip()
                    if re.match(r"(?i)^(figure|fig\.?|table)\s+\d+", sline):
                        asset_meta["detected_captions"].append({
                            "caption": sline,
                            "page": page_idx + 1,
                            "page_has_images": images_on_page > 0,
                            "page_has_tables": len(tables_on_page) > 0,
                        })

        return "\n".join(pages), asset_meta

    raise ValueError("Unsupported file format. Please upload a .docx or .pdf file.")


def find_section_details(lines: List[str]) -> Dict[str, Dict[str, Any]]:
    """Directive 1: Robust absence verification using regex headings and quote extracts."""
    found: Dict[str, Dict[str, Any]] = {}
    for i, line in enumerate(lines):
        clean = normalize(line)
        if not clean or len(clean) > 90:
            continue
        for canonical, pattern in SECTION_HEADING_PATTERNS.items():
            if canonical in found:
                continue
            if re.match(pattern, clean):
                # Fetch first substantial non-empty subsequent line as evidence quote
                first_line = ""
                for nxt in lines[i + 1: i + 6]:
                    if nxt.strip() and not any(re.match(p, nxt.strip()) for p in SECTION_HEADING_PATTERNS.values()):
                        first_line = nxt.strip()[:140]
                        break
                found[canonical] = {
                    "line_index": i,
                    "heading_text": clean,
                    "first_line_quote": first_line or "Heading detected without direct text body.",
                }
    return found


def extract_section_text(text: str, canonical: str) -> str:
    lines = text.splitlines()
    sections = find_section_details(lines)
    if canonical not in sections:
        return ""
    start = sections[canonical]["line_index"]
    later_starts = [s["line_index"] for k, s in sections.items() if s["line_index"] > start]
    end = min(later_starts) if later_starts else len(lines)
    return "\n".join(lines[start + 1:end]).strip()


def extract_abstract(text: str) -> str:
    lines = text.splitlines()
    sections = find_section_details(lines)
    if "abstract" in sections:
        start = sections["abstract"]["line_index"]
        end_candidates = [
            i for i, line in enumerate(lines)
            if i > start and (re.match(r"(?i)^\s*keywords?\s*:", line) or any(re.match(p, line.strip()) for p in SECTION_HEADING_PATTERNS.values()))
        ]
        end = min(end_candidates) if end_candidates else len(lines)
        val = "\n".join(lines[start + 1:end]).strip()
        if not val and ":" in lines[start]:
            val = lines[start].split(":", 1)[1]
        return normalize(val)
    m = re.search(r"(?is)\babstract\s*:\s*(.*?)(?:\bkeywords?\s*:|$)", text)
    return normalize(m.group(1)) if m else ""


def extract_keywords(text: str) -> List[str]:
    m = re.search(r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:(?:1\.?\s+)?introduction|materials|abstract)\b|$)", text)
    if not m:
        return []
    raw = m.group(1).strip().splitlines()[0]
    return [normalize(x) for x in re.split(r"[;,]", raw) if normalize(x)]


# ============================================================
# DIRECTIVES 1, 2, 3, 4, 5: DETERMINISTIC AUDIT
# ============================================================

def run_evidence_based_precheck(text: str, asset_meta: Dict[str, Any]) -> Dict[str, Any]:
    lines = text.splitlines()
    section_map = find_section_details(lines)
    abstract_txt = extract_abstract(text)
    keywords_list = extract_keywords(text)
    methods_txt = extract_section_text(text, "materials and methods")
    results_txt = extract_section_text(text, "results and discussion")
    references_txt = extract_section_text(text, "references")

    # Title detection
    title_candidate = lines[0].strip() if lines else "Manuscript Title Not Detected"
    for l in lines[:5]:
        if len(l.strip()) > 15 and not re.match(r"(?i)^(abstract|volume|issn|page)", l.strip()):
            title_candidate = l.strip()
            break

    # 1. Structural Section Checks
    structural_checks = []
    for req in MRJ_RULES["required_sections"]:
        canonical = req.lower()
        if canonical in section_map:
            det = section_map[canonical]
            structural_checks.append({
                "section_name": req,
                "detected": True,
                "detected_heading_text": det["heading_text"],
                "first_line_quote": det["first_line_quote"],
                "status": "PASS",
            })
        else:
            structural_checks.append({
                "section_name": req,
                "detected": False,
                "detected_heading_text": "None",
                "first_line_quote": "Section heading not identified in layout.",
                "status": "FAIL" if req in ["Introduction", "Materials and Methods", "Results and Discussion", "Conclusions", "References"] else "WARN",
            })

    # 2. Directive 2: Visual & Layout Asset Audit
    visual_asset_checks = []
    detected_caps = asset_meta.get("detected_captions", [])

    if not detected_caps:
        if asset_meta["total_images"] > 0 or asset_meta["total_tables"] > 0:
            visual_asset_checks.append({
                "label": "General Assets",
                "caption_found": False,
                "visual_image_present": True,
                "placement_location": f"Embedded: {asset_meta['total_images']} image(s), {asset_meta['total_tables']} table(s)",
                "status": "WARN",
                "notes": "Visual elements exist but formal numbered captions (e.g. 'Figure 1:') were not detected.",
            })
        else:
            visual_asset_checks.append({
                "label": "Visual Assets",
                "caption_found": False,
                "visual_image_present": False,
                "placement_location": "None",
                "status": "PASS",
                "notes": "No visual figures or graphic assets claimed or detected.",
            })
    else:
        for item in detected_caps:
            if isinstance(item, dict):
                cap_text = item["caption"]
                has_img = item["page_has_images"]
                has_tbl = item["page_has_tables"]
                loc = f"Page {item['page']}"
            else:
                cap_text = str(item)
                has_img = asset_meta["total_images"] > 0
                has_tbl = asset_meta["total_tables"] > 0
                loc = "Document Body"

            is_fig = bool(re.search(r"(?i)\bfig", cap_text))
            is_tbl = bool(re.search(r"(?i)\btable", cap_text))

            if is_fig and not has_img and asset_meta["total_images"] == 0:
                visual_asset_checks.append({
                    "label": cap_text[:35],
                    "caption_found": True,
                    "visual_image_present": False,
                    "placement_location": loc,
                    "status": "FAIL",
                    "notes": "Figure caption detected but rendering contains no graphic/image element.",
                })
            elif is_tbl and not has_tbl and asset_meta["total_tables"] == 0:
                visual_asset_checks.append({
                    "label": cap_text[:35],
                    "caption_found": True,
                    "visual_image_present": False,
                    "placement_location": loc,
                    "status": "WARN",
                    "notes": "Table caption detected without formatted table structure.",
                })
            else:
                visual_asset_checks.append({
                    "label": cap_text[:35],
                    "caption_found": True,
                    "visual_image_present": True,
                    "placement_location": loc,
                    "status": "PASS",
                    "notes": "Caption matches presence of embedded asset.",
                })

    # 3. Directive 4 & 5: Citation & Formatting Issues
    citation_formatting_issues = []

    # Directive 5: Abstract word count pragmatic thresholding
    if abstract_txt:
        cnt = word_count(abstract_txt)
        if cnt > MRJ_RULES["abstract_soft_limit_words"]:
            citation_formatting_issues.append({
                "issue_type": "Word Count Limit",
                "evidence_quote": f"Abstract contains {cnt} words (standard target is {MRJ_RULES['abstract_target_words']}).",
                "severity": "Major",
            })
        elif cnt > MRJ_RULES["abstract_target_words"]:
            citation_formatting_issues.append({
                "issue_type": "Word Count Limit",
                "evidence_quote": f"Abstract is {cnt} words; minor overflow over 200-word limit.",
                "severity": "Minor",
            })
    else:
        citation_formatting_issues.append({
            "issue_type": "Word Count Limit",
            "evidence_quote": "No abstract detected in manuscript.",
            "severity": "Major",
        })

    # Directive 4: Inline bracket checks in body
    body_without_refs = text[: text.lower().find("references")] if "references" in text.lower() else text
    inline_brackets = re.findall(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]", body_without_refs)
    if not inline_brackets:
        citation_formatting_issues.append({
            "issue_type": "Missing Inline Brackets",
            "evidence_quote": "Discussion and body text lack standard bracketed inline citations (e.g., [1]).",
            "severity": "Major",
        })

    # Directive 4: Truncated references check
    truncated_refs = []
    for ref_line in references_txt.splitlines():
        ref_s = ref_line.strip()
        if re.match(r"^\[\d+\]\s*(?:reference\s+\d+|incomplete|\.{3,})$", ref_s, re.I) or (
            re.match(r"^\[\d+\]", ref_s) and len(ref_s) < 18
        ):
            truncated_refs.append(ref_s)

    if truncated_refs:
        citation_formatting_issues.append({
            "issue_type": "Truncated Reference",
            "evidence_quote": f"Detected {len(truncated_refs)} incomplete reference entry(ies): {', '.join(truncated_refs[:3])}",
            "severity": "Major",
        })

    # Ethics & IRB Approval checks
    participant_terms = ["participant", "patient", "human subject", "survey", "interview", "questionnaire"]
    has_participants = any(term in (methods_txt.lower() + " " + results_txt.lower()) for term in participant_terms)
    irb_terms = ["irb", "institutional review board", "ethics committee", "ethical approval", "protocol code", "waiver"]
    has_irb = any(term in (methods_txt.lower() + " " + text.lower()) for term in irb_terms)

    if has_participants and not has_irb:
        citation_formatting_issues.append({
            "issue_type": "Ethics Statement Missing",
            "evidence_quote": "Participant/survey terminology present in methods without explicit IRB protocol approval or waiver statement.",
            "severity": "Major",
        })

    return {
        "manuscript_title": title_candidate,
        "structural_checks": structural_checks,
        "visual_asset_checks": visual_asset_checks,
        "citation_formatting_issues": citation_formatting_issues,
        "abstract": abstract_txt,
        "keywords": keywords_list,
        "methods_text": methods_txt,
        "results_text": results_txt,
    }


# ============================================================
# GROQ AI STRUCTURED AUDIT (SCHEMA COMPLIANT)
# ============================================================

EVALUATION_JSON_SCHEMA = {
    "name": "manuscript_editorial_audit",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "manuscript_title": {"type": "string"},
            "structural_section_checks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "section_name": {"type": "string"},
                        "detected": {"type": "boolean"},
                        "detected_heading_text": {"type": "string"},
                        "first_line_quote": {"type": "string"},
                        "status": {"type": "string", "enum": ["PASS", "WARN", "FAIL"]},
                    },
                    "required": ["section_name", "detected", "detected_heading_text", "first_line_quote", "status"],
                    "additionalProperties": False,
                },
            },
            "visual_asset_checks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "caption_found": {"type": "boolean"},
                        "visual_image_present": {"type": "boolean"},
                        "placement_location": {"type": "string"},
                        "status": {"type": "string", "enum": ["PASS", "FAIL", "WARN"]},
                        "notes": {"type": "string"},
                    },
                    "required": ["label", "caption_found", "visual_image_present", "placement_location", "status", "notes"],
                    "additionalProperties": False,
                },
            },
            "methodology_and_math_logic": {
                "type": "object",
                "properties": {
                    "sample_size_check": {"type": "string", "enum": ["PASS", "WARN"]},
                    "math_discrepancies": {"type": "array", "items": {"type": "string"}},
                    "missing_domain_metrics": {"type": "array", "items": {"type": "string"}},
                    "software_parameter_notes": {"type": "string"},
                },
                "required": ["sample_size_check", "math_discrepancies", "missing_domain_metrics", "software_parameter_notes"],
                "additionalProperties": False,
            },
            "citation_and_formatting_issues": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "issue_type": {"type": "string"},
                        "evidence_quote": {"type": "string"},
                        "severity": {"type": "string", "enum": ["Minor", "Major"]},
                    },
                    "required": ["issue_type", "evidence_quote", "severity"],
                    "additionalProperties": False,
                },
            },
            "editorial_recommendation": {
                "type": "object",
                "properties": {
                    "verdict": {
                        "type": "string",
                        "enum": ["Accept as is", "Accept with Minor Revisions", "Major Revisions", "Reject"],
                    },
                    "key_reasons": {"type": "array", "items": {"type": "string"}},
                },
                "required": ["verdict", "key_reasons"],
                "additionalProperties": False,
            },
            "research_area": {"type": "string"},
            "reviewer_search_keywords": {"type": "array", "items": {"type": "string"}},
        },
        "required": [
            "manuscript_title",
            "structural_section_checks",
            "visual_asset_checks",
            "methodology_and_math_logic",
            "citation_and_formatting_issues",
            "editorial_recommendation",
            "research_area",
            "reviewer_search_keywords",
        ],
        "additionalProperties": False,
    },
}


def run_groq_evidence_audit(
    raw_text: str,
    deterministic_data: Dict[str, Any],
    client: Groq,
) -> Dict[str, Any]:
    """Directive 3 & 5: AI cross-examination of sample size logic, software metrics, and recommendations."""
    prompt = f"""EVIDENCE-BASED MANUSCRIPT AUDIT & PRE-SCREENING:

MANUSCRIPT TITLE DETECTED: {deterministic_data['manuscript_title']}

DETERMINISTIC PRE-SCREENING DATA:
- Structural Sections Detected: {json.dumps(deterministic_data['structural_checks'], indent=2)}
- Visual Asset Audit: {json.dumps(deterministic_data['visual_asset_checks'], indent=2)}
- Citation & Formatting Issues Flagged: {json.dumps(deterministic_data['citation_formatting_issues'], indent=2)}

MANUSCRIPT EXCERPTS:
--- ABSTRACT ---
{deterministic_data['abstract'][:2500]}

--- METHODS ---
{deterministic_data['methods_text'][:4500]}

--- RESULTS & DISCUSSION ---
{deterministic_data['results_text'][:4500]}

EVALUATION DIRECTIVES:
1. ABSENCE VERIFICATION: Maintain structural check findings unless incontrovertible body text proves presence.
2. VISUAL ASSETS: Confirm whether visual element drops or misplaced figures exist.
3. METHODOLOGY & MATH LOGIC: Audit sample-size arithmetic. If subtotals sum to > N, evaluate whether full or fractional counting was declared. Verify domain-specific metrics (e.g., h-index, g-index, software versions, VOSviewer/R package parameters) and note un-filtered category noise.
4. CITATION & ETHICS: Flag any truncated bibliography entries and missing IRB statements.
5. PRAGMATIC THRESHOLDING: Soften formatting penalties (abstract overflow of <20 words is a Minor issue). Reserve 'Major Revisions' or 'Reject' strictly for missing body text, dropped figures, or flawed data logic.

Produce strictly the required JSON format."""

    resp = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a rigorous, evidence-based academic pre-screening assistant. Return strictly valid JSON following the schema.",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0,
        response_format={"type": "json_schema", "json_schema": EVALUATION_JSON_SCHEMA},
    )
    return json.loads(resp.choices[0].message.content or "{}")


# ============================================================
# OPENALEX REVIEWER DISCOVERY
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    res = requests.get(url, params=params, timeout=30)
    res.raise_for_status()
    return res.json()


def get_openalex_author(author_id: str, mailto: str = "") -> Dict[str, Any]:
    try:
        return openalex_get(f"{OPENALEX_AUTHOR_URL}/{author_id}", {}, mailto)
    except Exception:
        return {}


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


def search_openalex_reviewers(
    search_terms: List[str],
    region: str,
    mailto: str = "",
    max_candidates: int = 5,
) -> List[Dict[str, Any]]:
    query = " ".join(search_terms[:6]).strip()
    if not query:
        return []

    data = openalex_get(
        OPENALEX_URL,
        {"search": query, "per-page": 35, "sort": "publication_year:desc"},
        mailto=mailto,
    )

    candidates: Dict[str, Dict[str, Any]] = {}
    for work in data.get("results", []):
        year = work.get("publication_year") or 0
        title = work.get("display_name") or ""
        doi = work.get("doi") or ""
        cited = work.get("cited_by_count") or 0

        for authorship in work.get("authorships", []):
            author = authorship.get("author") or {}
            author_id = author.get("id")
            author_name = author.get("display_name")
            if not author_id or not author_name:
                continue

            institutions = authorship.get("institutions") or []
            if not institutions:
                continue

            for inst in institutions:
                inst_name = inst.get("display_name") or ""
                inst_country = inst.get("country_code") or ""
                candidate = {
                    "author_id": author_id,
                    "name": author_name,
                    "institution": inst_name,
                    "country": inst_country,
                    "city": (inst.get("geo") or {}).get("city") or "",
                    "region": (inst.get("geo") or {}).get("region") or "",
                    "recent_publications": [],
                    "score": 0.0,
                }

                if not candidate_region_match(candidate, region):
                    continue

                if author_id not in candidates:
                    candidates[author_id] = candidate

                candidates[author_id]["recent_publications"].append({
                    "title": title, "year": year, "doi": doi, "cited_by_count": cited
                })
                candidates[author_id]["score"] += max(0, year - 2018) * 0.8 + min(cited, 100) * 0.02

    verified = list(candidates.values())
    verified.sort(key=lambda x: (len(x["recent_publications"]), x["score"]), reverse=True)
    return verified[:max_candidates]


def reviewer_search_report(audit_json: Dict[str, Any], mailto: str = "") -> Dict[str, List[Dict[str, Any]]]:
    terms = audit_json.get("reviewer_search_keywords", []) + [audit_json.get("research_area", "")]
    terms = [normalize(t) for t in terms if normalize(t)]
    output = {}
    for r in ["India", "Northeast India", "Assam"]:
        try:
            output[r] = search_openalex_reviewers(terms, r, mailto=mailto)
        except Exception as exc:
            output[r] = [{"error": f"OpenAlex query error: {exc}"}]
    return output


# ============================================================
# BLIND REVIEW COPY BUILDER (NO AI REWRITING)
# ============================================================

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ORCID_RE = re.compile(r"\b(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}\b", re.I)
PHONE_RE = re.compile(r"(?<!\d)(?:\+?\d[\d\s().-]{7,}\d)(?!\d)")
URL_RE = re.compile(r"https?://\S+", re.I)

def redact_text(text: str) -> str:
    text = EMAIL_RE.sub("[REDACTED EMAIL]", text)
    text = ORCID_RE.sub("[REDACTED ORCID]", text)
    text = PHONE_RE.sub("[REDACTED PHONE]", text)
    return text


def blind_copy_docx(original_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(original_bytes))

    # Strip Word document metadata
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""
    props.subject = ""

    body_paras = list(doc.paragraphs)
    abstract_idx = None
    for i, p in enumerate(body_paras):
        if re.match(r"(?i)^\s*abstract\s*:?", p.text.strip()):
            abstract_idx = i
            break

    # In front matter: preserve title (first non-empty paragraph), delete author block
    if abstract_idx is not None:
        first_title_kept = False
        for i in range(abstract_idx):
            txt = body_paras[i].text.strip()
            if not txt:
                continue
            if not first_title_kept:
                first_title_kept = True
                continue
            p_elem = body_paras[i]._element
            p_elem.getparent().remove(p_elem)

    # Redact remaining paragraphs
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = redact_text(r.text)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# REPORT GENERATION (.DOCX)
# ============================================================

def generate_report_docx(audit: Dict[str, Any], reviewers: Dict[str, List[Dict[str, Any]]]) -> bytes:
    doc = Document()
    doc.add_heading("Evidence-Based Scientific Manuscript Audit", 0)
    doc.add_paragraph(f"Manuscript Title: {audit.get('manuscript_title', 'Untitled')}")
    rec = audit.get("editorial_recommendation", {})
    doc.add_paragraph(f"Verdict: {rec.get('verdict', 'Under Review')}")

    doc.add_heading("1. Structural Section Verification", level=1)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Section", "Detected", "Heading Evidence", "Status"
    for s in audit.get("structural_section_checks", []):
        row = tbl.add_row().cells
        row[0].text = s.get("section_name", "")
        row[1].text = "Yes" if s.get("detected") else "No"
        row[2].text = s.get("first_line_quote", "")
        row[3].text = s.get("status", "")

    doc.add_heading("2. Visual Asset Audit", level=1)
    for v in audit.get("visual_asset_checks", []):
        doc.add_paragraph(f"• {v.get('label', '')} [{v.get('status')}]: {v.get('notes', '')} (Location: {v.get('placement_location')})")

    doc.add_heading("3. Methodology & Mathematical Logic", level=1)
    mlog = audit.get("methodology_and_math_logic", {})
    doc.add_paragraph(f"Sample Size Check: {mlog.get('sample_size_check', 'PASS')}")
    doc.add_paragraph(f"Software & Parameter Notes: {mlog.get('software_parameter_notes', 'None')}")
    for d in mlog.get("math_discrepancies", []):
        doc.add_paragraph(f"• Discrepancy: {d}")

    doc.add_heading("4. Reviewer Candidates (OpenAlex)", level=1)
    for reg, cands in reviewers.items():
        doc.add_heading(reg, level=2)
        for c in cands:
            if "name" in c:
                doc.add_paragraph(f"- {c['name']} ({c.get('institution', 'N/A')}): {len(c.get('recent_publications', []))} recent publications")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT USER INTERFACE
# ============================================================

st.set_page_config(page_title="Evidence-Based Manuscript Pre-Screening", page_icon="🔬", layout="wide")
st.title("🔬 Evidence-Based Manuscript Audit & Editorial Pre-Screening")
st.caption("Cross-modal visual auditing, robust structural verification, sample-size mathematical cross-checks, and peer reviewer discovery.")

with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key = st.text_input("Groq API Key", value=get_secret("GROQ_API_KEY"), type="password")
    openalex_mailto = st.text_input("OpenAlex Mailto Email", value=get_secret("OPENALEX_MAILTO", "editorial@mrjournal.org"))
    st.markdown("---")
    st.markdown("### Evaluation Directives Active")
    st.markdown("1. **Absence Verification**: Scanning layout & numeric headings.")
    st.markdown("2. **Visual Asset Audit**: Cross-checking images vs. captions.")
    st.markdown("3. **Mathematical Logic**: Checking sample size ($N$) arithmetic.")
    st.markdown("4. **Citation & Ethics**: Checking brackets, truncation, and IRB.")
    st.markdown("5. **Pragmatic Thresholding**: Soft 220w abstract limit.")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file and st.button("🚀 Run Complete Evidence-Based Audit", type="primary"):
    if not groq_api_key:
        st.error("Please provide a Groq API Key.")
        st.stop()

    with st.spinner("Extracting text and performing visual asset cross-checks..."):
        raw_text, asset_meta = extract_text_and_assets(uploaded_file)
        deterministic_data = run_evidence_based_precheck(raw_text, asset_meta)

    with st.spinner("Executing Groq AI Structured Cross-Examination..."):
        client = Groq(api_key=groq_api_key)
        audit_json = run_groq_evidence_audit(raw_text, deterministic_data, client)

    with st.spinner("Retrieving verified reviewer profiles from OpenAlex..."):
        reviewers = reviewer_search_report(audit_json, mailto=openalex_mailto)

    with st.spinner("Preparing anonymized DOCX copy..."):
        orig_bytes = uploaded_file.getvalue()
        blind_docx_bytes = blind_copy_docx(orig_bytes) if uploaded_file.name.endswith(".docx") else orig_bytes

    report_docx_bytes = generate_report_docx(audit_json, reviewers)

    st.session_state["audit_result"] = audit_json
    st.session_state["reviewers"] = reviewers
    st.session_state["blind_bytes"] = blind_docx_bytes
    st.session_state["report_docx"] = report_docx_bytes
    st.success("Manuscript audit completed successfully.")

if "audit_result" in st.session_state:
    res = st.session_state["audit_result"]
    revs = st.session_state["reviewers"]

    # Top metrics banner
    rec = res.get("editorial_recommendation", {})
    verdict = rec.get("verdict", "Under Review")
    color = "green" if "Accept" in verdict else ("orange" if "Minor" in verdict else "red")
    st.markdown(f"### Editorial Verdict: :{color}[{verdict}]")
    for r in rec.get("key_reasons", []):
        st.markdown(f"- {r}")

    t1, t2, t3, t4, t5 = st.tabs([
        "📋 Full Audit JSON",
        "🏗️ Section Checks",
        "🖼️ Visual Assets",
        "📐 Math & Methodology",
        "👥 OpenAlex Reviewers",
    ])

    with t1:
        st.subheader("Strict Audit JSON Output")
        st.json(res)
        st.download_button(
            "📥 Download Audit JSON",
            data=json.dumps(res, indent=2),
            file_name="manuscript_audit_report.json",
            mime="application/json",
        )

    with t2:
        st.subheader("Structural Absence Verification")
        st.dataframe(res.get("structural_section_checks", []), use_container_width=True)

    with t3:
        st.subheader("Visual Asset & Layout Breaks")
        st.dataframe(res.get("visual_asset_checks", []), use_container_width=True)

    with t4:
        st.subheader("Methodological & Sample-Size Logic")
        mlog = res.get("methodology_and_math_logic", {})
        c1, c2 = st.columns(2)
        c1.metric("Sample Size Logic", mlog.get("sample_size_check", "PASS"))
        c2.write(f"**Software / Parameters:** {mlog.get('software_parameter_notes')}")
        if mlog.get("math_discrepancies"):
            st.error("Mathematical Discrepancies:")
            for disc in mlog["math_discrepancies"]:
                st.write(f"• {disc}")
        if mlog.get("missing_domain_metrics"):
            st.warning("Missing Domain Metrics:")
            for m in mlog["missing_domain_metrics"]:
                st.write(f"• {m}")

    with t5:
        st.subheader("Verified OpenAlex Reviewers")
        for reg, cands in revs.items():
            st.markdown(f"#### {reg}")
            if not cands or "error" in cands[0]:
                st.write("No candidates found or query error.")
                continue
            rows = []
            for c in cands:
                rows.append({
                    "Name": c.get("name"),
                    "Institution": c.get("institution"),
                    "Publications": len(c.get("recent_publications", [])),
                })
            st.dataframe(rows, use_container_width=True)

    st.markdown("---")
    c_down1, c_down2 = st.columns(2)
    c_down1.download_button(
        "📄 Download DOCX Editorial Report",
        data=st.session_state["report_docx"],
        file_name="Editorial_PreScreening_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    if uploaded_file and uploaded_file.name.endswith(".docx"):
        c_down2.download_button(
            "🙈 Download Anonymized Blind DOCX Copy",
            data=st.session_state["blind_bytes"],
            file_name="Anonymized_Blind_Reviewer_Copy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
