import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber
import requests
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from groq import Groq

# ============================================================
# CONFIGURATION & CONSTANTS
# ============================================================

DEFAULT_GROQ_MODEL = "openai/gpt-oss-120b"
OPENALEX_URL = "https://api.openalex.org/works"
OPENALEX_AUTHOR_URL = "https://api.openalex.org/authors"

MRJ_RULES = {
    "abstract_max_words": 200,
    "abstract_soft_overflow_limit": 220,  # <= 10% overflow is WARN, not FAIL
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

SECTION_ALIASES = {
    "introduction": ["introduction", "background and introduction"],
    "materials and methods": [
        "materials and methods",
        "materials & methods",
        "methods and materials",
        "methodology",
        "methods",
        "research methodology",
    ],
    "results and discussion": [
        "results and discussion",
        "results & discussion",
        "results",
        "discussion",
        "findings and discussion",
    ],
    "conclusions": [
        "conclusions",
        "conclusion",
        "summary and conclusions",
        "summary and conclusion",
        "concluding remarks",
        "summary",
    ],
    "multidisciplinary domains": ["multidisciplinary domains", "research domains"],
    "funding": ["funding", "financial support", "funding statement"],
    "acknowledgments": ["acknowledgments", "acknowledgements"],
    "conflicts of interest": [
        "conflicts of interest",
        "conflict of interest",
        "competing interests",
        "declaration of competing interest",
    ],
    "declaration on ai usage": [
        "declaration on ai usage",
        "ai usage statement",
        "declaration of generative ai",
        "artificial intelligence",
    ],
    "references": ["references", "reference", "bibliography", "literature cited"],
}

NORTHEAST_STATES = {
    "assam", "arunachal pradesh", "manipur", "meghalaya",
    "mizoram", "nagaland", "sikkim", "tripura"
}

NORTHEAST_INSTITUTIONS = [
    "iit guwahati", "tezpur university", "tezu", "nit silchar",
    "assam university", "gauhati university", "cotton university",
    "dibrugarh university", "nehu", "north-eastern hill university",
    "niser", "nit agartala", "manipur university", "mizoram university",
    "nagaland university", "tripura university", "rajiv gandhi university",
    "sikkim university"
]

ASSAM_TERMS = [
    "assam", "guwahati", "silchar", "tezpur", "dibrugarh", "jorhat",
    "iit guwahati", "gauhati university", "cotton university",
    "dibrugarh university", "assam university", "tezpur university",
    "nit silchar", "indian institute of technology guwahati"
]


# ============================================================
# EXTRACTION & VISUAL ASSET AUDIT
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


def strip_numerical_prefix(line: str) -> str:
    """Strip numerical prefixes such as '1.', 'Section 4:', 'IV.', or '3.2.'."""
    cleaned = re.sub(
        r"^(?:section\s*)?(?:[0-9]+(?:\.[0-9]+)*|[ivxlcdm]+)[\s.:–-]+\s*",
        "",
        line.strip(),
        flags=re.I,
    )
    return normalize(cleaned).lower().rstrip(":")


def extract_document_assets(uploaded_file) -> Tuple[str, List[Dict[str, Any]], List[str]]:
    """Extract raw text and audit visual assets (figures, drawings, tables)."""
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    full_text = ""
    visual_assets = []
    lines_list = []

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                parts.append(txt)
        for t_idx, table in enumerate(doc.tables, start=1):
            table_rows = []
            for row in table.rows:
                table_rows.append(" | ".join(cell.text.strip() for cell in row.cells))
            table_repr = "\n".join(table_rows)
            parts.append(f"[Table {t_idx}]\n" + table_repr)
            visual_assets.append({
                "label": f"Table {t_idx}",
                "caption_found": True,
                "visual_image_present": True,
                "placement_location": "In-line document body",
                "status": "PASS",
                "notes": f"Table detected with {len(table.rows)} rows and {len(table.columns)} columns."
            })
        full_text = "\n".join(parts)

    elif name.endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            page_text_list = []
            for page_idx, page in enumerate(pdf.pages, start=1):
                p_text = page.extract_text() or ""
                page_text_list.append(p_text)

                # Cross-modal visual asset checks
                images = page.images or []
                curves = page.curves or []
                tables = page.extract_tables() or []
                has_visual_stream = len(images) > 0 or len(curves) > 10

                # Detect figure captions on this page
                fig_captions = re.findall(
                    r"(?i)\b(Fig(?:ure)?\.?\s*\d+[a-z]?)\b\s*[:.-]?\s*([^\n]{5,100})",
                    p_text
                )
                for fig_id, cap_text in fig_captions:
                    clean_id = re.sub(r"\s+", " ", fig_id).title()
                    status = "PASS" if has_visual_stream else "FAIL"
                    visual_assets.append({
                        "label": clean_id,
                        "caption_found": True,
                        "visual_image_present": has_visual_stream,
                        "placement_location": f"Page {page_idx}",
                        "status": status,
                        "notes": (
                            f"Caption '{cap_text[:40]}...' confirmed with rendered visual data."
                            if has_visual_stream else
                            f"Caption detected on Page {page_idx} but NO visual stream/image element rendered."
                        )
                    })

                # Detect table captions
                tbl_captions = re.findall(
                    r"(?i)\b(Table\s*\d+[a-z]?)\b\s*[:.-]?\s*([^\n]{5,100})",
                    p_text
                )
                for tbl_id, cap_text in tbl_captions:
                    clean_id = re.sub(r"\s+", " ", tbl_id).title()
                    has_table_data = len(tables) > 0 or "|" in p_text
                    status = "PASS" if has_table_data else "WARN"
                    visual_assets.append({
                        "label": clean_id,
                        "caption_found": True,
                        "visual_image_present": has_table_data,
                        "placement_location": f"Page {page_idx}",
                        "status": status,
                        "notes": (
                            f"Table layout structure confirmed on Page {page_idx}."
                            if has_table_data else
                            f"Table caption detected without distinct cell grid on Page {page_idx}."
                        )
                    })

            full_text = "\n".join(page_text_list)
    else:
        raise ValueError("Unsupported format. Please upload .docx or .pdf.")

    return full_text, visual_assets, full_text.splitlines()


# ============================================================
# DETERMINISTIC & STRUCTURAL AUDITS
# ============================================================

def find_robust_sections(lines: List[str]) -> Dict[str, Dict[str, Any]]:
    """Detect section locations with numerical prefix stripping and empty-body checks."""
    found = {}
    for idx, line in enumerate(lines):
        raw = line.strip()
        cleaned = strip_numerical_prefix(raw)
        for canonical, aliases in SECTION_ALIASES.items():
            if cleaned in aliases and canonical not in found:
                # Lookahead to verify text actually follows
                next_non_empty = ""
                for next_line in lines[idx + 1: idx + 10]:
                    if next_line.strip():
                        next_non_empty = next_line.strip()
                        break
                found[canonical] = {
                    "line_index": idx,
                    "raw_heading": raw,
                    "first_line_quote": next_non_empty[:120],
                    "has_body": bool(next_non_empty),
                }
    return found


def extract_abstract(text: str) -> str:
    lines = text.splitlines()
    start = None
    for idx, line in enumerate(lines):
        cleaned = strip_numerical_prefix(line)
        if cleaned in ("abstract",):
            start = idx
            break

    if start is None:
        m = re.search(r"(?is)\babstract\s*:\s*(.*?)(?:\bkeywords\s*:|$)", text)
        return normalize(m.group(1)) if m else ""

    end_candidates = [
        i for i, line in enumerate(lines)
        if i > start and (
            strip_numerical_prefix(line).startswith("keywords") or
            strip_numerical_prefix(line) in ("introduction", "1. introduction")
        )
    ]
    end = min(end_candidates) if end_candidates else len(lines)
    val = "\n".join(lines[start + 1:end]).strip()
    if not val and ":" in lines[start]:
        val = lines[start].split(":", 1)[1]
    return normalize(val)


def extract_keywords(text: str) -> List[str]:
    m = re.search(
        r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:(?:[0-9]+\.?\s*)?introduction\b|\n\s*\n\s*[A-Z]|$))",
        text,
    )
    if not m:
        return []
    raw = " ".join([line.strip() for line in m.group(1).splitlines() if line.strip()])
    return [normalize(k) for k in re.split(r"[;,]", raw) if normalize(k)]


def audit_citations_and_references(text: str, lines: List[str], ref_index: Optional[int]) -> List[Dict[str, Any]]:
    """Audit square-bracket citations, truncated reference entries, and cross-references."""
    issues = []
    body_lines = lines[:ref_index] if ref_index is not None else lines
    body_text = "\n".join(body_lines)

    # 1. Inline citations
    inline_citations = re.findall(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]", body_text)
    if not inline_citations:
        issues.append({
            "issue_type": "Missing Inline Brackets",
            "evidence_quote": body_text[:140] + "...",
            "severity": "Major",
        })

    # 2. Reference list truncation audit
    if ref_index is not None:
        ref_text = "\n".join(lines[ref_index:])
        truncated = re.findall(
            r"(?i)(\[\d+\]\s*(?:Reference\s*\d+|et al\.?$|[a-zA-Z\s]{1,15}$))",
            ref_text
        )
        for t in truncated:
            issues.append({
                "issue_type": "Truncated Reference",
                "evidence_quote": t.strip(),
                "severity": "Major",
            })
    return issues


def run_pragmatic_rule_checks(text: str, visual_assets: List[Dict[str, Any]]) -> Dict[str, Any]:
    lines = text.splitlines()
    sections = find_robust_sections(lines)
    abstract = extract_abstract(text)
    keywords = extract_keywords(text)

    # 1. Structural Section Checks
    structural_checks = []
    for canonical in MRJ_RULES["required_sections"]:
        key = canonical.lower()
        sec_info = sections.get(key)
        if sec_info and sec_info["has_body"]:
            status = "PASS"
            notes = f"Section heading detected with body text."
        elif sec_info and not sec_info["has_body"]:
            status = "WARN"
            notes = f"Section heading detected but immediately followed by empty lines or EOF."
        else:
            status = "FAIL"
            notes = f"Section '{canonical}' not detected across full text layout."

        structural_checks.append({
            "section_name": canonical,
            "detected": bool(sec_info),
            "detected_heading_text": sec_info["raw_heading"] if sec_info else "None",
            "first_line_quote": sec_info["first_line_quote"] if sec_info else "None",
            "status": status,
        })

    # 2. Pragmatic Word Count Check
    citation_formatting_issues = []
    abstract_words = word_count(abstract)
    if abstract_words == 0:
        citation_formatting_issues.append({
            "issue_type": "Word Count Limit",
            "evidence_quote": "No abstract detected.",
            "severity": "Major"
        })
    elif abstract_words <= MRJ_RULES["abstract_max_words"]:
        pass
    elif abstract_words <= MRJ_RULES["abstract_soft_overflow_limit"]:
        citation_formatting_issues.append({
            "issue_type": "Word Count Limit",
            "evidence_quote": f"Abstract contains {abstract_words} words (guideline is {MRJ_RULES['abstract_max_words']}).",
            "severity": "Minor"
        })
    else:
        citation_formatting_issues.append({
            "issue_type": "Word Count Limit",
            "evidence_quote": f"Abstract contains {abstract_words} words (exceeds {MRJ_RULES['abstract_max_words']}).",
            "severity": "Major"
        })

    # 3. Citation issues
    ref_idx = sections.get("references", {}).get("line_index")
    citation_formatting_issues.extend(audit_citations_and_references(text, lines, ref_idx))

    return {
        "structural_section_checks": structural_checks,
        "visual_asset_checks": visual_assets,
        "citation_and_formatting_issues": citation_formatting_issues,
        "keywords": keywords,
        "abstract": abstract,
        "sections": sections,
    }


# ============================================================
# GROQ STRUCTURED METHODOLOGY & MATH AUDIT
# ============================================================

EVALUATION_SCHEMA = {
    "name": "manuscript_pre_screen_evaluation",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "manuscript_title": {"type": "string"},
            "methodology_and_math_logic": {
                "type": "object",
                "properties": {
                    "sample_size_check": {"type": "string", "enum": ["PASS", "WARN"]},
                    "math_discrepancies": {"type": "array", "items": {"type": "string"}},
                    "missing_domain_metrics": {"type": "array", "items": {"type": "string"}},
                    "software_parameter_notes": {"type": "string"},
                },
                "required": [
                    "sample_size_check",
                    "math_discrepancies",
                    "missing_domain_metrics",
                    "software_parameter_notes",
                ],
                "additionalProperties": False,
            },
            "research_area": {"type": "string"},
            "reviewer_search_keywords": {"type": "array", "items": {"type": "string"}},
            "editorial_recommendation": {
                "type": "object",
                "properties": {
                    "verdict": {
                        "type": "string",
                        "enum": [
                            "Accept as is",
                            "Accept with Minor Revisions",
                            "Major Revisions",
                            "Reject",
                        ],
                    },
                    "key_reasons": {"type": "array", "items": {"type": "string"}},
                },
                "required": ["verdict", "key_reasons"],
                "additionalProperties": False,
            },
        },
        "required": [
            "manuscript_title",
            "methodology_and_math_logic",
            "research_area",
            "reviewer_search_keywords",
            "editorial_recommendation",
        ],
        "additionalProperties": False,
    },
}


def run_groq_math_and_logic_audit(
    text: str,
    pre_checks: Dict[str, Any],
    client: Groq,
    model: str = DEFAULT_GROQ_MODEL,
) -> Dict[str, Any]:
    """Audit arithmetic subtotals, bibliometric parameters, and software versions via Groq."""
    intro = text[:4000]
    methods = ""
    res_disc = ""

    lines = text.splitlines()
    sections = pre_checks["sections"]
    if "materials and methods" in sections:
        start = sections["materials and methods"]["line_index"]
        methods = "\n".join(lines[start: start + 120])
    if "results and discussion" in sections:
        start = sections["results and discussion"]["line_index"]
        res_disc = "\n".join(lines[start: start + 120])

    prompt = f"""AUDIT DIRECTIVES FOR ACADEMIC PRE-SCREENING:
1. MATHEMATICAL & SAMPLE SIZE LOGIC: Audit entity sample sizes (N). Check if country counts, category distributions, or table subtotals sum to > N without clarifying fractional/full counting conventions.
2. LOW-FREQUENCY NOISE & DOMAIN METRICS: If bibliometric or empirical data is used, check for missing parameters (h-index, g-index, m-index, software versions e.g., VOSviewer, Bibliometrix, R packages, normalizations).
3. PRAGMATIC THRESHOLDS: Reserve 'Major Revisions'/'Reject' strictly for missing body text, unverified data logic, or dropped visual assets. Use 'Accept with Minor Revisions' for minor word limits or formatting overflows.

EXTRACTS FROM MANUSCRIPT:
Title/Intro region:
{intro[:2000]}

Materials & Methods:
{methods[:3000]}

Results / Data Excerpts:
{res_disc[:3000]}

Respond ONLY in the structured JSON schema provided."""

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert scientific manuscript pre-screener. Audit mathematical logic and domain standards without hallucination.",
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.1,
            max_completion_tokens=2200,
            response_format={"type": "json_schema", "json_schema": EVALUATION_SCHEMA},
        )
        return json.loads(response.choices[0].message.content or "{}")
    except Exception as exc:
        return {
            "manuscript_title": "Undetermined Title",
            "methodology_and_math_logic": {
                "sample_size_check": "WARN",
                "math_discrepancies": [f"AI evaluation could not be completed: {exc}"],
                "missing_domain_metrics": [],
                "software_parameter_notes": "Manual inspection required.",
            },
            "research_area": "General Multidisciplinary",
            "reviewer_search_keywords": pre_checks.get("keywords", [])[:5],
            "editorial_recommendation": {
                "verdict": "Accept with Minor Revisions",
                "key_reasons": ["Automated mathematical review encountered a network or token issue."],
            },
        }


# ============================================================
# LIVE OPENALEX REVIEWER HARVESTING (CACHED)
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    res = requests.get(url, params=params, timeout=20)
    res.raise_for_status()
    return res.json()


@lru_cache(maxsize=256)
def get_openalex_author(author_id: str, mailto: str = "") -> Dict[str, Any]:
    try:
        return openalex_get(f"{OPENALEX_AUTHOR_URL}/{author_id}", {}, mailto)
    except Exception:
        return {}


def candidate_matches_region(c: Dict[str, Any], region: str) -> bool:
    meta = f"{c.get('institution','')} {c.get('city','')} {c.get('region','')} {c.get('country','')}".lower()
    if region == "India":
        return c.get("country", "").lower() == "in" or "india" in meta
    if region == "Northeast India":
        return any(t in meta for t in (set(NORTHEAST_STATES) | set(NORTHEAST_INSTITUTIONS)))
    if region == "Assam":
        return any(t in meta for t in ASSAM_TERMS)
    return False


def harvest_openalex_reviewers(
    keywords: List[str],
    research_area: str,
    mailto: str = "",
) -> Dict[str, List[Dict[str, Any]]]:
    clean_terms = [normalize(k) for k in (keywords[:4] + [research_area]) if normalize(k)]
    query = " ".join(clean_terms[:5]).strip()
    out = {"India": [], "Northeast India": [], "Assam": []}
    if not query:
        return out

    try:
        data = openalex_get(
            OPENALEX_URL,
            {"search": query, "per-page": 40, "sort": "publication_year:desc"},
            mailto,
        )
    except Exception as exc:
        err = [{"error": f"OpenAlex query error: {exc}"}]
        return {k: err for k in out}

    candidates: Dict[str, Dict[str, Any]] = {}
    for work in data.get("results", []):
        year = work.get("publication_year") or 0
        title = work.get("display_name") or ""
        doi = work.get("doi") or ""
        cited = work.get("cited_by_count") or 0

        for authorship in work.get("authorships", []):
            author = authorship.get("author") or {}
            aid = author.get("id")
            aname = author.get("display_name")
            if not aid or not aname:
                continue

            insts = authorship.get("institutions") or []
            if not insts:
                continue

            inst = insts[0]
            if aid not in candidates:
                candidates[aid] = {
                    "id": aid,
                    "name": aname,
                    "institution": inst.get("display_name") or "",
                    "country": inst.get("country_code") or "",
                    "city": (inst.get("geo") or {}).get("city") or "",
                    "region": (inst.get("geo") or {}).get("region") or "",
                    "publications": [],
                    "score": 0.0,
                }
            candidates[aid]["publications"].append({"title": title, "year": year, "doi": doi})
            candidates[aid]["score"] += max(0, year - 2018) * 0.8 + min(cited, 100) * 0.03

    # Resolve last-known affiliations for top 20 candidates
    pool = sorted(candidates.values(), key=lambda x: x["score"], reverse=True)[:20]
    for c in pool:
        auth_info = get_openalex_author(c["id"], mailto)
        lk = (auth_info.get("last_known_institutions") or [{}])[0]
        geo = lk.get("geo") or {}
        c["last_institution"] = lk.get("display_name") or c["institution"]
        c["last_country"] = lk.get("country_code") or c["country"]
        c["last_city"] = geo.get("city") or c["city"]
        c["last_region"] = geo.get("region") or c["region"]

    for reg in ["Assam", "Northeast India", "India"]:
        matched = [
            c for c in pool
            if candidate_matches_region({
                "institution": c["last_institution"],
                "country": c["last_country"],
                "city": c["last_city"],
                "region": c["last_region"],
            }, reg)
        ]
        matched.sort(key=lambda x: (len(x["publications"]), x["score"]), reverse=True)
        out[reg] = matched[:5]

    return out


# ============================================================
# BLIND REVIEW COPY GENERATOR
# ============================================================

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ORCID_RE = re.compile(r"\b(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}\b", re.I)

def delete_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    if p is not None and p.getparent() is not None:
        p.getparent().remove(p)


def blind_copy_docx(original_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(original_bytes))
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""

    paras = list(doc.paragraphs)
    abstract_idx = None
    for idx, p in enumerate(paras):
        if re.match(r"(?i)^\s*(?:[0-9]+\.?\s*)?abstract\b", p.text.strip()):
            abstract_idx = idx
            break

    # Strip author blocks before Abstract while retaining title
    if abstract_idx is not None:
        nonempty = [i for i, p in enumerate(paras[:abstract_idx]) if p.text.strip()]
        if nonempty:
            title_idx = nonempty[0]
            for i in nonempty[1:]:
                delete_paragraph(paras[i])

    # Redact sensitive sections and emails
    removable = {"funding", "acknowledgments", "acknowledgements", "author contributions", "competing interests"}
    for p in list(doc.paragraphs):
        txt = normalize(p.text).lower()
        if any(txt.startswith(r) for r in removable):
            delete_paragraph(p)
        else:
            for run in p.runs:
                if EMAIL_RE.search(run.text) or ORCID_RE.search(run.text):
                    run.text = EMAIL_RE.sub("[REDACTED EMAIL]", run.text)
                    run.text = ORCID_RE.sub("[REDACTED ORCID]", run.text)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# DOCX REPORT COMPILER
# ============================================================

def compile_audit_docx_report(audit_json: Dict[str, Any]) -> bytes:
    doc = Document()
    doc.add_heading("Academic Manuscript Editorial Pre-Screening Audit", 0)
    doc.add_paragraph(f"Manuscript Title: {audit_json.get('manuscript_title', 'Not specified')}")

    # Editorial Recommendation
    rec = audit_json.get("editorial_recommendation", {})
    doc.add_heading(f"Verdict: {rec.get('verdict', 'Under Review')}", level=1)
    for r in rec.get("key_reasons", []):
        doc.add_paragraph(r, style="List Bullet")

    # Section 1: Structural Audit
    doc.add_heading("1. Structural Section Checks", level=2)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Section", "Detected Heading", "First Line Quote", "Status"
    for s in audit_json.get("structural_section_checks", []):
        row = tbl.add_row().cells
        row[0].text = s["section_name"]
        row[1].text = s["detected_heading_text"]
        row[2].text = s["first_line_quote"][:80]
        row[3].text = s["status"]

    # Section 2: Visual Asset Audit
    doc.add_heading("2. Visual & Layout Asset Audit", level=2)
    v_assets = audit_json.get("visual_asset_checks", [])
    if v_assets:
        vtbl = doc.add_table(rows=1, cols=4)
        vtbl.style = "Table Grid"
        vhdr = vtbl.rows[0].cells
        vhdr[0].text, vhdr[1].text, vhdr[2].text, vhdr[3].text = "Label", "Placement", "Visual Present", "Status"
        for v in v_assets:
            row = vtbl.add_row().cells
            row[0].text = v["label"]
            row[1].text = v["placement_location"]
            row[2].text = "Yes" if v["visual_image_present"] else "No"
            row[3].text = v["status"]
    else:
        doc.add_paragraph("No explicit Figure or Table captions detected.")

    # Section 3: Methodology & Math Logic
    doc.add_heading("3. Methodological & Sample-Size Logic", level=2)
    math_info = audit_json.get("methodology_and_math_logic", {})
    doc.add_paragraph(f"Sample Size Verification: {math_info.get('sample_size_check', 'PASS')}")
    for d in math_info.get("math_discrepancies", []):
        doc.add_paragraph(f"- Discrepancy: {d}")
    for m in math_info.get("missing_domain_metrics", []):
        doc.add_paragraph(f"- Missing Domain Parameter: {m}")
    doc.add_paragraph(f"Software / Parameters: {math_info.get('software_parameter_notes', 'None recorded')}")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(
    page_title="Academic Manuscript Pre-Screening Engine",
    page_icon="🔬",
    layout="wide"
)

with st.sidebar:
    st.header("⚙️ Editorial Settings")
    groq_key = st.text_input(
        "Groq API Key",
        value=get_secret("GROQ_API_KEY"),
        type="password"
    )
    openalex_mail = st.text_input(
        "OpenAlex Polite Email",
        value=get_secret("OPENALEX_MAILTO", "editor@scholarly.org")
    )
    model_choice = st.selectbox(
        "Groq Reasoning Model",
        options=[DEFAULT_GROQ_MODEL, "llama-3.3-70b-versatile"],
        index=0
    )
    st.info("System evaluates section presence, cross-modal figure/image parity, arithmetic logic, and OpenAlex reviewers.")

st.title("🔬 Scientific Manuscript Pre-Screening Assistant")
st.caption("Evidence-based pre-screening adhering to visual layout audits, sample-size logic, and pragmatic thresholding.")

uploaded_file = st.file_uploader("Upload Academic Manuscript (PDF or DOCX)", type=["pdf", "docx"])

if uploaded_file:
    if st.button("Run Comprehensive Manuscript Audit", type="primary"):
        if not groq_key:
            st.error("Please supply a Groq API Key.")
            st.stop()

        with st.status("Performing cross-modal audit...", expanded=True) as status:
            st.write("Extracting layout & visual streams (images, vector drawings, tables)...")
            raw_text, visual_assets, lines = extract_document_assets(uploaded_file)

            st.write("Verifying numerical section prefixes and reference citations...")
            pre_checks = run_pragmatic_rule_checks(raw_text, visual_assets)

            st.write("Auditing mathematical logic, counting conventions, and domain metrics...")
            groq_client = Groq(api_key=groq_key)
            ai_eval = run_groq_math_and_logic_audit(raw_text, pre_checks, groq_client, model=model_choice)

            # Consolidate output adhering to the exact required schema
            final_audit = {
                "manuscript_title": ai_eval.get("manuscript_title", "Undetermined"),
                "structural_section_checks": pre_checks["structural_section_checks"],
                "visual_asset_checks": visual_assets,
                "methodology_and_math_logic": ai_eval.get("methodology_and_math_logic", {}),
                "citation_and_formatting_issues": pre_checks["citation_and_formatting_issues"],
                "editorial_recommendation": ai_eval.get("editorial_recommendation", {
                    "verdict": "Accept with Minor Revisions",
                    "key_reasons": ["Automated pre-screening complete."]
                })
            }

            st.write("Harvesting live reviewer candidates from OpenAlex...")
            rev_candidates = harvest_openalex_reviewers(
                ai_eval.get("reviewer_search_keywords", []),
                ai_eval.get("research_area", "General"),
                openalex_mail
            )

            # Generate blind copy
            blind_bytes = None
            if uploaded_file.name.lower().endswith(".docx"):
                blind_bytes = blind_copy_docx(uploaded_file.getvalue())

            report_doc = compile_audit_docx_report(final_audit)

            st.session_state["audit_result"] = final_audit
            st.session_state["reviewers"] = rev_candidates
            st.session_state["report_doc"] = report_doc
            st.session_state["blind_docx"] = blind_bytes
            status.update(label="Manuscript pre-screening completed!", state="complete")

if "audit_result" in st.session_state:
    res = st.session_state["audit_result"]
    verdict = res["editorial_recommendation"]["verdict"]

    st.divider()
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader(f"📄 {res['manuscript_title']}")
    with col2:
        badge_color = "green" if "Accept" in verdict else "orange" if "Minor" in verdict else "red"
        st.markdown(f"### Verdict: :{badge_color}[{verdict}]")

    st.write("**Key Decision Reasons:**")
    for r in res["editorial_recommendation"]["key_reasons"]:
        st.write(f"- {r}")

    # Tabs for Audited Directives
    t1, t2, t3, t4, t5 = st.tabs([
        "1. Sections Audit",
        "2. Visual & Layout Assets",
        "3. Math & Methodology Logic",
        "4. Citations & Ethics",
        "5. Reviewers & Downloads"
    ])

    with t1:
        st.write("#### Numerical Prefix & Absence Verification")
        st.dataframe(
            [
                {
                    "Section": s["section_name"],
                    "Detected Heading": s["detected_heading_text"],
                    "First Line Quote": s["first_line_quote"],
                    "Status": s["status"]
                }
                for s in res["structural_section_checks"]
            ],
            use_container_width=True,
            hide_index=True
        )

    with t2:
        st.write("#### Cross-Modal Visual Parity (Captions vs Rendered Assets)")
        if res["visual_asset_checks"]:
            st.dataframe(
                [
                    {
                        "Asset": v["label"],
                        "Page / Location": v["placement_location"],
                        "Caption Found": v["caption_found"],
                        "Rendered Visual Present": v["visual_image_present"],
                        "Status": v["status"],
                        "Audit Notes": v["notes"]
                    }
                    for v in res["visual_asset_checks"]
                ],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No Figure or Table captions identified in layout.")

    with t3:
        st.write("#### Mathematical Consistency & Domain Indices")
        m = res["methodology_and_math_logic"]
        st.metric("Sample Size Verification", m.get("sample_size_check", "PASS"))
        if m.get("math_discrepancies"):
            st.warning("Sample Size / Subtotal Inconsistencies:")
            for d in m["math_discrepancies"]:
                st.write(f"- {d}")
        if m.get("missing_domain_metrics"):
            st.info("Unreported Domain Metrics (e.g. h/g-index, parameter versions):")
            for met in m["missing_domain_metrics"]:
                st.write(f"- {met}")
        st.write("**Software & Methodological Parameters:**", m.get("software_parameter_notes", "None reported"))

    with t4:
        st.write("#### Citation Integrity & Pragmatic Thresholds")
        if res["citation_and_formatting_issues"]:
            st.dataframe(
                [
                    {
                        "Issue Type": c["issue_type"],
                        "Severity": c["severity"],
                        "Evidence Quote": c["evidence_quote"]
                    }
                    for c in res["citation_and_formatting_issues"]
                ],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("No truncated references or citation formatting penalties detected.")

    with t5:
        st.write("#### OpenAlex Live Reviewer Pool")
        revs = st.session_state.get("reviewers", {})
        for region, candidates in revs.items():
            st.markdown(f"**{region} Candidates:**")
            if not candidates or "error" in candidates[0]:
                st.caption("No regional candidates found.")
                continue
            st.dataframe(
                [
                    {
                        "Name": c["name"],
                        "Affiliation": f"{c.get('last_institution','')} ({c.get('last_city','')}, {c.get('last_country','')})",
                        "Recent Work": c.get("publications", [{}])[0].get("title", "")
                    }
                    for c in candidates
                ],
                use_container_width=True,
                hide_index=True
            )

        st.divider()
        st.write("#### Downloads")
        d1, d2 = st.columns(2)
        d1.download_button(
            "📥 Download Audit Report (.docx)",
            data=st.session_state["report_doc"],
            file_name="Manuscript_Audit_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        if st.session_state.get("blind_docx"):
            d2.download_button(
                "📥 Download Blind Reviewer Copy (.docx)",
                data=st.session_state["blind_docx"],
                file_name="Blind_Reviewer_Copy.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
