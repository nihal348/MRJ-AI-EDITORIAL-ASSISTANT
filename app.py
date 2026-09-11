import io
import json
import os
import re
from copy import deepcopy
from typing import Any, Dict, List, Tuple

import pdfplumber
import requests
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from groq import Groq


# ============================================================
# CONFIGURATION
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
OPENALEX_URL = "https://api.openalex.org/works"
OPENALEX_AUTHOR_URL = "https://api.openalex.org/authors"

MRJ_RULES = {
    "abstract_max_words": 200,
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
    "reference_style": "Numbered references in order of appearance using square-bracket citations.",
    "doi_required_where_available": True,
}

SECTION_ALIASES = {
    "introduction": ["introduction", "1. introduction"],
    "materials and methods": [
        "materials and methods",
        "materials & methods",
        "methods",
        "2. materials and methods",
    ],
    "results and discussion": [
        "results and discussion",
        "results & discussion",
        "3. results and discussion",
        "results",
        "discussion",
    ],
    "conclusions": ["conclusions", "conclusion", "4. conclusions"],
    "multidisciplinary domains": ["multidisciplinary domains"],
    "funding": ["funding"],
    "acknowledgments": ["acknowledgments", "acknowledgements"],
    "conflicts of interest": ["conflicts of interest", "conflict of interest"],
    "declaration on ai usage": [
        "declaration on ai usage",
        "ai usage",
        "artificial intelligence",
    ],
    "references": ["references", "reference"],
}

NORTHEAST_STATES = {
    "assam",
    "arunachal pradesh",
    "manipur",
    "meghalaya",
    "mizoram",
    "nagaland",
    "sikkim",
    "tripura",
}

NORTHEAST_INSTITUTION_TERMS = [
    "iit guwahati",
    "tezu university",
    "tezu",
    "nit silchar",
    "assam university",
    "gauhati university",
    "cotton university",
    "dibrugarh university",
    "nehu",
    "north-eastern hill university",
    "niser",
    "nit agartala",
    "central university of jharkhand",
    "manipur university",
    "mizoram university",
    "nagaland university",
    "tripura university",
    "rajiv gandhi university",
    "arunachal university",
    "sikkim university",
]

ASSAM_TERMS = [
    "assam",
    "iit guwahati",
    "gauhati university",
    "cotton university",
    "dibrugarh university",
    "assam university",
    "tezpur university",
    "nit silchar",
    "indian institute of technology guwahati",
]


# ============================================================
# GENERAL HELPERS
# ============================================================

def get_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def word_count(text: str) -> int:
    return len(re.findall(r"\b[\w'-]+\b", text or ""))


def first_nonempty_lines(text: str, limit: int = 20) -> List[str]:
    return [x.strip() for x in (text or "").splitlines() if x.strip()][:limit]


def find_section_positions(lines: List[str]) -> Dict[str, int]:
    positions = {}
    for i, line in enumerate(lines):
        n = normalize(line).lower().rstrip(":")
        for canonical, aliases in SECTION_ALIASES.items():
            if n in aliases and canonical not in positions:
                positions[canonical] = i
    return positions


def extract_section_text(text: str, canonical: str) -> str:
    lines = text.splitlines()
    positions = find_section_positions(lines)
    start = positions.get(canonical)
    if start is None:
        return ""
    next_positions = [p for p in positions.values() if p > start]
    end = min(next_positions) if next_positions else len(lines)
    return "\n".join(lines[start + 1:end]).strip()


def extract_abstract(text: str) -> str:
    lines = text.splitlines()
    positions = find_section_positions(lines)
    start = None
    for key in ["abstract"]:
        for i, line in enumerate(lines):
            if normalize(line).lower().rstrip(":") == key:
                start = i
                break
    if start is None:
        # Fall back to text between an "Abstract:" marker and Keywords.
        m = re.search(r"(?is)\babstract\s*:\s*(.*?)(?:\bkeywords\s*:|$)", text)
        return normalize(m.group(1)) if m else ""
    end_candidates = [
        i for i, line in enumerate(lines)
        if i > start and normalize(line).lower().startswith("keywords")
    ]
    end = min(end_candidates) if end_candidates else len(lines)
    value = "\n".join(lines[start + 1:end]).strip()
    if not value and ":" in lines[start]:
        value = lines[start].split(":", 1)[1]
    return normalize(value)


def extract_keywords(text: str) -> List[str]:
    m = re.search(
        r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:1\.?\s+)?introduction\b|\n\s*abstract\b|$)",
        text,
    )
    if not m:
        return []
    raw = m.group(1).strip().splitlines()[0]
    return [normalize(x) for x in re.split(r"[;,]", raw) if normalize(x)]


def extract_domain_statement(text: str) -> str:
    section = extract_section_text(text, "multidisciplinary domains")
    if not section:
        m = re.search(
            r"(?is)this research covers the domains\s*:\s*(.*?)(?:\n\s*(?:funding|acknowledg|conflicts|declaration|references)\b|$)",
            text,
        )
        return normalize(m.group(0)) if m else ""
    return normalize(section)


def extract_author_block(text: str) -> str:
    """Return the front-matter region between title and abstract.

    Used only for COI/blind-review processing. It is never sent to the
    reviewer-search prompt as an instruction to invent names.
    """
    lines = [x for x in text.splitlines()]
    abstract_idx = None
    for i, line in enumerate(lines):
        if re.match(r"(?i)^\s*abstract\s*:?\s*$", line) or re.match(
            r"(?i)^\s*abstract\s*:", line
        ):
            abstract_idx = i
            break
    if abstract_idx is None:
        return ""
    return "\n".join(lines[:abstract_idx])


# ============================================================
# FILE EXTRACTION
# ============================================================

def extract_text(uploaded_file) -> str:
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = []

        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text)

        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))

        return "\n".join(parts)

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text() or "")
        return "\n".join(pages)

    raise ValueError("Unsupported file type.")


# ============================================================
# DETERMINISTIC MRJ PRE-SCREENING
# ============================================================

def run_mrj_rule_checks(text: str) -> List[Dict[str, Any]]:
    checks = []
    lower = text.lower()

    abstract = extract_abstract(text)
    keywords = extract_keywords(text)
    positions = find_section_positions(text.splitlines())

    # Abstract
    if not abstract:
        checks.append({
            "requirement": "Abstract",
            "status": "FAIL",
            "evidence": "No abstract was detected.",
            "action": "Add an abstract following the MRJ template.",
        })
    else:
        n = word_count(abstract)
        status = "PASS" if n <= MRJ_RULES["abstract_max_words"] else "FAIL"
        checks.append({
            "requirement": "Abstract ≤ 200 words",
            "status": status,
            "evidence": f"Detected approximately {n} words.",
            "action": "Keep the abstract at or below 200 words." if status == "FAIL" else "No action required.",
        })

    # Keywords
    if MRJ_RULES["keywords_min"] <= len(keywords) <= MRJ_RULES["keywords_max"]:
        keyword_status = "PASS"
        keyword_action = "No action required."
    elif len(keywords) == 0:
        keyword_status = "FAIL"
        keyword_action = "Add 3–10 pertinent keywords."
    else:
        keyword_status = "FAIL"
        keyword_action = "Use 3–10 pertinent keywords."
    checks.append({
        "requirement": "3–10 keywords",
        "status": keyword_status,
        "evidence": f"Detected {len(keywords)} keyword(s): {', '.join(keywords) if keywords else 'none detected'}.",
        "action": keyword_action,
    })

    # Required sections
    required_map = [
        ("Introduction", "introduction"),
        ("Materials and Methods", "materials and methods"),
        ("Results and Discussion", "results and discussion"),
        ("Conclusions", "conclusions"),
        ("Multidisciplinary Domains", "multidisciplinary domains"),
        ("Funding", "funding"),
        ("Acknowledgments", "acknowledgments"),
        ("Conflicts of Interest", "conflicts of interest"),
        ("Declaration on AI Usage", "declaration on ai usage"),
        ("References", "references"),
    ]

    for label, canonical in required_map:
        exists = canonical in positions
        checks.append({
            "requirement": f"Required section: {label}",
            "status": "PASS" if exists else "FAIL",
            "evidence": "Section heading detected." if exists else "Section heading not detected.",
            "action": "No action required." if exists else f"Add the '{label}' section required by the MRJ template.",
        })

    # Domains
    domain_statement = extract_domain_statement(text)
    domain_matches = re.findall(r"\([a-z]\)\s*([^,;.]+)", domain_statement, flags=re.I)
    domain_count = len(domain_matches)
    if domain_count >= MRJ_RULES["minimum_domains"]:
        domain_status = "PASS"
        domain_action = "No action required."
    else:
        domain_status = "WARN" if domain_statement else "FAIL"
        domain_action = "State at least two research domains using the MRJ domain statement."

    checks.append({
        "requirement": "At least two multidisciplinary domains",
        "status": domain_status,
        "evidence": f"Detected {domain_count} explicit domain item(s)." if domain_statement else "No MRJ domain statement detected.",
        "action": domain_action,
    })

    # Citation style
    body_before_refs = text[: text.lower().find("references")] if "references" in lower else text
    numbered_citations = re.findall(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]", body_before_refs)
    citation_status = "PASS" if numbered_citations else "WARN"
    checks.append({
        "requirement": "Numbered square-bracket citations",
        "status": citation_status,
        "evidence": f"Detected {len(numbered_citations)} square-bracket citation pattern(s).",
        "action": "Check that references are numbered in order of appearance and citations use square brackets.",
    })

    # Ethical approval
    methods = extract_section_text(text, "materials and methods").lower()
    ethical_terms = [
        "ethics approval",
        "ethical approval",
        "ethics committee",
        "institutional review board",
        "irb",
        "approval code",
        "ethical clearance",
    ]
    ethics_relevant = any(
        term in (methods + " " + lower) for term in
        ["human", "patient", "participant", "animal", "clinical", "intervention"]
    )
    if ethics_relevant:
        ethics_found = any(term in methods for term in ethical_terms)
        checks.append({
            "requirement": "Ethics approval where applicable",
            "status": "PASS" if ethics_found else "WARN",
            "evidence": "Potential human/animal/intervention study language detected; ethics wording was " +
                        ("detected." if ethics_found else "not detected in Materials and Methods."),
            "action": "Verify that the approving authority and approval code are stated where required.",
        })

    # DOI presence in references
    references = extract_section_text(text, "references")
    doi_count = len(re.findall(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", references, flags=re.I))
    checks.append({
        "requirement": "DOI included for references where available",
        "status": "PASS" if doi_count else "WARN",
        "evidence": f"Detected {doi_count} DOI-like reference string(s).",
        "action": "Verify DOI information for references where a DOI exists.",
    })

    return checks


# ============================================================
# GROQ STRUCTURED ANALYSIS
# ============================================================

AI_SCHEMA = {
    "name": "mrj_editorial_assessment",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "summary": {"type": "string"},
            "research_area": {"type": "string"},
            "keywords": {
                "type": "array",
                "items": {"type": "string"},
            },
            "methodology_assessment": {"type": "string"},
            "results_discussion_assessment": {"type": "string"},
            "abstract_alignment": {"type": "string"},
            "potential_major_concerns": {
                "type": "array",
                "items": {"type": "string"},
            },
            "editorial_recommendation": {
                "type": "string",
                "enum": ["Proceed to editorial review", "Needs author correction", "Major concern"],
            },
        },
        "required": [
            "summary",
            "research_area",
            "keywords",
            "methodology_assessment",
            "results_discussion_assessment",
            "abstract_alignment",
            "potential_major_concerns",
            "editorial_recommendation",
        ],
        "additionalProperties": False,
    },
}


def run_ai_analysis(
    text: str,
    deterministic_checks: List[Dict[str, Any]],
    client: Groq,
) -> Dict[str, Any]:
    """
    Run the qualitative MRJ assessment with a deliberately small prompt.

    Groq's on-demand tier can impose a tokens-per-minute limit. The previous
    implementation sent the same manuscript information several times
    (section extracts + a 30,000-character full excerpt), which unnecessarily
    inflated the request. This version sends only the information needed for
    the qualitative assessment.
    """
    abstract = extract_abstract(text)
    methods = extract_section_text(text, "materials and methods")
    results = extract_section_text(text, "results and discussion")
    conclusions = extract_section_text(text, "conclusions")

    # Deterministic checks are already performed in Python. Give the model only
    # their compact status/evidence summary; do not resend the manuscript.
    rule_summary = "\n".join(
        f"- {c['requirement']}: {c['status']} — {c['evidence']}"
        for c in deterministic_checks
    )[:7000]

    # Hard character caps keep the request comfortably below the 8,000 TPM
    # limit on the user's current Groq on-demand tier.
    prompt = f"""
You are a cautious academic journal editorial pre-screening assistant for MRJ.

Rules:
- Assess only supplied evidence.
- Never invent authors, reviewers, institutions, citations, results, or facts.
- Do not rewrite the manuscript.
- Deterministic checks below are authoritative for formatting/compliance.
- If evidence is insufficient, say so.
- Do not make an acceptance/rejection decision.
- Return only the requested structured assessment.

DETERMINISTIC MRJ CHECKS:
{rule_summary}

ABSTRACT:
{abstract[:3500]}

MATERIALS AND METHODS:
{methods[:6500]}

RESULTS AND DISCUSSION:
{results[:7000]}

CONCLUSIONS:
{conclusions[:3000]}
"""

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a cautious journal-editor assistant. "
                        "Return only the requested structured assessment."
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            reasoning_effort="medium",
            max_tokens=900,
            response_format={
                "type": "json_schema",
                "json_schema": AI_SCHEMA,
            },
        )

        content = response.choices[0].message.content or ""
        return json.loads(content)

    except Exception as exc:
        # Keep the application usable when Groq throttles the request.
        # The deterministic MRJ checks remain available and are not discarded.
        error_text = str(exc)
        if "413" in error_text or "tokens per minute" in error_text.lower():
            return {
                "summary": (
                    "The deterministic MRJ checks were completed, but the "
                    "qualitative Groq assessment was skipped because the "
                    "current Groq tokens-per-minute limit was exceeded."
                ),
                "research_area": "Not assessed",
                "keywords": [],
                "methodology_assessment": "Not assessed because the AI request was rate-limited.",
                "results_discussion_assessment": "Not assessed because the AI request was rate-limited.",
                "abstract_alignment": "Not assessed because the AI request was rate-limited.",
                "potential_major_concerns": [
                    "Groq request exceeded the current tokens-per-minute limit. "
                    "The manuscript was not modified."
                ],
                "editorial_recommendation": "Needs author correction",
            }

        raise


# ============================================================
# REVIEWER SEARCH: OPENALEX
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    response = requests.get(url, params=params, timeout=30)
    response.raise_for_status()
    return response.json()


def get_openalex_author(author_id: str, mailto: str = "") -> Dict[str, Any]:
    try:
        return openalex_get(f"{OPENALEX_AUTHOR_URL}/{author_id}", {}, mailto)
    except Exception:
        return {}


def institution_location_text(inst: Dict[str, Any]) -> str:
    if not inst:
        return ""
    country = (inst.get("country_code") or "").lower()
    display = inst.get("display_name") or ""
    geo = inst.get("geo") or {}
    return " ".join(
        str(x or "") for x in [
            display,
            country,
            geo.get("city"),
            geo.get("region"),
        ]
    ).lower()


def candidate_region_match(candidate: Dict[str, Any], region: str) -> bool:
    text = " ".join(
        [
            candidate.get("institution", ""),
            candidate.get("city", ""),
            candidate.get("region", ""),
            candidate.get("country", ""),
        ]
    ).lower()

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
    max_candidates: int = 6,
) -> List[Dict[str, Any]]:
    query = " ".join(search_terms[:8]).strip()
    if not query:
        return []

    data = openalex_get(
        OPENALEX_URL,
        {
            "search": query,
            "per-page": 40,
            "sort": "publication_year:desc",
        },
        mailto,
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
                    "city": ((inst.get("geo") or {}).get("city") or ""),
                    "region": ((inst.get("geo") or {}).get("region") or ""),
                    "recent_publications": [],
                    "score": 0.0,
                }

                if not candidate_region_match(candidate, region):
                    continue

                key = author_id
                if key not in candidates:
                    candidates[key] = candidate

                rec = candidates[key]
                rec["recent_publications"].append({
                    "title": title,
                    "year": year,
                    "doi": doi,
                    "cited_by_count": cited,
                })

                # Recent publication + citation evidence, not fabricated expertise.
                rec["score"] += max(0, year - 2018) * 0.8
                rec["score"] += min(cited, 100) * 0.03

    results = list(candidates.values())

    # Try to use the author's last-known institution as an additional verification signal.
    for candidate in results[:30]:
        author = get_openalex_author(candidate["author_id"], mailto)
        last_known = author.get("last_known_institutions") or []
        if last_known:
            lk = last_known[0]
            candidate["last_known_institution"] = lk.get("display_name") or candidate["institution"]
            candidate["last_known_country"] = lk.get("country_code") or candidate["country"]
            candidate["last_known_city"] = ((lk.get("geo") or {}).get("city") or candidate["city"])
            candidate["last_known_region"] = ((lk.get("geo") or {}).get("region") or candidate["region"])
        else:
            candidate["last_known_institution"] = candidate["institution"]
            candidate["last_known_country"] = candidate["country"]
            candidate["last_known_city"] = candidate["city"]
            candidate["last_known_region"] = candidate["region"]

    # Re-check region against last-known institution where possible.
    verified = []
    for candidate in results:
        check_candidate = dict(candidate)
        check_candidate["institution"] = candidate["last_known_institution"]
        check_candidate["country"] = candidate["last_known_country"]
        check_candidate["city"] = candidate["last_known_city"]
        check_candidate["region"] = candidate["last_known_region"]

        if candidate_region_match(check_candidate, region):
            candidate["verification_note"] = (
                "OpenAlex publication affiliation plus last-known institution match."
            )
            verified.append(candidate)
        else:
            # Keep the publication evidence but make the uncertainty explicit.
            candidate["verification_note"] = (
                "Publication-associated affiliation matched the requested region; "
                "current affiliation could not be independently confirmed."
            )
            if region == "India" and candidate["country"] == "in":
                verified.append(candidate)

    verified.sort(
        key=lambda x: (
            len(x["recent_publications"]),
            x["score"],
            max((p["year"] for p in x["recent_publications"]), default=0),
        ),
        reverse=True,
    )

    return verified[:max_candidates]


def reviewer_search_report(
    ai_data: Dict[str, Any],
    mailto: str = "",
) -> Dict[str, List[Dict[str, Any]]]:
    terms = ai_data.get("keywords", []) + [ai_data.get("research_area", "")]
    terms = [normalize(x) for x in terms if normalize(x)]

    output = {}
    for region in ["India", "Northeast India", "Assam"]:
        try:
            output[region] = search_openalex_reviewers(
                terms,
                region,
                mailto=mailto,
                max_candidates=6,
            )
        except Exception as exc:
            output[region] = [{
                "error": f"OpenAlex search failed: {exc}"
            }]
    return output


# ============================================================
# BLIND REVIEW COPY
# ============================================================

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ORCID_RE = re.compile(r"\b(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}\b", re.I)
PHONE_RE = re.compile(r"(?<!\d)(?:\+?\d[\d\s().-]{7,}\d)(?!\d)")
URL_RE = re.compile(r"https?://\S+", re.I)

IDENTIFYING_LABELS = [
    "correspondence",
    "corresponding author",
    "email",
    "orcid",
    "scopus author id",
    "affiliation",
    "department of",
    "faculty of",
    "university",
    "institute",
    "institution",
    "hospital",
    "laboratory",
    "research centre",
    "research center",
    "address",
    "postal",
    "street",
]


def delete_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def paragraph_is_identifying(text: str) -> bool:
    low = normalize(text).lower()
    if not low:
        return False

    # These are safe to remove wherever they occur.
    if EMAIL_RE.search(text) or ORCID_RE.search(text):
        return True

    # Avoid deleting normal scientific/reference prose merely because it mentions
    # an institution. Standalone metadata blocks are normally short.
    if len(text) <= 180 and any(label in low for label in IDENTIFYING_LABELS):
        return True

    # Author/affiliation lines with superscript-style numbering in front matter.
    if re.search(r"\b\d+\s*[,;]\s*", text) and len(text) < 300:
        return True

    return False


def redact_run_text(text: str) -> str:
    if not text:
        return text
    text = EMAIL_RE.sub("[REDACTED]", text)
    text = ORCID_RE.sub("[REDACTED]", text)
    text = PHONE_RE.sub("[REDACTED]", text)
    # URLs can expose institutional pages or profiles.
    text = URL_RE.sub("[REDACTED]", text)
    return text


def iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

    for section in doc.sections:
        for container in [section.header, section.footer]:
            for p in container.paragraphs:
                yield p


def blind_copy_docx(original_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(original_bytes))

    # 1. Remove core metadata that can expose the author.
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""
    props.subject = ""
    props.keywords = ""

    # 2. Determine the abstract boundary.
    body_paragraphs = list(doc.paragraphs)
    abstract_idx = None
    for i, p in enumerate(body_paragraphs):
        if re.match(r"(?i)^\s*abstract\s*:?", p.text.strip()):
            abstract_idx = i
            break

    # 3. In the front matter, preserve the title but remove the author block.
    #    We keep the first substantial paragraph as the title and remove everything
    #    else before Abstract. This avoids trying to guess author names with AI.
    if abstract_idx is not None:
        nonempty_before = [
            (i, p) for i, p in enumerate(body_paragraphs[:abstract_idx])
            if p.text.strip()
        ]

        if nonempty_before:
            title_idx = nonempty_before[0][0]
            for i, p in enumerate(body_paragraphs[:abstract_idx]):
                if i == title_idx:
                    continue
                if p.text.strip():
                    delete_paragraph(p)

    # 4. Remove identifying standalone paragraphs throughout the document.
    #    Do not alter scientific prose simply because it contains a word like
    #    "university" in a reference or methods sentence; only standalone
    #    identifying blocks are deleted here.
    for p in list(iter_all_paragraphs(doc)):
        if p._element.getparent() is None:
            continue
        txt = p.text.strip()
        if not txt:
            continue
        if paragraph_is_identifying(txt):
            delete_paragraph(p)

    # 5. Remove sections commonly used to identify authors/funding.
    #    This is conservative for blind review: no replacement prose is inserted.
    removable_section_starts = {
        "funding",
        "acknowledgments",
        "acknowledgements",
    }

    paragraphs = list(doc.paragraphs)
    for i, p in enumerate(paragraphs):
        heading = normalize(p.text).lower().rstrip(":")
        if heading not in removable_section_starts:
            continue

        # Delete heading and following paragraphs until the next obvious heading.
        delete_paragraph(p)
        for q in paragraphs[i + 1:]:
            qtxt = normalize(q.text)
            if re.match(r"^(?:\d+(?:\.\d+)*)?\s*[A-Z][A-Za-z &/-]{2,60}$", qtxt):
                break
            if q._element.getparent() is not None:
                delete_paragraph(q)

    # 6. Redact explicit email/ORCID/phone/URLs in remaining runs while
    #    retaining paragraph/run formatting.
    for p in list(iter_all_paragraphs(doc)):
        for run in p.runs:
            run.text = redact_run_text(run.text)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def blind_copy_pdf_as_docx(pdf_bytes: bytes) -> bytes:
    """PDF -> DOCX fallback.

    A true layout-preserving PDF redaction requires PDF redaction tooling.
    This fallback produces a clean reviewer-copy DOCX without adding editorial
    text, but cannot guarantee identical PDF pagination.
    """
    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    # Use the same conservative text rules for a PDF fallback.
    lines = text.splitlines()
    output_lines = []
    abstract_seen = False
    title_kept = False

    for line in lines:
        stripped = line.strip()
        low = stripped.lower()

        if not abstract_seen:
            if not title_kept and stripped:
                output_lines.append(stripped)
                title_kept = True
                continue
            if low.startswith("abstract"):
                abstract_seen = True
                output_lines.append(stripped)
                continue
            # Remove front matter before abstract.
            continue

        if paragraph_is_identifying(stripped):
            continue

        if low in {"funding", "acknowledgments", "acknowledgements"}:
            continue

        output_lines.append(redact_run_text(line))

    doc = Document()
    for line in output_lines:
        if line.strip():
            doc.add_paragraph(line)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# REPORT GENERATION
# ============================================================

def add_status_table(doc: Document, checks: List[Dict[str, Any]]) -> None:
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "MRJ Requirement"
    hdr[1].text = "Status"
    hdr[2].text = "Evidence"
    hdr[3].text = "Action"

    for check in checks:
        cells = table.add_row().cells
        cells[0].text = check["requirement"]
        cells[1].text = check["status"]
        cells[2].text = check["evidence"]
        cells[3].text = check["action"]


def generate_report_docx(
    filename: str,
    checks: List[Dict[str, Any]],
    ai_data: Dict[str, Any],
    reviewers: Dict[str, List[Dict[str, Any]]],
) -> bytes:
    doc = Document()

    doc.add_heading("MRJ Editorial Pre-Screening Report", 0)
    doc.add_paragraph(
        "This report is an editorial pre-screening aid. It does not replace human editorial judgment or peer review."
    )

    doc.add_heading("1. MRJ Template Compliance", level=1)
    add_status_table(doc, checks)

    doc.add_heading("2. AI-Assisted Manuscript Assessment", level=1)

    fields = [
        ("Summary", ai_data.get("summary", "")),
        ("Research area", ai_data.get("research_area", "")),
        ("Methodology assessment", ai_data.get("methodology_assessment", "")),
        ("Results and discussion assessment", ai_data.get("results_discussion_assessment", "")),
        ("Abstract alignment", ai_data.get("abstract_alignment", "")),
        ("Editorial recommendation", ai_data.get("editorial_recommendation", "")),
    ]

    for label, value in fields:
        doc.add_heading(label, level=2)
        doc.add_paragraph(value or "Not available.")

    concerns = ai_data.get("potential_major_concerns", [])
    doc.add_heading("Potential major concerns", level=2)
    if concerns:
        for concern in concerns:
            doc.add_paragraph(concern, style="List Bullet")
    else:
        doc.add_paragraph("No major concerns were identified by the AI assessment.")

    doc.add_heading("3. Research Keywords Used for Reviewer Search", level=1)
    keywords = ai_data.get("keywords", [])
    doc.add_paragraph(", ".join(keywords) if keywords else "No keywords returned.")

    doc.add_heading("4. Reviewer Candidates — Verification Required", level=1)
    doc.add_paragraph(
        "Names below are retrieved from OpenAlex publication records; they are not generated by the language model. "
        "The listed institution is publication-associated or OpenAlex last-known affiliation and must be independently "
        "checked by the editor before invitation. A candidate is not automatically suitable and COI screening remains human responsibility."
    )

    for region in ["India", "Northeast India", "Assam"]:
        doc.add_heading(region, level=2)
        candidates = reviewers.get(region, [])

        if not candidates:
            doc.add_paragraph("No candidate found from the available OpenAlex results.")
            continue

        if "error" in candidates[0]:
            doc.add_paragraph(candidates[0]["error"])
            continue

        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        headers = ["Name", "Institution", "Location", "Evidence", "Verification"]
        for cell, header in zip(table.rows[0].cells, headers):
            cell.text = header

        for c in candidates:
            cells = table.add_row().cells
            cells[0].text = c["name"]
            cells[1].text = c.get("last_known_institution") or c.get("institution", "")
            cells[2].text = ", ".join(
                x for x in [
                    c.get("last_known_city", ""),
                    c.get("last_known_region", ""),
                    c.get("last_known_country", ""),
                ] if x
            )

            pubs = sorted(
                c.get("recent_publications", []),
                key=lambda x: x.get("year", 0),
                reverse=True,
            )[:3]

            evidence = []
            for p in pubs:
                doi = f" | {p['doi']}" if p.get("doi") else ""
                evidence.append(f"{p.get('year', '')}: {p.get('title', '')}{doi}")

            cells[3].text = "\n".join(evidence)
            cells[4].text = c.get("verification_note", "")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(
    page_title="MRJ AI Editorial Pre-Screening",
    page_icon="📄",
    layout="wide",
)

st.title("MRJ AI Editorial Pre-Screening")
st.caption(
    "MRJ-template-based compliance checking, Groq structured analysis, "
    "OpenAlex reviewer discovery, and blind-review document preparation."
)

with st.expander("Important workflow rules", expanded=False):
    st.markdown(
        """
- Reviewer names are **not generated by the AI**.
- MRJ formatting/compliance checks are deterministic wherever possible.
- Groq is used for interpretive editorial assessment and research-keyword extraction.
- The blind-review copy does **not** use AI rewriting.
- The editor must verify reviewer identity, current affiliation, expertise, availability and conflicts of interest before invitation.
        """
    )

uploaded_file = st.file_uploader(
    "Upload manuscript",
    type=["docx", "pdf"],
    help="DOCX is recommended when you need the closest possible formatting preservation for the blind copy.",
)

if uploaded_file:
    st.info(
        "For the best formatting-preserving blind copy, upload the original DOCX. "
        "A PDF can be analyzed, but PDF-to-DOCX reconstruction cannot guarantee identical pagination/layout."
    )

    groq_key = get_secret("GROQ_API_KEY")
    openalex_mailto = get_secret("OPENALEX_MAILTO")

    if not groq_key:
        st.error("GROQ_API_KEY is missing. Add it to Streamlit Secrets or the environment.")
        st.stop()

    if st.button("Run MRJ Pre-Screening", type="primary"):
        try:
            with st.spinner("Extracting manuscript..."):
                raw_text = extract_text(uploaded_file)

            if not raw_text.strip():
                st.error("No readable text was extracted from the uploaded file.")
                st.stop()

            with st.spinner("Running deterministic MRJ checks..."):
                deterministic_checks = run_mrj_rule_checks(raw_text)

            with st.spinner("Running structured Groq editorial assessment..."):
                client = Groq(api_key=groq_key)
                ai_data = run_ai_analysis(raw_text, deterministic_checks, client)

            with st.spinner("Searching OpenAlex for real reviewer candidates..."):
                reviewers = reviewer_search_report(ai_data, openalex_mailto)

            with st.spinner("Creating blind-review copy..."):
                original_bytes = uploaded_file.getvalue()
                if uploaded_file.name.lower().endswith(".docx"):
                    blind_bytes = blind_copy_docx(original_bytes)
                    blind_filename = "MRJ_Blind_Reviewer_Copy.docx"
                else:
                    blind_bytes = blind_copy_pdf_as_docx(original_bytes)
                    blind_filename = "MRJ_Blind_Reviewer_Copy.docx"

            with st.spinner("Building editorial report..."):
                report_bytes = generate_report_docx(
                    uploaded_file.name,
                    deterministic_checks,
                    ai_data,
                    reviewers,
                )

            st.session_state["mrj_checks"] = deterministic_checks
            st.session_state["mrj_ai"] = ai_data
            st.session_state["mrj_reviewers"] = reviewers
            st.session_state["mrj_blind"] = blind_bytes
            st.session_state["mrj_blind_filename"] = blind_filename
            st.session_state["mrj_report"] = report_bytes
            st.session_state["mrj_raw_text"] = raw_text

            st.success("Pre-screening completed.")

        except requests.HTTPError as exc:
            st.error(f"OpenAlex/API error: {exc}")
        except Exception as exc:
            st.exception(exc)

if "mrj_checks" in st.session_state:
    checks = st.session_state["mrj_checks"]
    ai_data = st.session_state["mrj_ai"]
    reviewers = st.session_state["mrj_reviewers"]

    st.subheader("1. MRJ Compliance")
    status_counts = {
        "PASS": sum(x["status"] == "PASS" for x in checks),
        "WARN": sum(x["status"] == "WARN" for x in checks),
        "FAIL": sum(x["status"] == "FAIL" for x in checks),
    }

    c1, c2, c3 = st.columns(3)
    c1.metric("Pass", status_counts["PASS"])
    c2.metric("Warnings", status_counts["WARN"])
    c3.metric("Failures", status_counts["FAIL"])

    st.dataframe(
        [
            {
                "Requirement": x["requirement"],
                "Status": x["status"],
                "Evidence": x["evidence"],
                "Action": x["action"],
            }
            for x in checks
        ],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("2. AI Editorial Assessment")
    st.write(ai_data.get("summary", ""))
    st.write("**Research area:**", ai_data.get("research_area", ""))
    st.write("**Methodology:**", ai_data.get("methodology_assessment", ""))
    st.write("**Results & Discussion:**", ai_data.get("results_discussion_assessment", ""))
    st.write("**Abstract alignment:**", ai_data.get("abstract_alignment", ""))
    st.write("**Editorial recommendation:**", ai_data.get("editorial_recommendation", ""))

    concerns = ai_data.get("potential_major_concerns", [])
    if concerns:
        st.warning("Potential concerns identified by AI")
        for item in concerns:
            st.write(f"- {item}")

    st.subheader("3. Reviewer Candidates")
    st.caption(
        "These are retrieved from OpenAlex publication records. They are not hallucinated by the LLM. "
        "Verify current affiliation, expertise and conflicts of interest before contacting anyone."
    )

    for region, candidates in reviewers.items():
        st.markdown(f"### {region}")

        if not candidates:
            st.write("No candidates found.")
            continue

        if "error" in candidates[0]:
            st.error(candidates[0]["error"])
            continue

        rows = []
        for c in candidates:
            pubs = sorted(
                c.get("recent_publications", []),
                key=lambda p: p.get("year", 0),
                reverse=True,
            )[:3]
            rows.append({
                "Name": c["name"],
                "Institution": c.get("last_known_institution") or c.get("institution", ""),
                "Location": ", ".join(
                    x for x in [
                        c.get("last_known_city", ""),
                        c.get("last_known_region", ""),
                        c.get("last_known_country", ""),
                    ] if x
                ),
                "Recent publication evidence": "\n".join(
                    f"{p.get('year', '')}: {p.get('title', '')}" for p in pubs
                ),
                "Verification": c.get("verification_note", ""),
            })

        st.dataframe(rows, use_container_width=True, hide_index=True)

    st.subheader("4. Downloads")

    st.download_button(
        "Download MRJ Editorial Report (.docx)",
        data=st.session_state["mrj_report"],
        file_name="MRJ_Editorial_PreScreening_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        "Download Blind Reviewer Copy (.docx)",
        data=st.session_state["mrj_blind"],
        file_name=st.session_state["mrj_blind_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    with st.expander("Technical note about blind-copy formatting"):
        st.write(
            "DOCX input is edited in-place at the document level: author/front-matter blocks and "
            "identifying sections are removed, while existing runs and document formatting are retained "
            "where possible. The application does not ask the LLM to rewrite the manuscript."
        )
