import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
from typing import Any, Dict, List, Tuple

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
    "conflicts of interest": ["conflicts of interest", "conflict of interest", "competing interests"],
    "declaration on ai usage": [
        "declaration on ai usage",
        "ai usage",
        "artificial intelligence",
    ],
    "references": ["references", "reference", "bibliography"],
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
    "tezpur university",
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
    "manipur university",
    "mizoram university",
    "nagaland university",
    "tripura university",
    "rajiv gandhi university",
    "sikkim university",
]

ASSAM_TERMS = [
    "assam",
    "guwahati",
    "silchar",
    "tezpur",
    "dibrugarh",
    "jorhat",
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
    """Retrieve environment secrets from Streamlit secrets or OS environment."""
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
    start = None
    for i, line in enumerate(lines):
        norm = normalize(line).lower().rstrip(":")
        if norm in ("abstract", "1. abstract"):
            start = i
            break

    if start is None:
        m = re.search(r"(?is)\babstract\s*:\s*(.*?)(?:\bkeywords\s*:|$)", text)
        return normalize(m.group(1)) if m else ""

    end_candidates = [
        i for i, line in enumerate(lines)
        if i > start and (
            normalize(line).lower().startswith("keywords") or
            normalize(line).lower().startswith("1. introduction") or
            normalize(line).lower() == "introduction"
        )
    ]
    end = min(end_candidates) if end_candidates else len(lines)
    value = "\n".join(lines[start + 1:end]).strip()
    if not value and ":" in lines[start]:
        value = lines[start].split(":", 1)[1]
    return normalize(value)


def extract_keywords(text: str) -> List[str]:
    m = re.search(
        r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:1\.?\s+)?introduction\b|\n\s*abstract\b|\n\s*\n\s*[A-Z]|$)",
        text,
    )
    if not m:
        return []
    # Join multi-line keywords before splitting
    raw = " ".join([line.strip() for line in m.group(1).strip().splitlines() if line.strip()])
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

    raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")


# ============================================================
# DETERMINISTIC MRJ PRE-SCREENING
# ============================================================

def run_mrj_rule_checks(text: str) -> List[Dict[str, Any]]:
    checks = []
    lower = text.lower()
    lines = text.splitlines()

    abstract = extract_abstract(text)
    keywords = extract_keywords(text)
    positions = find_section_positions(lines)

    # 1. Abstract check
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
            "action": "Keep abstract at or below 200 words." if status == "FAIL" else "No action required.",
        })

    # 2. Keywords check
    if MRJ_RULES["keywords_min"] <= len(keywords) <= MRJ_RULES["keywords_max"]:
        keyword_status = "PASS"
        keyword_action = "No action required."
    elif len(keywords) == 0:
        keyword_status = "FAIL"
        keyword_action = "Add 3–10 pertinent keywords."
    else:
        keyword_status = "FAIL"
        keyword_action = "Adjust keywords count to between 3 and 10."
    checks.append({
        "requirement": "3–10 keywords",
        "status": keyword_status,
        "evidence": f"Detected {len(keywords)} keyword(s): {', '.join(keywords) if keywords else 'none detected'}.",
        "action": keyword_action,
    })

    # 3. Required sections check
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
            "action": "No action required." if exists else f"Add the '{label}' section required by MRJ.",
        })

    # 4. Multidisciplinary domains check
    domain_statement = extract_domain_statement(text)
    domain_matches = re.findall(r"(?:\([a-z0-9]+\)|\b\d+\.|\*|-)\s*([^,;.\n]+)", domain_statement, flags=re.I)
    domain_count = len(domain_matches)
    if domain_count >= MRJ_RULES["minimum_domains"]:
        domain_status = "PASS"
        domain_action = "No action required."
    else:
        domain_status = "WARN" if domain_statement else "FAIL"
        domain_action = "Explicitly list at least two domains in the Multidisciplinary Domains statement."

    checks.append({
        "requirement": "At least two multidisciplinary domains",
        "status": domain_status,
        "evidence": f"Detected {domain_count} domain items." if domain_statement else "No domain statement detected.",
        "action": domain_action,
    })

    # 5. Citation style check (using last reference index to avoid early false match)
    ref_pos = positions.get("references")
    if ref_pos is not None:
        body_before_refs = "\n".join(lines[:ref_pos])
    else:
        last_ref_idx = lower.rfind("references")
        body_before_refs = text[:last_ref_idx] if last_ref_idx != -1 else text

    numbered_citations = re.findall(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]", body_before_refs)
    citation_status = "PASS" if numbered_citations else "WARN"
    checks.append({
        "requirement": "Numbered square-bracket citations",
        "status": citation_status,
        "evidence": f"Detected {len(numbered_citations)} square-bracket citation pattern(s).",
        "action": "Check that citations follow sequential square-bracket formatting (e.g., [1, 2]).",
    })

    # 6. Ethical approval check
    methods = extract_section_text(text, "materials and methods").lower()
    ethical_terms = [
        "ethics approval", "ethical approval", "ethics committee",
        "institutional review board", "irb", "approval code", "ethical clearance"
    ]
    ethics_relevant = any(
        term in (methods + " " + lower) for term in
        ["human", "patient", "participant", "animal", "clinical", "intervention", "survey"]
    )
    if ethics_relevant:
        ethics_found = any(term in methods for term in ethical_terms)
        checks.append({
            "requirement": "Ethics approval statement",
            "status": "PASS" if ethics_found else "WARN",
            "evidence": "Study involves human/animal/participant subjects; ethics wording was " +
                        ("detected in Methods." if ethics_found else "NOT clearly identified in Materials and Methods."),
            "action": "Verify IRB/ethics committee name and protocol approval code are clearly specified.",
        })

    # 7. DOI check
    references = extract_section_text(text, "references")
    doi_count = len(re.findall(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", references, flags=re.I))
    checks.append({
        "requirement": "DOI inclusion in references",
        "status": "PASS" if doi_count > 0 else "WARN",
        "evidence": f"Detected {doi_count} DOI reference pattern(s).",
        "action": "Verify that digital object identifiers (DOIs) are supplied where available.",
    })

    return checks


# ============================================================
# GROQ STRUCTURED ANALYSIS
# ============================================================

STATUS_VALUES = ["PASS", "CONCERN", "MAJOR CONCERN", "NOT ASSESSABLE"]

def assessment_schema(name: str, properties: Dict[str, Any], required: List[str]) -> Dict[str, Any]:
    return {
        "name": name,
        "strict": True,
        "schema": {
            "type": "object",
            "properties": properties,
            "required": required,
            "additionalProperties": False,
        },
    }

ITEM_SCHEMA = {
    "type": "object",
    "properties": {
        "status": {"type": "string", "enum": STATUS_VALUES},
        "finding": {"type": "string"},
        "evidence": {"type": "string"},
        "action": {"type": "string"},
    },
    "required": ["status", "finding", "evidence", "action"],
    "additionalProperties": False,
}

AI_CALL_1_SCHEMA = assessment_schema(
    "mrj_scope_abstract_introduction",
    {
        "research_area": {"type": "string"},
        "keywords": {"type": "array", "items": {"type": "string"}},
        "research_question": ITEM_SCHEMA,
        "abstract": ITEM_SCHEMA,
        "introduction_novelty": ITEM_SCHEMA,
    },
    ["research_area", "keywords", "research_question", "abstract", "introduction_novelty"],
)

AI_CALL_2_SCHEMA = assessment_schema(
    "mrj_methodology_statistics_ethics",
    {
        "methodology": ITEM_SCHEMA,
        "statistics": ITEM_SCHEMA,
        "ethics_reproducibility": ITEM_SCHEMA,
    },
    ["methodology", "statistics", "ethics_reproducibility"],
)

AI_CALL_3_SCHEMA = assessment_schema(
    "mrj_results_discussion_conclusion",
    {
        "results": ITEM_SCHEMA,
        "discussion": ITEM_SCHEMA,
        "conclusion": ITEM_SCHEMA,
        "abstract_conclusion_alignment": ITEM_SCHEMA,
        "reference_use": ITEM_SCHEMA,
        "major_red_flags": {"type": "array", "items": {"type": "string"}},
        "editorial_recommendation": {
            "type": "string",
            "enum": ["Proceed to editorial review", "Needs author correction", "Major concern"],
        },
    },
    [
        "results", "discussion", "conclusion",
        "abstract_conclusion_alignment", "reference_use",
        "major_red_flags", "editorial_recommendation"
    ],
)


def _groq_json_call(client: Groq, model: str, schema: Dict[str, Any], prompt: str, max_tokens: int = 2000) -> Dict[str, Any]:
    """Call Groq API with structured JSON output and automatic retry on budget overrun."""
    request_params = dict(
        model=model,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a strict, cautious academic pre-screening editorial assistant. "
                    "Base your assessment entirely on the provided excerpt. Never invent citations, results, "
                    "institutions, or facts. If evidence is absent, use NOT ASSESSABLE. Keep findings concise. "
                    "Output strictly valid JSON complying with the provided schema."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        max_completion_tokens=max_tokens,
        response_format={"type": "json_schema", "json_schema": schema},
    )

    try:
        response = client.chat.completions.create(**request_params)
    except Exception as exc:
        msg = str(exc).lower()
        if "max completion tokens" in msg or "json_validate_failed" in msg:
            request_params["max_completion_tokens"] = 3500
            response = client.chat.completions.create(**request_params)
        else:
            raise exc

    return json.loads(response.choices[0].message.content or "{}")


def _ai_not_assessed(reason: str) -> Dict[str, Any]:
    item = {"status": "NOT ASSESSABLE", "finding": reason, "evidence": "", "action": "Review manually."}
    keys = [
        "research_question", "abstract", "introduction_novelty", "methodology",
        "statistics", "ethics_reproducibility", "results", "discussion",
        "conclusion", "abstract_conclusion_alignment", "reference_use"
    ]
    return {
        "summary": reason,
        "research_area": "Not assessed",
        "keywords": [],
        **{k: item.copy() for k in keys},
        "major_red_flags": [reason],
        "editorial_recommendation": "Needs author correction",
    }


def _compact_rule_summary(checks: List[Dict[str, Any]]) -> str:
    return "\n".join(f"- {c['requirement']}: {c['status']} — {c['evidence']}" for c in checks)


def run_ai_analysis(
    text: str,
    deterministic_checks: List[Dict[str, Any]],
    client: Groq,
    model: str = DEFAULT_GROQ_MODEL
) -> Dict[str, Any]:
    """Execute three focused, parallel AI assessments for low latency."""
    abstract = extract_abstract(text)
    intro = extract_section_text(text, "introduction")
    methods = extract_section_text(text, "materials and methods")
    rd = extract_section_text(text, "results and discussion")
    conclusion = extract_section_text(text, "conclusions")
    refs = extract_section_text(text, "references")
    rules = _compact_rule_summary(deterministic_checks)

    prompt1 = f"""MRJ PRE-SCREEN: RESEARCH QUESTION, ABSTRACT, INTRODUCTION
MRJ DETERMINISTIC CHECKS:
{rules[:3000]}

ABSTRACT:
{abstract[:2500]}

INTRODUCTION:
{intro[:3500]}

Assess: explicit research question/objective; abstract components; and introduction gap/novelty. Extract specific research area and 3-5 search keywords."""

    prompt2 = f"""MRJ PRE-SCREEN: METHODOLOGY, STATISTICS, ETHICS
MRJ DETERMINISTIC CHECKS:
{rules[:3000]}

MATERIALS AND METHODS:
{methods[:5000]}

Assess: methodology completeness and reproducibility; appropriateness of statistics/data analysis; and ethics/IRB approvals if applicable."""

    prompt3 = f"""MRJ PRE-SCREEN: RESULTS, DISCUSSION, CONCLUSION, REFERENCES
ABSTRACT:
{abstract[:2000]}

RESULTS AND DISCUSSION:
{rd[:5000]}

CONCLUSIONS:
{conclusion[:2000]}

REFERENCES:
{refs[:2500]}

Assess: whether results address the research question; discussion interprets rather than merely re-states; conclusions are supported; citation coherence. List evidence-backed red flags."""

    try:
        with ThreadPoolExecutor(max_workers=3) as executor:
            fut1 = executor.submit(_groq_json_call, client, model, AI_CALL_1_SCHEMA, prompt1, 2000)
            fut2 = executor.submit(_groq_json_call, client, model, AI_CALL_2_SCHEMA, prompt2, 2000)
            fut3 = executor.submit(_groq_json_call, client, model, AI_CALL_3_SCHEMA, prompt3, 2200)

            res1 = fut1.result()
            res2 = fut2.result()
            res3 = fut3.result()

        out = {}
        out.update(res1)
        out.update(res2)
        out.update(res3)

        assess_keys = [
            "research_question", "abstract", "introduction_novelty",
            "methodology", "statistics", "ethics_reproducibility",
            "results", "discussion", "conclusion",
            "abstract_conclusion_alignment", "reference_use"
        ]
        concern_count = sum(out.get(k, {}).get("status") in ("CONCERN", "MAJOR CONCERN") for k in assess_keys)
        out["summary"] = (
            f"Focused AI screening completed across 11 key scientific facets. {concern_count} item(s) flagged with concerns."
            if concern_count else
            "Focused AI assessment completed across 11 scientific areas. No critical methodological or structural red flags identified."
        )
        return out

    except Exception as exc:
        msg = str(exc)
        if "413" in msg or "tokens per minute" in msg.lower():
            return _ai_not_assessed("Groq rate limit encountered. Deterministic checks remain valid.")
        return _ai_not_assessed(f"AI assessment failed: {msg}")


def _format_ai_item(item: Dict[str, Any]) -> str:
    return (
        f"Status: {item.get('status', 'NOT ASSESSABLE')}\n"
        f"Finding: {item.get('finding', '')}\n"
        f"Evidence: {item.get('evidence', '') or 'Not supplied.'}\n"
        f"Action: {item.get('action', '') or 'Manual review required.'}"
    )


# ============================================================
# REVIEWER SEARCH: OPENALEX (OPTIMIZED)
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    response = requests.get(url, params=params, timeout=20)
    response.raise_for_status()
    return response.json()


@lru_cache(maxsize=256)
def get_openalex_author(author_id: str, mailto: str = "") -> Dict[str, Any]:
    try:
        return openalex_get(f"{OPENALEX_AUTHOR_URL}/{author_id}", {}, mailto)
    except Exception:
        return {}


def candidate_region_match(candidate: Dict[str, Any], region: str) -> bool:
    text = " ".join([
        str(candidate.get("institution") or ""),
        str(candidate.get("city") or ""),
        str(candidate.get("region") or ""),
        str(candidate.get("country") or ""),
    ]).lower()

    if region == "India":
        return candidate.get("country", "").lower() == "in" or "india" in text

    if region == "Northeast India":
        return any(term in text for term in (set(NORTHEAST_STATES) | set(NORTHEAST_INSTITUTION_TERMS)))

    if region == "Assam":
        return any(term in text for term in ASSAM_TERMS)

    return False


def reviewer_search_report(
    ai_data: Dict[str, Any],
    mailto: str = "",
) -> Dict[str, List[Dict[str, Any]]]:
    """Single-pass OpenAlex query with author-level caching and regional partitioning."""
    terms = ai_data.get("keywords", [])[:4] + [ai_data.get("research_area", "")]
    terms = [normalize(x) for x in terms if normalize(x)]
    query = " ".join(terms[:5]).strip()

    output = {"India": [], "Northeast India": [], "Assam": []}
    if not query:
        return output

    try:
        data = openalex_get(
            OPENALEX_URL,
            {
                "search": query,
                "per-page": 50,
                "sort": "publication_year:desc",
            },
            mailto,
        )
    except Exception as exc:
        err = [{"error": f"OpenAlex search failed: {exc}"}]
        return {k: err for k in output}

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
                geo = inst.get("geo") or {}

                if author_id not in candidates:
                    candidates[author_id] = {
                        "author_id": author_id,
                        "name": author_name,
                        "institution": inst_name,
                        "country": inst_country,
                        "city": geo.get("city") or "",
                        "region": geo.get("region") or "",
                        "recent_publications": [],
                        "score": 0.0,
                    }

                rec = candidates[author_id]
                rec["recent_publications"].append({
                    "title": title,
                    "year": year,
                    "doi": doi,
                    "cited_by_count": cited,
                })
                rec["score"] += max(0, year - 2018) * 0.8 + min(cited, 100) * 0.03

    # Resolve affiliations only for top scoring candidate pool (max 25)
    top_candidates = sorted(candidates.values(), key=lambda x: x["score"], reverse=True)[:25]
    for c in top_candidates:
        author_meta = get_openalex_author(c["author_id"], mailto)
        last_known = (author_meta.get("last_known_institutions") or [{}])[0]
        geo = last_known.get("geo") or {}

        c["last_known_institution"] = last_known.get("display_name") or c["institution"]
        c["last_known_country"] = last_known.get("country_code") or c["country"]
        c["last_known_city"] = geo.get("city") or c["city"]
        c["last_known_region"] = geo.get("region") or c["region"]

    # Filter into regional buckets locally
    for reg in ["Assam", "Northeast India", "India"]:
        matched = []
        for c in top_candidates:
            check_obj = {
                "institution": c["last_known_institution"],
                "country": c["last_known_country"],
                "city": c["last_known_city"],
                "region": c["last_known_region"],
            }
            if candidate_region_match(check_obj, reg):
                c["verification_note"] = "OpenAlex last-known affiliation matches region."
                matched.append(c)
            elif candidate_region_match(c, reg):
                c["verification_note"] = "Historical publication affiliation matches region."
                matched.append(c)

        matched.sort(
            key=lambda x: (
                len(x["recent_publications"]),
                x["score"],
                max((p["year"] for p in x["recent_publications"]), default=0),
            ),
            reverse=True,
        )
        output[reg] = matched[:6]

    return output


# ============================================================
# BLIND REVIEW COPY
# ============================================================

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ORCID_RE = re.compile(r"\b(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}\b", re.I)
PHONE_RE = re.compile(r"(?<!\d)(?:\+?\d[\d\s().-]{7,}\d)(?!\d)")
URL_RE = re.compile(r"https?://\S+", re.I)

IDENTIFYING_LABELS = [
    "correspondence", "corresponding author", "email", "orcid",
    "scopus author id", "affiliation", "department of", "faculty of",
    "university", "institute", "hospital", "laboratory",
    "address", "postal"
]

REMOVABLE_SECTIONS = {
    "funding", "acknowledgments", "acknowledgements",
    "author contributions", "author's contributions", "authors' contributions",
    "competing interests", "biography", "about the authors",
}


def delete_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    if p is not None and p.getparent() is not None:
        p.getparent().remove(p)


def paragraph_is_identifying(text: str) -> bool:
    low = normalize(text).lower()
    if not low:
        return False
    if EMAIL_RE.search(text) or ORCID_RE.search(text):
        return True
    if len(text) <= 180 and any(label in low for label in IDENTIFYING_LABELS):
        return True
    if re.search(r"\b\d+\s*[,;]\s*", text) and len(text) < 250 and any(w in low for w in ["dept", "department", "univ"]):
        return True
    return False


def redact_run_text(text: str) -> str:
    if not text:
        return text
    text = EMAIL_RE.sub("[REDACTED EMAIL]", text)
    text = ORCID_RE.sub("[REDACTED ORCID]", text)
    text = PHONE_RE.sub("[REDACTED PHONE]", text)
    text = URL_RE.sub("[REDACTED LINK]", text)
    return text


def iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        if p is not None:
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

    # 1. Clear core metadata
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""
    props.subject = ""
    props.keywords = ""

    # 2. Identify Abstract boundary
    body_paragraphs = list(doc.paragraphs)
    abstract_idx = None
    for i, p in enumerate(body_paragraphs):
        if re.match(r"(?i)^\s*abstract\s*:?", p.text.strip()):
            abstract_idx = i
            break

    # 3. Retain first paragraph (Title) and delete front-matter authors/affiliations
    if abstract_idx is not None:
        nonempty_before = [(i, p) for i, p in enumerate(body_paragraphs[:abstract_idx]) if p.text.strip()]
        if nonempty_before:
            title_idx = nonempty_before[0][0]
            for i, p in enumerate(body_paragraphs[:abstract_idx]):
                if i != title_idx and p.text.strip():
                    delete_paragraph(p)

    # 4. Remove standalone identifying paragraphs
    for p in list(iter_all_paragraphs(doc)):
        txt = p.text.strip()
        if txt and paragraph_is_identifying(txt):
            delete_paragraph(p)

    # 5. Remove funding, acknowledgment, and contribution sections
    paragraphs = [p for p in doc.paragraphs if p is not None]
    for i, p in enumerate(paragraphs):
        heading = normalize(p.text).lower().rstrip(":")
        if heading in REMOVABLE_SECTIONS:
            delete_paragraph(p)
            for q in paragraphs[i + 1:]:
                qtxt = normalize(q.text)
                if re.match(r"^(?:\d+(?:\.\d+)*)?\s*[A-Z][A-Za-z &/-]{2,60}$", qtxt):
                    break
                delete_paragraph(q)

    # 6. Redact identifiable runs (emails, links, ORCIDs)
    for p in list(iter_all_paragraphs(doc)):
        for run in p.runs:
            run.text = redact_run_text(run.text)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def blind_copy_pdf_as_docx(pdf_bytes: bytes) -> bytes:
    """PDF text fallback for blind reviewer copy."""
    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

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
            continue

        if paragraph_is_identifying(stripped):
            continue
        if low in REMOVABLE_SECTIONS:
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
# REPORT GENERATION (.DOCX)
# ============================================================

def generate_report_docx(
    filename: str,
    checks: List[Dict[str, Any]],
    ai_data: Dict[str, Any],
    reviewers: Dict[str, List[Dict[str, Any]]],
) -> bytes:
    doc = Document()
    doc.add_heading("MRJ Editorial Pre-Screening Report", 0)
    doc.add_paragraph("Editorial decision-support audit. Peer review and editor verification remain required.")

    # Section 1: Template Compliance
    doc.add_heading("1. MRJ Template Compliance", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Requirement", "Status", "Evidence", "Action"

    for check in checks:
        cells = table.add_row().cells
        cells[0].text = check["requirement"]
        cells[1].text = check["status"]
        cells[2].text = check["evidence"]
        cells[3].text = check["action"]

    # Section 2: AI Scientific Review
    doc.add_heading("2. AI-Assisted Manuscript Assessment", level=1)
    doc.add_paragraph(f"Summary: {ai_data.get('summary', 'Not available.')}")
    doc.add_paragraph(f"Research Area: {ai_data.get('research_area', 'Not assessed')}")

    groups = [
        ("Research question / objective", "research_question"),
        ("Abstract quality", "abstract"),
        ("Introduction and novelty", "introduction_novelty"),
        ("Methodology completeness", "methodology"),
        ("Statistics / data analysis", "statistics"),
        ("Ethics and reproducibility", "ethics_reproducibility"),
        ("Results", "results"),
        ("Discussion", "discussion"),
        ("Conclusion", "conclusion"),
        ("Abstract–conclusion alignment", "abstract_conclusion_alignment"),
        ("Reference use", "reference_use"),
    ]

    for label, key in groups:
        doc.add_heading(label, level=2)
        doc.add_paragraph(_format_ai_item(ai_data.get(key, {})))

    doc.add_heading("Major Red Flags", level=2)
    concerns = ai_data.get("major_red_flags", [])
    if concerns:
        for c in concerns:
            doc.add_paragraph(c, style="List Bullet")
    else:
        doc.add_paragraph("No major red flags detected from supplied evidence.")

    doc.add_heading("Pre-Screening Editorial Recommendation", level=2)
    doc.add_paragraph(ai_data.get("editorial_recommendation", "Not assessed"))

    # Section 3: Reviewers
    doc.add_heading("3. Reviewer Candidates (OpenAlex Live Verified)", level=1)
    for region in ["India", "Northeast India", "Assam"]:
        doc.add_heading(region, level=2)
        candidates = reviewers.get(region, [])

        if not candidates:
            doc.add_paragraph("No candidate found from current OpenAlex query.")
            continue
        if "error" in candidates[0]:
            doc.add_paragraph(candidates[0]["error"])
            continue

        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        h = tbl.rows[0].cells
        h[0].text, h[1].text, h[2].text, h[3].text = "Name", "Affiliation", "Recent Publications", "Verification"

        for c in candidates:
            row = tbl.add_row().cells
            row[0].text = c["name"]
            row[1].text = f"{c.get('last_known_institution', '')} ({c.get('last_known_city', '')}, {c.get('last_known_country', '')})"
            pubs = c.get("recent_publications", [])[:2]
            row[2].text = "\n".join(f"- {p.get('year', '')}: {p.get('title', '')}" for p in pubs)
            row[3].text = c.get("verification_note", "")

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

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    groq_key = st.text_input(
        "Groq API Key",
        value=get_secret("GROQ_API_KEY"),
        type="password",
        help="Enter your Groq API key (starts with gsk_)"
    )
    openalex_mailto = st.text_input(
        "OpenAlex Mailto Email",
        value=get_secret("OPENALEX_MAILTO", "editor@example.com"),
        help="Email sent in User-Agent for OpenAlex courteous pool"
    )
    selected_model = st.selectbox(
        "Groq Model",
        options=[DEFAULT_GROQ_MODEL, "llama-3.3-70b-versatile", "llama-3.1-8b-instant"],
        index=0
    )
    st.markdown("---")
    st.markdown("**MRJ Pre-Screening Engine v2.1**\n- Deterministic rule checks\n- Parallel structured LLM audit\n- Live OpenAlex discovery")

st.title("MRJ AI Editorial Pre-Screening")
st.caption("Standardized compliance evaluation, concurrent Groq structured analysis, and automated double-blind copy generator.")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file:
    if st.button("Run Pre-Screening Analysis", type="primary"):
        if not groq_key:
            st.error("Groq API Key is required. Please provide it in the sidebar or Streamlit secrets.")
            st.stop()

        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            status_text.text("Extracting manuscript text...")
            raw_text = extract_text(uploaded_file)
            progress_bar.progress(20)

            if not raw_text.strip():
                st.error("No readable text found in manuscript.")
                st.stop()

            status_text.text("Executing deterministic MRJ checks...")
            deterministic_checks = run_mrj_rule_checks(raw_text)
            progress_bar.progress(40)

            status_text.text("Running parallel Groq scientific analysis...")
            client = Groq(api_key=groq_key)
            ai_data = run_ai_analysis(raw_text, deterministic_checks, client, model=selected_model)
            progress_bar.progress(65)

            status_text.text("Searching OpenAlex for candidates...")
            reviewers = reviewer_search_report(ai_data, openalex_mailto)
            progress_bar.progress(80)

            status_text.text("Generating blind reviewer copy...")
            orig_bytes = uploaded_file.getvalue()
            if uploaded_file.name.lower().endswith(".docx"):
                blind_bytes = blind_copy_docx(orig_bytes)
            else:
                blind_bytes = blind_copy_pdf_as_docx(orig_bytes)
            progress_bar.progress(90)

            status_text.text("Compiling editorial DOCX report...")
            report_bytes = generate_report_docx(
                uploaded_file.name,
                deterministic_checks,
                ai_data,
                reviewers
            )
            progress_bar.progress(100)
            status_text.empty()

            # Store in session state
            st.session_state["mrj_checks"] = deterministic_checks
            st.session_state["mrj_ai"] = ai_data
            st.session_state["mrj_reviewers"] = reviewers
            st.session_state["mrj_blind"] = blind_bytes
            st.session_state["mrj_report"] = report_bytes
            st.success("Pre-screening completed successfully.")

        except Exception as exc:
            st.exception(exc)

# Display Results if available
if "mrj_checks" in st.session_state:
    checks = st.session_state["mrj_checks"]
    ai_data = st.session_state["mrj_ai"]
    reviewers = st.session_state["mrj_reviewers"]

    st.subheader("1. MRJ Formatting & Rule Compliance")
    c1, c2, c3 = st.columns(3)
    c1.metric("Pass", sum(x["status"] == "PASS" for x in checks))
    c2.metric("Warnings", sum(x["status"] == "WARN" for x in checks))
    c3.metric("Fails", sum(x["status"] == "FAIL" for x in checks))

    st.dataframe(
        [{"Requirement": x["requirement"], "Status": x["status"], "Evidence": x["evidence"], "Action": x["action"]} for x in checks],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("2. AI Scientific Pre-Screening Assessment")
    st.info(f"**Overview:** {ai_data.get('summary', '')}")
    st.write(f"**Research Area:** {ai_data.get('research_area', '')}")
    st.write(f"**Keywords Extracted:** {', '.join(ai_data.get('keywords', []))}")

    groups = [
        ("Research question / objective", "research_question"),
        ("Abstract", "abstract"),
        ("Introduction & Novelty", "introduction_novelty"),
        ("Methodology", "methodology"),
        ("Statistics", "statistics"),
        ("Ethics & Reproducibility", "ethics_reproducibility"),
        ("Results", "results"),
        ("Discussion", "discussion"),
        ("Conclusions", "conclusion"),
        ("Alignment", "abstract_conclusion_alignment"),
        ("References", "reference_use"),
    ]

    for label, key in groups:
        item = ai_data.get(key, {})
        with st.expander(f"{label} — {item.get('status', 'NOT ASSESSABLE')}"):
            st.write(f"**Finding:** {item.get('finding', '')}")
            st.write(f"**Evidence:** {item.get('evidence', '') or 'None supplied'}")
            st.write(f"**Action:** {item.get('action', '') or 'Manual check'}")

    red_flags = ai_data.get("major_red_flags", [])
    if red_flags:
        st.warning("⚠️ Flagged Major Concerns:")
        for rf in red_flags:
            st.write(f"- {rf}")

    st.write(f"**Recommendation:** `{ai_data.get('editorial_recommendation', 'Not assessed')}`")

    st.subheader("3. Reviewer Candidates (Live OpenAlex)")
    for region, candidates in reviewers.items():
        st.markdown(f"#### {region}")
        if not candidates or "error" in candidates[0]:
            st.write("No candidates found or query error.")
            continue

        rows = []
        for c in candidates:
            pubs = c.get("recent_publications", [])[:2]
            pub_text = " | ".join(f"{p.get('year', '')}: {p.get('title', '')[:45]}..." for p in pubs)
            rows.append({
                "Name": c["name"],
                "Institution": c.get("last_known_institution") or c.get("institution", ""),
                "Location": f"{c.get('last_known_city', '')}, {c.get('last_known_country', '')}",
                "Recent Works": pub_text,
                "Match Note": c.get("verification_note", "")
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)

    st.subheader("4. Downloads")
    d1, d2 = st.columns(2)
    d1.download_button(
        "📥 Download Editorial Report (.docx)",
        data=st.session_state["mrj_report"],
        file_name="MRJ_PreScreening_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    d2.download_button(
        "📥 Download Blind Reviewer Copy (.docx)",
        data=st.session_state["mrj_blind"],
        file_name="MRJ_Blind_Review_Copy.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
