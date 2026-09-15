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

STATUS_VALUES = ["PASS", "CONCERN", "MAJOR CONCERN", "NOT ASSESSABLE"]


def assessment_schema(name: str, properties: Dict[str, Any], required: List[str]) -> Dict[str, Any]:
    return {"name": name, "strict": True, "schema": {"type": "object", "properties": properties, "required": required, "additionalProperties": False}}


ITEM_SCHEMA = {"type": "object", "properties": {
    "status": {"type": "string", "enum": STATUS_VALUES}, "finding": {"type": "string"},
    "evidence": {"type": "string"}, "action": {"type": "string"}},
    "required": ["status", "finding", "evidence", "action"], "additionalProperties": False}

AI_CALL_1_SCHEMA = assessment_schema("mrj_scope_abstract_introduction", {
    "research_area": {"type": "string"}, "keywords": {"type": "array", "items": {"type": "string"}},
    "research_question": ITEM_SCHEMA, "abstract": ITEM_SCHEMA, "introduction_novelty": ITEM_SCHEMA},
    ["research_area", "keywords", "research_question", "abstract", "introduction_novelty"])
AI_CALL_2_SCHEMA = assessment_schema("mrj_methodology_statistics_ethics", {
    "methodology": ITEM_SCHEMA, "statistics": ITEM_SCHEMA, "ethics_reproducibility": ITEM_SCHEMA},
    ["methodology", "statistics", "ethics_reproducibility"])
AI_CALL_3_SCHEMA = assessment_schema("mrj_results_discussion_conclusion", {
    "results": ITEM_SCHEMA, "discussion": ITEM_SCHEMA, "conclusion": ITEM_SCHEMA,
    "abstract_conclusion_alignment": ITEM_SCHEMA, "reference_use": ITEM_SCHEMA,
    "major_red_flags": {"type": "array", "items": {"type": "string"}},
    "editorial_recommendation": {"type": "string", "enum": ["Proceed to editorial review", "Needs author correction", "Major concern"]}},
    ["results", "discussion", "conclusion", "abstract_conclusion_alignment", "reference_use", "major_red_flags", "editorial_recommendation"])


def _groq_json_call(client: Groq, schema: Dict[str, Any], prompt: str, max_tokens: int = 1200) -> Dict[str, Any]:
    """Call Groq Structured Outputs with enough budget for GPT-OSS JSON generation."""
    request = dict(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a cautious academic journal pre-screening assistant. "
                    "Assess only evidence supplied. Never invent facts, citations, "
                    "authors, reviewers, institutions, sample sizes, results, "
                    "statistical tests, ethics approvals, or research questions. "
                    "If evidence is missing, use NOT ASSESSABLE. Do not rewrite "
                    "manuscript text. Keep findings, evidence, and actions concise. "
                    "Return only the JSON required by the schema."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0,
        reasoning_effort="low",
        include_reasoning=False,
        max_completion_tokens=max_tokens,
        response_format={"type": "json_schema", "json_schema": schema},
    )

    try:
        response = client.chat.completions.create(**request)
    except Exception as exc:
        # If GPT-OSS exhausts its completion budget before constrained JSON is
        # complete, retry once with a larger completion budget.
        msg = str(exc).lower()
        if "max completion tokens" not in msg and "json_validate_failed" not in msg:
            raise
        request["max_completion_tokens"] = max(2400, max_tokens * 2)
        response = client.chat.completions.create(**request)

    return json.loads(response.choices[0].message.content or "{}")


def _ai_not_assessed(reason: str) -> Dict[str, Any]:
    item = {"status": "NOT ASSESSABLE", "finding": reason, "evidence": "", "action": "Review this item manually."}
    keys = ["research_question","abstract","introduction_novelty","methodology","statistics","ethics_reproducibility","results","discussion","conclusion","abstract_conclusion_alignment","reference_use"]
    return {"summary": reason, "research_area": "Not assessed", "keywords": [], **{k: item.copy() for k in keys}, "major_red_flags": [reason], "editorial_recommendation": "Needs author correction"}


def _compact_rule_summary(checks: List[Dict[str, Any]]) -> str:
    return "\n".join(f"- {c['requirement']}: {c['status']} — {c['evidence']}" for c in checks)


def run_ai_analysis(text: str, deterministic_checks: List[Dict[str, Any]], client: Groq) -> Dict[str, Any]:
    """Run three focused AI assessments instead of one oversized manuscript prompt."""
    abstract = extract_abstract(text); intro = extract_section_text(text, "introduction")
    methods = extract_section_text(text, "materials and methods"); rd = extract_section_text(text, "results and discussion")
    conclusion = extract_section_text(text, "conclusions"); refs = extract_section_text(text, "references")
    rules = _compact_rule_summary(deterministic_checks)
    try:
        prompt1 = f"""MRJ PRE-SCREEN: RESEARCH QUESTION, ABSTRACT, INTRODUCTION

MRJ CHECKS:
{rules[:3500]}

ABSTRACT:
{abstract[:2600]}

INTRODUCTION:
{intro[:3800]}

Assess: explicit research question/objective; abstract coverage of background, methods, results and conclusion; and whether the introduction establishes a supported gap/objective/novelty. Extract a concise research area and useful search keywords. Never infer missing facts. Use NOT ASSESSABLE when evidence is insufficient."""
        a = _groq_json_call(client, AI_CALL_1_SCHEMA, prompt1, 1200)
        prompt2 = f"""MRJ PRE-SCREEN: METHODOLOGY, STATISTICS, ETHICS

MRJ CHECKS:
{rules[:3500]}

MATERIALS AND METHODS:
{methods[:5000]}

Assess methodology completeness/reproducibility; appropriateness and reporting of statistics/data analysis; and ethics/consent/animal/reproducibility information where relevant. Do not invent missing sample sizes, tests or approvals. Use NOT ASSESSABLE when evidence is insufficient."""
        b = _groq_json_call(client, AI_CALL_2_SCHEMA, prompt2, 1200)
        prompt3 = f"""MRJ PRE-SCREEN: RESULTS, DISCUSSION, CONCLUSION, REFERENCES

ABSTRACT:
{abstract[:2200]}

RESULTS AND DISCUSSION:
{rd[:5200]}

CONCLUSIONS:
{conclusion[:2200]}

REFERENCES:
{refs[:2400]}

Assess whether results answer the objective; discussion interprets rather than merely repeats; conclusion is supported and appropriately limited; abstract/conclusion are aligned; and citations/references are used coherently. Do not verify reference facts from memory. List only evidence-supported major red flags. Give a pre-screening recommendation, not an acceptance/rejection decision."""
        c = _groq_json_call(client, AI_CALL_3_SCHEMA, prompt3, 1400)
        out = {}; out.update(a); out.update(b); out.update(c)
        assess_keys = ["research_question","abstract","introduction_novelty","methodology","statistics","ethics_reproducibility","results","discussion","conclusion","abstract_conclusion_alignment","reference_use"]
        concern_count = sum(out.get(k, {}).get("status") in ("CONCERN", "MAJOR CONCERN") for k in assess_keys)
        out["summary"] = (f"Focused AI assessment completed across 11 scientific areas; {concern_count} area(s) were flagged as CONCERN or MAJOR CONCERN. Objective MRJ compliance remains based on deterministic checks." if concern_count else "Focused AI assessment completed across 11 scientific areas. No CONCERN or MAJOR CONCERN item was flagged from the supplied evidence; this does not replace peer review.")
        return out
    except Exception as exc:
        msg = str(exc)
        if "413" in msg or "tokens per minute" in msg.lower():
            return _ai_not_assessed("Groq rate limit was reached. Deterministic MRJ checks remain valid and the manuscript was not modified.")
        return _ai_not_assessed(f"Structured AI assessment failed: {msg}")


def _format_ai_item(item: Dict[str, Any]) -> str:
    return "\n".join([f"Status: {item.get('status', 'NOT ASSESSABLE')}", f"Finding: {item.get('finding', '')}", f"Evidence: {item.get('evidence', '') or 'Not supplied.'}", f"Action: {item.get('action', '') or 'Manual review required.'}"])


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
        if p is None:
            continue
        if re.match(r"(?i)^\s*abstract\s*:?", p.text.strip()):
            abstract_idx = i
            break

    # 3. In the front matter, preserve the title but remove the author block.
    #    We keep the first substantial paragraph as the title and remove everything
    #    else before Abstract. This avoids trying to guess author names with AI.
    if abstract_idx is not None:
        nonempty_before = [
            (i, p) for i, p in enumerate(body_paragraphs[:abstract_idx])
            if p is not None and p.text.strip()
        ]

        if nonempty_before:
            title_idx = nonempty_before[0][0]
            for i, p in enumerate(body_paragraphs[:abstract_idx]):
                if i == title_idx:
                    continue
                if p is not None and p.text.strip():
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

    paragraphs = [p for p in doc.paragraphs if p is not None]
    for i, p in enumerate(paragraphs):
        if p._element is None or p._element.getparent() is None:
            continue
        heading = normalize(p.text).lower().rstrip(":")
        if heading not in removable_section_starts:
            continue

        # Delete heading and following paragraphs until the next obvious heading.
        delete_paragraph(p)
        for q in paragraphs[i + 1:]:
            if q is None or q._element is None or q._element.getparent() is None:
                continue
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

    doc.add_heading("Assessment overview", level=2)
    doc.add_paragraph(ai_data.get("summary", "Not available."))
    doc.add_heading("Research area and reviewer-search keywords", level=2)
    doc.add_paragraph(f"Research area: {ai_data.get('research_area', 'Not assessed')}")
    doc.add_paragraph("Keywords: " + (", ".join(ai_data.get("keywords", [])) or "Not assessed"))
    groups=[("Research question / objective","research_question"),("Abstract","abstract"),("Introduction and novelty","introduction_novelty"),("Methodology","methodology"),("Statistics / data analysis","statistics"),("Ethics and reproducibility","ethics_reproducibility"),("Results","results"),("Discussion","discussion"),("Conclusion","conclusion"),("Abstract–conclusion alignment","abstract_conclusion_alignment"),("Reference use","reference_use")]
    doc.add_heading("Scientific pre-screening assessment", level=2)
    for label,key in groups:
        doc.add_heading(label, level=3); doc.add_paragraph(_format_ai_item(ai_data.get(key, {})))
    doc.add_heading("Potential major red flags", level=2)
    concerns=ai_data.get("major_red_flags", [])
    if concerns:
        for concern in concerns: doc.add_paragraph(concern, style="List Bullet")
    else: doc.add_paragraph("No major red flags were identified from the supplied evidence.")
    doc.add_heading("Editorial recommendation", level=2)
    doc.add_paragraph(ai_data.get("editorial_recommendation", "Not assessed"))
    doc.add_paragraph("This is a pre-screening aid only. Final editorial and peer-review decisions remain human responsibilities.")

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

    st.subheader("2. AI Scientific Pre-Screening")
    st.write(ai_data.get("summary", ""))
    st.write("**Research area:**", ai_data.get("research_area", ""))
    st.write("**Reviewer-search keywords:**", ", ".join(ai_data.get("keywords", [])) or "Not assessed")
    groups=[("Research question / objective","research_question"),("Abstract","abstract"),("Introduction and novelty","introduction_novelty"),("Methodology","methodology"),("Statistics / data analysis","statistics"),("Ethics and reproducibility","ethics_reproducibility"),("Results","results"),("Discussion","discussion"),("Conclusion","conclusion"),("Abstract–conclusion alignment","abstract_conclusion_alignment"),("Reference use","reference_use")]
    for label,key in groups:
        item=ai_data.get(key,{})
        with st.expander(f"{label} — {item.get('status','NOT ASSESSABLE')}"):
            st.write("**Finding:**", item.get("finding", "")); st.write("**Evidence:**", item.get("evidence", "") or "Not supplied."); st.write("**Action:**", item.get("action", "") or "Manual review required.")
    concerns=ai_data.get("major_red_flags", [])
    if concerns:
        st.warning("Potential major red flags")
        for item in concerns: st.write(f"- {item}")
    st.write("**Editorial pre-screening recommendation:**", ai_data.get("editorial_recommendation", "Not assessed"))

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
