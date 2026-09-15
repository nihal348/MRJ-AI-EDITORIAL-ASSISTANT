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
# CONFIGURATION & MRJ TEMPLATE CONSTANTS
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
    "Declaration on AI Usage",
    "References",
    "Ethics Statement",
]

CORE_SECTIONS = {
    "Abstract",
    "Introduction",
    "Materials and Methods",
    "Results and Discussion",
    "Conclusions",
    "Multidisciplinary Domains",
    "References",
}

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
    "Declaration on AI Usage": r"(?i)^\s*(?:declaration\s+on\s+ai(?:\s+usage)?|generative\s+ai\s+statement|ai\s+usage|declaration\s+on\s+artificial\s+intelligence)\s*:?$",
    "References": r"(?i)^\s*(?:references?|bibliography|literature\s+cited)\s*:?$",
    "Ethics Statement": r"(?i)^\s*(?:ethics\s+statement|ethical\s+approval|ethics\s+approval|institutional\s+review\s+board|irb\s+statement)\s*:?$",
}

THESIS_SUBHEADING_PATTERNS = [
    r"(?i)\b(?:review of (?:related )?literature|literature review)\b",
    r"(?i)\b(?:statement of (?:the )?problem|problem statement)\b",
    r"(?i)\b(?:hypotheses|hypothesis development)\b",
    r"(?i)\b(?:research questions?)\b",
    r"(?i)\b(?:objectives of (?:the )?study)\b",
    r"(?i)\b(?:delimitations?|scope and delimitations?)\b",
    r"(?i)\b(?:significance of (?:the )?study)\b",
    r"(?i)\b(?:conceptual framework)\b",
    r"(?i)\b(?:operational definitions?)\b",
]

TEMPLATE_PLACEHOLDER_PATTERNS = [
    r"(?i)\b(?:issn\s*[:\-]?\s*xxxx-xxxx|issn\s+xxxx)\b",
    r"(?i)\b(?:doi\s*[:\-]?\s*10\.xxxx[^\s]*)\b",
    r"(?i)\b(?:volume\s*xx|vol\.\s*xx|issue\s*xx)\b",
    r"(?i)\b(?:received\s*:\s*date|accepted\s*:\s*date|published\s*:\s*date)\b",
    r"(?i)\[insert\s+(?:figure|table|author|text|affiliation)\b.*?\]",
]

NORTHEAST_STATES = {
    "assam", "arunachal pradesh", "manipur", "meghalaya", "mizoram", "nagaland", "sikkim", "tripura"
}

NORTHEAST_INSTITUTION_TERMS = [
    "iit guwahati", "tezu university", "tezpur university", "nit silchar", "assam university",
    "gauhati university", "cotton university", "dibrugarh university", "nehu",
    "north-eastern hill university", "nit agartala", "manipur university",
    "mizoram university", "nagaland university", "tripura university", "rajiv gandhi university",
    "sikkim university", "arunachal university"
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
    clean = re.sub(r"^```(?:json)?\s*", "", raw_resp.strip(), flags=re.MULTILINE)
    clean = re.sub(r"```\s*$", "", clean.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(clean)
    except Exception:
        match = re.search(r"(\{.*\})", clean, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        raise ValueError("Could not extract valid JSON from LLM audit response.")


# ============================================================
# DETERMINISTIC MRJ TEMPLATE PRE-SCANNER
# ============================================================

def preaudit_mrj_template(text: str) -> Tuple[List[Dict[str, Any]], List[str], List[str], Dict[str, Any]]:
    lines = text.splitlines()
    detected_map = {}
    detected_thesis_headers = []
    detected_placeholders = []
    template_observations = {
        "abstract_word_count": 0,
        "keywords_detected": [],
        "multidisciplinary_domains_found": False,
        "multidisciplinary_domains_count": 0,
        "declaration_ai_usage_found": False,
        "funding_statement_found": False,
        "conflicts_statement_found": False,
        "square_bracket_citations_found": False,
        "improper_p_values_detected": [],
    }

    # Citations check
    citation_matches = re.findall(r"\[\d+(?:[\–\-–,]\s*\d+)*\]", text)
    template_observations["square_bracket_citations_found"] = len(citation_matches) > 0

    # P-value regex scan (flagging p=0.00 or p=0)
    p_zero_matches = re.findall(r"(?i)\bp\s*(?:=|is)\s*0(?:\.0+)?(?!\d)", text)
    if p_zero_matches:
        template_observations["improper_p_values_detected"] = list(set(p_zero_matches))

    # Placeholder detection
    for ph_pattern in TEMPLATE_PLACEHOLDER_PATTERNS:
        matches = re.findall(ph_pattern, text)
        for m in matches:
            clean_m = normalize(m)
            if clean_m and clean_m not in detected_placeholders:
                detected_placeholders.append(clean_m)

    for i, line in enumerate(lines):
        clean = normalize(line)
        if not clean or len(clean) > 85:
            continue

        # Detect sections
        for sec_name, pattern in SECTION_PATTERNS.items():
            if sec_name in detected_map:
                continue
            if re.match(pattern, clean):
                quote = ""
                for nxt in lines[i + 1: i + 6]:
                    cleaned_nxt = normalize(nxt)
                    if cleaned_nxt and not any(re.match(p, cleaned_nxt) for p in SECTION_PATTERNS.values()):
                        quote = cleaned_nxt[:140]
                        break
                detected_map[sec_name] = {
                    "heading": clean,
                    "first_line_quote": quote or "Section heading identified with inline content.",
                    "line_num": i,
                }

        # Check for thesis subheadings
        for t_pattern in THESIS_SUBHEADING_PATTERNS:
            if re.match(t_pattern, clean) and clean not in detected_thesis_headers:
                detected_thesis_headers.append(clean)

    # Implicit abstract detection
    if "Abstract" not in detected_map:
        intro_line = detected_map.get("Introduction", {}).get("line_num", len(lines))
        search_limit = min(len(lines), intro_line, 45)
        for idx in range(search_limit):
            l = lines[idx].strip()
            if (
                len(l.split()) >= 25
                and not re.search(r"(?i)@|department\b|university\b|institute\b|received:|accepted:|doi\.org|issn|vol\.\s*\d+", l)
                and not any(re.match(p, l) for p in SECTION_PATTERNS.values())
            ):
                detected_map["Abstract"] = {
                    "heading": "Implicit Abstract (Heading unformatted/omitted)",
                    "first_line_quote": l[:140],
                    "line_num": idx,
                }
                template_observations["abstract_word_count"] = word_count(l)
                break
    else:
        abs_line = detected_map["Abstract"]["line_num"]
        abs_text_parts = []
        for nxt in lines[abs_line + 1: abs_line + 18]:
            if re.match(r"(?i)^\s*(?:keywords?|1\.?\s+introduction)\b", nxt.strip()):
                break
            abs_text_parts.append(nxt)
        template_observations["abstract_word_count"] = word_count(" ".join(abs_text_parts))

    # Keywords detection
    m_kw = re.search(r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:(?:1\.?\s+)?introduction|materials|abstract)\b|$)", text)
    if m_kw:
        raw_kw = m_kw.group(1).strip().splitlines()[0]
        template_observations["keywords_detected"] = [
            normalize(k) for k in re.split(r"[;,]", raw_kw) if len(normalize(k)) > 1
        ]

    # Multidisciplinary Domains verification
    m_dom = re.search(
        r"(?is)\bmultidisciplinary\s+domains?\b.*?(?:this\s+research\s+covers\s+the\s+domains\s*:\s*|\(a\))(.*?)(?=\n\s*(?:funding|acknowledg|conflicts|declaration|references)\b|$)",
        text,
    )
    if m_dom:
        template_observations["multidisciplinary_domains_found"] = True
        domain_items = re.findall(r"\([a-z]\)\s*([^,;.]+)", m_dom.group(0), flags=re.I)
        template_observations["multidisciplinary_domains_count"] = len(domain_items)

    # AI Declaration check
    if re.search(r"(?is)\bdeclaration\s+on\s+ai\s+usage\b|prepared\s+without\s+the\s+use\s+of\s+ai\s+tools|ai\s+tools\b", text):
        template_observations["declaration_ai_usage_found"] = True

    # Assemble section pre-audit
    preaudited = []
    for sec in REQUIRED_TEMPLATE_SECTIONS:
        if sec in detected_map:
            info = detected_map[sec]
            preaudited.append({
                "section_name": sec,
                "detected_heading": info["heading"],
                "first_line_quote": info["first_line_quote"],
                "status_hint": "PASS",
            })
        else:
            preaudited.append({
                "section_name": sec,
                "detected_heading": "Section heading not identified",
                "first_line_quote": "N/A",
                "status_hint": "FAIL" if sec in CORE_SECTIONS else "WARN",
            })

    return preaudited, detected_thesis_headers, detected_placeholders, template_observations


def extract_text_and_assets(uploaded_file) -> Tuple[str, Dict[str, Any]]:
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
                            "has_visual": (asset_meta["total_images"] > 0 if "Fig" in label else asset_meta["total_tables"] > 0),
                        })

        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))

        return "\n".join(parts), asset_meta

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
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
                                "has_visual": (images_on_page > 0 if "Fig" in label else tables_on_page > 0) or (asset_meta["total_images"] > 0 if "Fig" in label else asset_meta["total_tables"] > 0),
                            })

        return "\n".join(pages), asset_meta

    raise ValueError("Unsupported format. Please upload a .docx or .pdf file.")


# ============================================================
# AUDIT JSON SCHEMA SPECIFICATION (5-PASS ENGINE)
# ============================================================

AUDIT_STRICT_SCHEMA = {
    "name": "manuscript_editorial_audit",
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
                        "enum": [
                            "Accept as is",
                            "Accept with Minor Revisions",
                            "Major Revisions",
                            "Reject",
                        ],
                    },
                    "executive_summary": {"type": "string"},
                    "actionable_revision_requirements": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
                "required": ["decision", "executive_summary", "actionable_revision_requirements"],
                "additionalProperties": False,
            },
            "pass_1_cross_section_integrity": {
                "type": "object",
                "properties": {
                    "abstract_claimed_items": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "abstract_body_discrepancies": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
                "required": ["abstract_claimed_items", "abstract_body_discrepancies"],
                "additionalProperties": False,
            },
            "pass_2_template_and_layout": {
                "type": "object",
                "properties": {
                    "journal_style_compliance": {
                        "type": "string",
                        "enum": ["PASS", "FAIL"],
                    },
                    "unwanted_thesis_subheadings": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "template_placeholders_detected": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "multidisciplinary_domains_compliance": {
                        "type": "string",
                        "enum": ["PASS", "FAIL"],
                    },
                    "abstract_quality_and_word_count": {"type": "string"},
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
                                    "enum": ["PASS", "WARN", "FAIL", "NOT EVALUATED"],
                                },
                            },
                            "required": ["section_name", "detected_heading", "first_line_quote", "status"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": [
                    "journal_style_compliance",
                    "unwanted_thesis_subheadings",
                    "template_placeholders_detected",
                    "multidisciplinary_domains_compliance",
                    "abstract_quality_and_word_count",
                    "structural_section_checks",
                ],
                "additionalProperties": False,
            },
            "pass_3_statistical_and_mathematical_rigor": {
                "type": "object",
                "properties": {
                    "p_value_reporting_issues": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "pseudoreplication_flags": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "statistical_test_completeness": {"type": "string"},
                },
                "required": [
                    "p_value_reporting_issues",
                    "pseudoreplication_flags",
                    "statistical_test_completeness",
                ],
                "additionalProperties": False,
            },
            "pass_4_algorithmic_reproducibility": {
                "type": "object",
                "properties": {
                    "is_computational_or_algorithm_paper": {"type": "boolean"},
                    "algorithmic_edge_cases_notes": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "software_dependency_versions": {"type": "string"},
                    "benchmarks_and_scalability": {"type": "string"},
                },
                "required": [
                    "is_computational_or_algorithm_paper",
                    "algorithmic_edge_cases_notes",
                    "software_dependency_versions",
                    "benchmarks_and_scalability",
                ],
                "additionalProperties": False,
            },
            "pass_5_literal_text_and_captions": {
                "type": "object",
                "properties": {
                    "enumeration_figure_inconsistencies": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "taxonomic_formatting_flags": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "typographical_and_spacing_issues": {
                        "type": "array",
                        "items": {"type": "string"},
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
                                    "enum": ["PASS", "FAIL", "NOT EVALUATED"],
                                },
                            },
                            "required": ["label", "placement", "visual_present", "status"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": [
                    "enumeration_figure_inconsistencies",
                    "taxonomic_formatting_flags",
                    "typographical_and_spacing_issues",
                    "visual_asset_audit",
                ],
                "additionalProperties": False,
            },
        },
        "required": [
            "manuscript_title",
            "editorial_verdict",
            "pass_1_cross_section_integrity",
            "pass_2_template_and_layout",
            "pass_3_statistical_and_mathematical_rigor",
            "pass_4_algorithmic_reproducibility",
            "pass_5_literal_text_and_captions",
        ],
        "additionalProperties": False,
    },
}


# ============================================================
# AUDIT ENGINE (GROQ API IMPLEMENTATION OF MANDATORY PASSES)
# ============================================================

def run_editorial_audit(raw_text: str, asset_meta: Dict[str, Any], client: Groq) -> Dict[str, Any]:
    preaudited, detected_thesis, detected_placeholders, template_obs = preaudit_mrj_template(raw_text)

    prompt = f"""You are an expert Lead Academic Editor and Manuscript Quality Auditor. Pre-screen this submission for the Multidisciplinary Research Journal (MRJ) by performing a strict 5-pass audit.

MANDATORY AUDIT PASSES:

PASS 1: CROSS-SECTION CONTENT INTEGRITY & DISCREPANCY AUDIT
- Extract every experiment, algorithm, dataset, or organism claimed in the ABSTRACT.
- Cross-examine the METHODS and RESULTS sections to verify if each claimed study is explicitly present with empirical data.
- Detail any discrepancies or phantom claims under 'abstract_body_discrepancies'.

PASS 2: TEMPLATE, LAYOUT & PLACEHOLDER AUDIT
- Verify MRJ required sections: Abstract, Introduction, Materials and Methods, Results and Discussion, Conclusions, Multidisciplinary Domains, Funding, Acknowledgments, Conflicts of Interest, Declaration on AI Usage, References.
- Check for unwanted thesis subheadings (e.g., 'Review of Related Literature', 'Statement of the Problem', 'Hypotheses', 'Research Questions', 'Objectives', 'Significance of the study', 'Delimitations').
- Detect leftover template placeholders (e.g., 'ISSN XXXX-XXXX', 'DOI 10.XXXX/...', 'Volume XX', 'Received: date').
- Verify Multidisciplinary Domains: must cover >= 2 domains with exact formulation 'This research covers the domains: (a) ..., (b) ...'.
- Check Abstract: single paragraph, <= 200 words, quantitative metrics instead of generic qualitative claims.

PASS 3: STATISTICAL & MATHEMATICAL RIGOR AUDIT
- Check for improper p-value reporting (e.g., 'p = 0.00' or 'p = 0' instead of 'p < 0.001').
- Check for potential pseudoreplication (e.g., treating image pixels, technical replicates, or subsamples as independent observational units).
- Verify all statistical tests state observational units, sample sizes, and degrees of freedom rationale.

PASS 4: ALGORITHMIC & METHODOLOGICAL REPRODUCIBILITY AUDIT
- If this is a software, algorithm, or data tool paper, audit whether the following edge cases are documented:
  * Handling of ambiguous bases (e.g., 'N')
  * Lowercase vs. uppercase sequences/strings
  * Overlapping vs. non-overlapping k-mers
  * Reverse complement handling
  * Multi-chromosomal / multi-contig handling
  * Exact software dependency version numbers (e.g., Python v3.10, OpenCV v4.5)
  * Memory/runtime benchmarks for scalability claims.
- If not a software paper, indicate is_computational_or_algorithm_paper: false and note methodological software reproducibility.

PASS 5: LITERAL TEXT & FIGURE CAPTION AUDIT
- Audit inline text enumerations against figures (check that lists match exact counts without omissions or duplicates).
- Verify species names for binomial italicization (e.g., *E. coli*, *S. pneumoniae*, *M. tuberculosis*).
- Detect punctuation anomalies, missing spaces after punctuation, or concatenated words.

PRE-SCANNED DETERMINISTIC SIGNALS:
- Pre-scanned Sections: {json.dumps(preaudited, indent=2)}
- Unwanted Thesis Subheadings: {json.dumps(detected_thesis, indent=2)}
- Template Placeholders Detected: {json.dumps(detected_placeholders, indent=2)}
- Abstract Word Count: ~{template_obs['abstract_word_count']} words
- Keywords Detected: {json.dumps(template_obs['keywords_detected'])}
- Multidisciplinary Domains Detected: {template_obs['multidisciplinary_domains_count']} (Found: {template_obs['multidisciplinary_domains_found']})
- Declaration on AI Usage Present: {template_obs['declaration_ai_usage_found']}
- Regex Detected Improper P-values: {json.dumps(template_obs['improper_p_values_detected'])}
- Visual Asset Captions: {json.dumps(asset_meta['detected_captions'], indent=2)}

MANUSCRIPT EXCERPT:
--- START OF MANUSCRIPT ---
{raw_text[:15000]}
--- END OF MANUSCRIPT ---

Output strictly valid JSON complying with the requested schema."""

    messages = [
        {
            "role": "system",
            "content": "You are a rigorous Lead Academic Editor. Pre-screen submissions according to the 5-pass audit rules and return strictly valid JSON matching the schema.",
        },
        {"role": "user", "content": prompt},
    ]

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

    resp = client.chat.completions.create(
        model=FALLBACK_MODEL,
        messages=messages,
        temperature=0.0,
        max_completion_tokens=4000,
        response_format={"type": "json_object"},
    )
    return clean_json_response(resp.choices[0].message.content or "{}")


# ============================================================
# OPENALEX REVIEWER DISCOVERY
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    res = requests.get(url, params=params, timeout=25)
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


def search_openalex_by_query(query: str, region: str, matched_topic: str, mailto: str = "") -> List[Dict[str, Any]]:
    if not query.strip():
        return []
    try:
        data = openalex_get(
            OPENALEX_URL,
            {"search": query, "per-page": 25, "sort": "publication_year:desc"},
            mailto=mailto,
        )
    except Exception:
        return []

    candidates = {}
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
                    "match_type": f"Specialist in {matched_topic}",
                    "recent_pubs": [],
                }
                if not candidate_region_match(candidate, region):
                    continue

                if author_id not in candidates:
                    candidates[author_id] = candidate
                candidates[author_id]["recent_pubs"].append({"year": year, "title": title, "doi": doi})

    return list(candidates.values())


def reviewer_discovery_report(
    title: str,
    extracted_keywords: List[str],
    mailto: str = "",
    max_per_region: int = 4,
) -> Dict[str, List[Dict[str, Any]]]:
    stopwords = {"a", "an", "the", "and", "or", "in", "on", "at", "to", "for", "with", "of", "by", "from", "using", "study", "analysis", "approach", "based"}
    title_terms = [w for w in re.findall(r"\b[A-Za-z]{3,}\b", title) if w.lower() not in stopwords]

    search_queries = []
    if title_terms:
        search_queries.append((" ".join(title_terms[:3]), "Manuscript Title Phrase"))
    for kw in extracted_keywords[:3]:
        if kw and len(kw) > 2:
            search_queries.append((kw, kw))
    for term in title_terms[:2]:
        if term not in [sq[0] for sq in search_queries]:
            search_queries.append((term, term))

    output = {}
    for region in ["India", "Northeast India", "Assam"]:
        region_candidates = {}
        for q_str, q_label in search_queries:
            if len(region_candidates) >= max_per_region:
                break
            found = search_openalex_by_query(q_str, region, matched_topic=q_label, mailto=mailto)
            for cand in found:
                aid = cand["author_id"]
                if aid not in region_candidates:
                    region_candidates[aid] = cand
                else:
                    region_candidates[aid]["recent_pubs"].extend(cand["recent_pubs"])

        results = list(region_candidates.values())
        results.sort(key=lambda x: len(x["recent_pubs"]), reverse=True)

        if len(results) < 2 and region in ["Assam", "Northeast India"]:
            fallback_kw = extracted_keywords[0] if extracted_keywords else (title_terms[0] if title_terms else "research")
            hub_query = f"IIT Guwahati {fallback_kw}" if region == "Assam" else f"Tezpur University {fallback_kw}"
            hub_found = search_openalex_by_query(hub_query, region, matched_topic=f"Regional Hub ({fallback_kw})", mailto=mailto)
            for c in hub_found:
                if c["author_id"] not in region_candidates:
                    results.append(c)

        output[region] = results[:max_per_region]

    return output


# ============================================================
# MARKDOWN REPORT GENERATOR (STRICT PROMPT TEMPLATE FORMAT)
# ============================================================

def generate_markdown_audit_report(audit: Dict[str, Any], reviewers: Dict[str, List[Dict[str, Any]]]) -> str:
    verd = audit.get("editorial_verdict", {})
    p1 = audit.get("pass_1_cross_section_integrity", {})
    p2 = audit.get("pass_2_template_and_layout", {})
    p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
    p4 = audit.get("pass_4_algorithmic_reproducibility", {})
    p5 = audit.get("pass_5_literal_text_and_captions", {})

    # Actionable requirements list
    reqs = verd.get("actionable_revision_requirements", [])
    reqs_md = "\n".join([f"- {r}" for r in reqs]) if reqs else "- No mandatory revisions requested."

    # Subheadings & Placeholders
    unwanted = p2.get("unwanted_thesis_subheadings", [])
    unwanted_str = ", ".join(unwanted) if unwanted else "None detected"

    placeholders = p2.get("template_placeholders_detected", [])
    placeholders_str = ", ".join(placeholders) if placeholders else "None"

    # Discrepancies
    disc = p1.get("abstract_body_discrepancies", [])
    disc_str = "\n".join([f"- {d}" for d in disc]) if disc else "- No phantom claims or body discrepancies detected."

    # P-values and Pseudoreplication
    p_issues = p3.get("p_value_reporting_issues", [])
    p_issues_str = "; ".join(p_issues) if p_issues else "None (properly formatted or p < 0.001 used)"

    pseudo = p3.get("pseudoreplication_flags", [])
    pseudo_str = "; ".join(pseudo) if pseudo else "None identified (sampling units accounted for)"

    edge_cases = p4.get("algorithmic_edge_cases_notes", [])
    edge_cases_str = "; ".join(edge_cases) if edge_cases else "None identified"
    if p4.get("software_dependency_versions"):
        edge_cases_str += f" | Versions: {p4.get('software_dependency_versions')}"

    # Literal text & captions
    enum_issues = p5.get("enumeration_figure_inconsistencies", [])
    enum_str = "; ".join(enum_issues) if enum_issues else "Consistent with figures"

    taxa = p5.get("taxonomic_formatting_flags", [])
    typos = p5.get("typographical_and_spacing_issues", [])
    taxa_str = "; ".join(taxa + typos) if (taxa or typos) else "Standard binomial nomenclature and formatting"

    # Reviewer section
    rev_lines = []
    for reg, cands in reviewers.items():
        rev_lines.append(f"#### {reg}")
        if not cands:
            rev_lines.append("- No verified OpenAlex profiles identified.")
        for c in cands:
            pubs_count = len(c.get("recent_pubs", []))
            rev_lines.append(f"- **{c.get('name')}** ({c.get('institution')}): {c.get('match_type')} [{pubs_count} indexed works]")

    reviewers_md = "\n".join(rev_lines)

    report_md = f"""# MRJ Academic Pre-Screening & Editorial Audit Report

**Manuscript Title:** {audit.get('manuscript_title', 'Not specified')}
**Decision:** {verd.get('decision', 'Major Revisions')}

### Executive Summary & Actionable Revision Requirements
{verd.get('executive_summary', '')}

{reqs_md}

### 1. Template Compliance & Structural Audit
- **Journal Style Compliance:** {p2.get('journal_style_compliance', 'FAIL')}
- **Unwanted Thesis Subheadings:** {unwanted_str}
- **Template Placeholders Detected:** {placeholders_str}
- **Multidisciplinary Domains Statement:** {p2.get('multidisciplinary_domains_compliance', 'FAIL')}
- **Abstract Quality & Word Count:** {p2.get('abstract_quality_and_word_count', 'N/A')}

### 2. Cross-Section Consistency & Phantom Claim Audit
- **Abstract vs. Body Discrepancies:**
{disc_str}

### 3. Statistical, Mathematical & Methodological Rigor
- **P-Value Reporting Issues:** {p_issues_str}
- **Pseudoreplication / Sampling Unit Flags:** {pseudo_str}
- **Algorithmic Reproducibility Edge Cases:** {edge_cases_str}

### 4. Text-Level & Technical Accuracy Flags
- **Enumeration / Text-Figure Inconsistencies:** {enum_str}
- **Taxonomic Formatting & Grammatical Typos:** {taxa_str}

### 5. Verified Relevant Reviewer Discovery
{reviewers_md}
"""
    return report_md


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

def generate_docx_report(
    audit: Dict[str, Any],
    reviewers: Dict[str, List[Dict[str, Any]]],
    template_obs: Dict[str, Any],
) -> bytes:
    doc = Document()
    doc.add_heading("MRJ Academic Pre-Screening & Editorial Audit Report", 0)
    doc.add_paragraph(f"Manuscript Title: {audit.get('manuscript_title', 'Not specified')}")

    verd = audit.get("editorial_verdict", {})
    doc.add_paragraph(f"Decision: {verd.get('decision', 'Under Review')}")
    doc.add_paragraph(f"Executive Summary: {verd.get('executive_summary', '')}")

    reqs = verd.get("actionable_revision_requirements", [])
    if reqs:
        doc.add_heading("Actionable Revision Requirements for Authors:", level=2)
        for r in reqs:
            doc.add_paragraph(f"• {r}")

    # Pass 1
    doc.add_heading("1. Cross-Section Content Integrity & Discrepancy Audit", level=1)
    p1 = audit.get("pass_1_cross_section_integrity", {})
    disc = p1.get("abstract_body_discrepancies", [])
    if disc:
        doc.add_paragraph("Abstract vs. Body Discrepancies Identified:")
        for d in disc:
            doc.add_paragraph(f"• {d}")
    else:
        doc.add_paragraph("Abstract claims are fully corroborated by empirical data in Methods and Results.")

    # Pass 2
    doc.add_heading("2. Template Compliance & Structural Audit", level=1)
    p2 = audit.get("pass_2_template_and_layout", {})
    doc.add_paragraph(f"Journal Style Compliance: {p2.get('journal_style_compliance', 'N/A')}")
    doc.add_paragraph(f"Multidisciplinary Domains Statement: {p2.get('multidisciplinary_domains_compliance', 'N/A')}")
    doc.add_paragraph(f"Abstract Word Count & Quality: {p2.get('abstract_quality_and_word_count', 'N/A')}")

    placeholders = p2.get("template_placeholders_detected", [])
    doc.add_paragraph(f"Template Placeholders: {', '.join(placeholders) if placeholders else 'None'}")

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    h = tbl.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Section", "Detected Heading", "First-Line Quote", "Status"
    for s in p2.get("structural_section_checks", []):
        row = tbl.add_row().cells
        row[0].text = s.get("section_name", "")
        row[1].text = s.get("detected_heading", "")
        row[2].text = s.get("first_line_quote", "")
        row[3].text = s.get("status", "")

    # Pass 3
    doc.add_heading("3. Statistical, Mathematical & Methodological Rigor", level=1)
    p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
    p_issues = p3.get("p_value_reporting_issues", [])
    doc.add_paragraph(f"P-Value Reporting Issues: {', '.join(p_issues) if p_issues else 'None detected'}")
    pseudos = p3.get("pseudoreplication_flags", [])
    doc.add_paragraph(f"Pseudoreplication Warnings: {', '.join(pseudos) if pseudos else 'None flagged'}")
    doc.add_paragraph(f"Statistical Completeness: {p3.get('statistical_test_completeness', 'N/A')}")

    # Pass 4
    doc.add_heading("4. Algorithmic Reproducibility & Edge Cases", level=1)
    p4 = audit.get("pass_4_algorithmic_reproducibility", {})
    doc.add_paragraph(f"Computational/Algorithm Manuscript: {'Yes' if p4.get('is_computational_or_algorithm_paper') else 'No'}")
    edge_notes = p4.get("algorithmic_edge_cases_notes", [])
    if edge_notes:
        for en in edge_notes:
            doc.add_paragraph(f"• Edge case note: {en}")
    doc.add_paragraph(f"Software Versions: {p4.get('software_dependency_versions', 'None specified')}")
    doc.add_paragraph(f"Benchmarks / Scalability: {p4.get('benchmarks_and_scalability', 'None specified')}")

    # Pass 5
    doc.add_heading("5. Text-Level Accuracy & Visual Captions", level=1)
    p5 = audit.get("pass_5_literal_text_and_captions", {})
    v_tbl = doc.add_table(rows=1, cols=4)
    v_tbl.style = "Table Grid"
    vh = v_tbl.rows[0].cells
    vh[0].text, vh[1].text, vh[2].text, vh[3].text = "Asset Label", "Placement", "Visual Present", "Status"
    for v in p5.get("visual_asset_audit", []):
        row = v_tbl.add_row().cells
        row[0].text = v.get("label", "")
        row[1].text = v.get("placement", "")
        row[2].text = "Yes" if v.get("visual_present") else "No"
        row[3].text = v.get("status", "")

    # Reviewers
    doc.add_heading("6. Verified Reviewer Discovery (OpenAlex)", level=1)
    for reg, cands in reviewers.items():
        doc.add_heading(reg, level=2)
        for c in cands:
            doc.add_paragraph(f"• {c.get('name')} ({c.get('institution')}): {c.get('match_type')} - {len(c.get('recent_pubs', []))} recent publications")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="MRJ Manuscript Auditor (5-Pass Engine)", page_icon="📄", layout="wide")
st.title("📄 MRJ Manuscript Quality Auditor & Pre-Screening Engine")
st.caption("Comprehensive 5-Pass editorial validation: Template compliance, cross-section consistency, statistical rigor, reproducibility, and OpenAlex reviewer discovery.")

with st.sidebar:
    st.header("⚙️ Editorial Settings")
    groq_api_key = st.text_input("Groq API Key", value=get_secret("GROQ_API_KEY"), type="password")
    openalex_mailto = st.text_input("OpenAlex Mailto Email", value=get_secret("OPENALEX_MAILTO", "editor@mrjournal.org"))
    st.markdown("---")
    st.markdown("**MRJ 5-Pass Audit Core:**")
    st.markdown("1. **Pass 1:** Cross-section integrity & phantom claims.")
    st.markdown("2. **Pass 2:** Template, placeholders & layout.")
    st.markdown("3. **Pass 3:** Statistical rigor & pseudoreplication.")
    st.markdown("4. **Pass 4:** Algorithmic reproducibility & edge cases.")
    st.markdown("5. **Pass 5:** Literal text accuracy & figure alignment.")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file and st.button("🚀 Run Rigorous 5-Pass MRJ Audit", type="primary"):
    if not groq_api_key:
        st.error("Please provide a valid Groq API Key.")
        st.stop()

    try:
        with st.spinner("Extracting manuscript body and scanning template landmarks..."):
            raw_text, asset_meta = extract_text_and_assets(uploaded_file)
            if not raw_text.strip():
                st.error("Could not extract readable text from the uploaded document.")
                st.stop()

        with st.spinner("Executing regex checks on placeholders, p-values, and thesis subheadings..."):
            _, _, _, template_obs = preaudit_mrj_template(raw_text)

        with st.spinner("Running 5-Pass Auditor with Groq AI..."):
            client = Groq(api_key=groq_api_key)
            audit_result = run_editorial_audit(raw_text, asset_meta, client)

        with st.spinner("Identifying relevant domain reviewers via OpenAlex..."):
            detected_title = audit_result.get("manuscript_title", "")
            keywords = template_obs.get("keywords_detected", [])
            reviewers = reviewer_discovery_report(detected_title, keywords, mailto=openalex_mailto)

        with st.spinner("Compiling structured Markdown and DOCX reports..."):
            md_report = generate_markdown_audit_report(audit_result, reviewers)
            orig_bytes = uploaded_file.getvalue()
            blind_bytes = blind_copy_docx(orig_bytes) if uploaded_file.name.endswith(".docx") else orig_bytes
            docx_report = generate_docx_report(audit_result, reviewers, template_obs)

        st.session_state["audit"] = audit_result
        st.session_state["reviewers"] = reviewers
        st.session_state["template_obs"] = template_obs
        st.session_state["md_report"] = md_report
        st.session_state["blind_bytes"] = blind_bytes
        st.session_state["docx_report"] = docx_report
        st.success("5-Pass editorial audit complete.")

    except Exception as e:
        st.error(f"Audit processing error: {e}")


# Display Audit Findings
if "audit" in st.session_state:
    audit = st.session_state["audit"]
    reviewers = st.session_state["reviewers"]
    template_obs = st.session_state["template_obs"]
    md_report = st.session_state["md_report"]
    verd = audit.get("editorial_verdict", {})
    decision = verd.get("decision", "Major Revisions")

    st.markdown("---")
    c1, c2 = st.columns([1, 3])
    with c1:
        if decision == "Accept as is":
            st.success(f"### Verdict:\n**{decision}**")
        elif decision == "Accept with Minor Revisions":
            st.info(f"### Verdict:\n**{decision}**")
        elif decision == "Major Revisions":
            st.warning(f"### Verdict:\n**{decision}**")
        else:
            st.error(f"### Verdict:\n**{decision}**")
    with c2:
        st.subheader(audit.get("manuscript_title", "Untitled Manuscript"))
        st.write(f"**Executive Summary:** {verd.get('executive_summary', '')}")
        reqs = verd.get("actionable_revision_requirements", [])
        if reqs:
            st.markdown("**Actionable Revision Requirements:**")
            for r in reqs:
                st.markdown(f"- ⚠️ {r}")

    t_md, t1, t2, t3, t4, t5, t_rev = st.tabs([
        "📄 Structured Markdown Report",
        "Pass 1: Cross-Section Integrity",
        "Pass 2: Template & Sections",
        "Pass 3: Statistical Rigor",
        "Pass 4: Reproducibility",
        "Pass 5: Text & Captions",
        "👥 Reviewer Discovery",
    ])

    with t_md:
        st.markdown(md_report)

    with t1:
        st.subheader("Pass 1: Cross-Section Content Integrity & Phantom Claim Audit")
        p1 = audit.get("pass_1_cross_section_integrity", {})
        st.markdown("**Experiments/Claims Extracted from Abstract:**")
        for item in p1.get("abstract_claimed_items", []):
            st.markdown(f"- 🔬 {item}")

        discs = p1.get("abstract_body_discrepancies", [])
        if discs:
            st.error("**Discrepancies & Phantom Claims Flagged:**")
            for d in discs:
                st.markdown(f"- ❌ {d}")
        else:
            st.success("✅ No discrepancies detected between Abstract claims and empirical data.")

    with t2:
        st.subheader("Pass 2: Template, Layout & Placeholder Audit")
        p2 = audit.get("pass_2_template_and_layout", {})
        col_p1, col_p2, col_p3 = st.columns(3)
        col_p1.metric("Journal Style Compliance", p2.get("journal_style_compliance", "N/A"))
        col_p2.metric("Multidisciplinary Domains", p2.get("multidisciplinary_domains_compliance", "N/A"))
        col_p3.metric("Abstract Words", f"~{template_obs.get('abstract_word_count', 0)}")

        phs = p2.get("template_placeholders_detected", [])
        if phs:
            st.warning("⚠️ **Template Placeholders Detected:**\n\n" + ", ".join([f"`{p}`" for p in phs]))
        else:
            st.success("✅ No template placeholders identified.")

        unw = p2.get("unwanted_thesis_subheadings", [])
        if unw:
            st.warning("⚠️ **Unwanted Dissertation Subheadings Identified:**\n\n" + ", ".join([f"`{u}`" for u in unw]))
        else:
            st.success("✅ Standard journal section hierarchy maintained.")

        st.dataframe(p2.get("structural_section_checks", []), use_container_width=True, hide_index=True)

    with t3:
        st.subheader("Pass 3: Statistical & Mathematical Rigor Audit")
        p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
        p_issues = p3.get("p_value_reporting_issues", [])
        if p_issues:
            st.warning("⚠️ **Improper P-Value Reporting Flagged (e.g. p = 0.00):**")
            for pi in p_issues:
                st.markdown(f"- {pi}")
        else:
            st.success("✅ P-values properly reported without absolute zero statements.")

        pseudos = p3.get("pseudoreplication_flags", [])
        if pseudos:
            st.warning("⚠️ **Potential Pseudoreplication / Unit Issues:**")
            for ps in pseudos:
                st.markdown(f"- {ps}")
        else:
            st.success("✅ Observational and experimental units appropriately differentiated.")

        st.info(f"**Statistical Test Completeness:** {p3.get('statistical_test_completeness', 'N/A')}")

    with t4:
        st.subheader("Pass 4: Algorithmic & Methodological Reproducibility Audit")
        p4 = audit.get("pass_4_algorithmic_reproducibility", {})
        st.write(f"**Computational/Tool Paper:** {'Yes' if p4.get('is_computational_or_algorithm_paper') else 'No'}")
        edge_cases = p4.get("algorithmic_edge_cases_notes", [])
        if edge_cases:
            st.warning("**Reproducibility & Edge Case Observations:**")
            for ec in edge_cases:
                st.markdown(f"- {ec}")
        else:
            st.success("✅ No major algorithmic reproducibility gaps identified.")

        st.write(f"**Software Dependency Versions:** {p4.get('software_dependency_versions', 'None specified')}")
        st.write(f"**Scalability & Benchmarks:** {p4.get('benchmarks_and_scalability', 'None recorded')}")

    with t5:
        st.subheader("Pass 5: Literal Text & Visual Caption Audit")
        p5 = audit.get("pass_5_literal_text_and_captions", {})
        enums = p5.get("enumeration_figure_inconsistencies", [])
        if enums:
            st.warning("**Enumeration & Figure Discrepancies:**")
            for en in enums:
                st.markdown(f"- {en}")
        else:
            st.success("✅ Textual enumerations and figure elements align.")

        taxa = p5.get("taxonomic_formatting_flags", [])
        if taxa:
            st.warning("**Taxonomic Nomenclature / Formatting Flags:**")
            for tx in taxa:
                st.markdown(f"- {tx}")
        else:
            st.success("✅ Taxonomic formatting compliant.")

        st.dataframe(p5.get("visual_asset_audit", []), use_container_width=True, hide_index=True)

    with t_rev:
        st.subheader("Verified Relevant Reviewer Discovery (OpenAlex)")
        for region, cands in reviewers.items():
            st.markdown(f"### Region: {region}")
            if not cands:
                st.write("No matching candidates located.")
                continue
            r_rows = []
            for c in cands:
                pubs = c.get("recent_pubs", [])
                r_rows.append({
                    "Candidate Name": c.get("name"),
                    "Affiliation": c.get("institution"),
                    "Specialization": c.get("match_type"),
                    "Indexed Publications": len(pubs),
                    "Representative Work": pubs[0]["title"] if pubs else "N/A",
                })
            st.dataframe(r_rows, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📥 Editorial Exports")
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button(
        "📝 Download Markdown Report (.md)",
        data=st.session_state["md_report"],
        file_name="MRJ_Editorial_Audit_Report.md",
        mime="text/markdown",
        use_container_width=True,
    )
    d2.download_button(
        "📄 Download Word Report (.docx)",
        data=st.session_state["docx_report"],
        file_name="MRJ_Editorial_Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    if uploaded_file.name.endswith(".docx"):
        d3.download_button(
            "🙈 Download Blind Reviewer Copy (.docx)",
            data=st.session_state["blind_bytes"],
            file_name="MRJ_Anonymized_Reviewer_Copy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    d4.download_button(
        "💾 Download Audit JSON (.json)",
        data=json.dumps(audit, indent=2),
        file_name="mrj_audit_data.json",
        mime="application/json",
        use_container_width=True,
    )
