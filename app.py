import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
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

DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"
OPENALEX_URL = "https://api.openalex.org/works"
OPENALEX_AUTHOR_URL = "https://api.openalex.org/authors"

MRJ_RULES = {
    "abstract_max_words": 200,
    "abstract_soft_overflow_limit": 220,
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
        "materials and methods", "materials & methods",
        "methods and materials", "methodology", "methods", "research methodology"
    ],
    "results and discussion": [
        "results and discussion", "results & discussion",
        "results", "discussion", "findings and discussion"
    ],
    "conclusions": [
        "conclusions", "conclusion", "summary and conclusions",
        "summary and conclusion", "concluding remarks", "summary"
    ],
    "multidisciplinary domains": ["multidisciplinary domains", "research domains"],
    "funding": ["funding", "financial support", "funding statement"],
    "acknowledgments": ["acknowledgments", "acknowledgements"],
    "conflicts of interest": [
        "conflicts of interest", "conflict of interest",
        "competing interests", "declaration of competing interest"
    ],
    "declaration on ai usage": [
        "declaration on ai usage", "ai usage statement",
        "declaration of generative ai", "artificial intelligence"
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
# EXTRACTION HELPERS
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
    cleaned = re.sub(
        r"^(?:section\s*)?(?:[0-9]+(?:\.[0-9]+)*|[ivxlcdm]+)[\s.:–-]+\s*",
        "",
        line.strip(),
        flags=re.I,
    )
    return normalize(cleaned).lower().rstrip(":")

def extract_document_assets(uploaded_file) -> Tuple[str, List[Dict[str, Any]], List[str]]:
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    full_text = ""
    visual_assets = []

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                parts.append(txt)
        for t_idx, table in enumerate(doc.tables, start=1):
            table_rows = [" | ".join(c.text.strip() for c in row.cells) for row in table.rows]
            parts.append(f"[Table {t_idx}]\n" + "\n".join(table_rows))
            visual_assets.append({
                "label": f"Table {t_idx}",
                "caption_found": True,
                "visual_image_present": True,
                "placement_location": "In-line document body",
                "status": "PASS",
                "notes": f"Table verified with {len(table.rows)} rows, {len(table.columns)} cols."
            })
        full_text = "\n".join(parts)

    elif name.endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            page_text_list = []
            for page_idx, page in enumerate(pdf.pages, start=1):
                p_text = page.extract_text() or ""
                page_text_list.append(p_text)

                images = page.images or []
                tables = page.find_tables() or []
                has_images = len(images) > 0

                fig_caps = re.findall(r"(?i)\b(Fig(?:ure)?\.?\s*\d+[a-z]?)\b\s*[:.-]?\s*([^\n]{5,80})", p_text)
                for fig_id, cap_text in fig_caps:
                    clean_id = re.sub(r"\s+", " ", fig_id).title()
                    status = "PASS" if has_images else "WARN"
                    visual_assets.append({
                        "label": clean_id,
                        "caption_found": True,
                        "visual_image_present": has_images,
                        "placement_location": f"Page {page_idx}",
                        "status": status,
                        "notes": f"Caption '{cap_text[:35]}...' | Embedded image stream present: {has_images}"
                    })

                tbl_caps = re.findall(r"(?i)\b(Table\s*\d+[a-z]?)\b\s*[:.-]?\s*([^\n]{5,80})", p_text)
                for tbl_id, cap_text in tbl_caps:
                    clean_id = re.sub(r"\s+", " ", tbl_id).title()
                    has_table = len(tables) > 0 or "|" in p_text
                    status = "PASS" if has_table else "WARN"
                    visual_assets.append({
                        "label": clean_id,
                        "caption_found": True,
                        "visual_image_present": has_table,
                        "placement_location": f"Page {page_idx}",
                        "status": status,
                        "notes": f"Table grid detected: {has_table}"
                    })
            full_text = "\n".join(page_text_list)
    else:
        raise ValueError("Unsupported format. Use .docx or .pdf.")

    return full_text, visual_assets, full_text.splitlines()

# ============================================================
# DETERMINISTIC AUDITS
# ============================================================

def find_robust_sections(lines: List[str]) -> Dict[str, Dict[str, Any]]:
    found = {}
    for idx, line in enumerate(lines):
        raw = line.strip()
        cleaned = strip_numerical_prefix(raw)
        for canonical, aliases in SECTION_ALIASES.items():
            if cleaned in aliases and canonical not in found:
                next_text = ""
                for n_line in lines[idx + 1: idx + 10]:
                    if n_line.strip():
                        next_text = n_line.strip()
                        break
                found[canonical] = {
                    "line_index": idx,
                    "raw_heading": raw,
                    "first_line_quote": next_text[:120],
                    "has_body": bool(next_text),
                }
    return found

def extract_abstract(text: str) -> str:
    lines = text.splitlines()
    start = None
    for idx, line in enumerate(lines):
        if strip_numerical_prefix(line) == "abstract":
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
    return normalize(val)

def extract_keywords(text: str) -> List[str]:
    m = re.search(r"(?is)\bkeywords?\s*:\s*(.*?)(?=\n\s*(?:[0-9]+\.?\s*)?introduction\b|\n\s*\n|$)", text)
    if not m:
        return []
    raw = " ".join([l.strip() for l in m.group(1).splitlines() if l.strip()])
    return [normalize(k) for k in re.split(r"[;,]", raw) if normalize(k)]

def run_pragmatic_rule_checks(text: str, visual_assets: List[Dict[str, Any]]) -> Dict[str, Any]:
    lines = text.splitlines()
    sections = find_robust_sections(lines)
    abstract = extract_abstract(text)
    keywords = extract_keywords(text)

    structural_checks = []
    for canonical in MRJ_RULES["required_sections"]:
        sec_info = sections.get(canonical.lower())
        if sec_info and sec_info["has_body"]:
            status = "PASS"
        elif sec_info and not sec_info["has_body"]:
            status = "WARN"
        else:
            status = "FAIL"

        structural_checks.append({
            "section_name": canonical,
            "detected": bool(sec_info),
            "detected_heading_text": sec_info["raw_heading"] if sec_info else "None",
            "first_line_quote": sec_info["first_line_quote"] if sec_info else "None",
            "status": status,
        })

    issues = []
    abs_words = word_count(abstract)
    if abs_words == 0:
        issues.append({"issue_type": "Abstract", "evidence_quote": "No abstract detected.", "severity": "Major"})
    elif abs_words > MRJ_RULES["abstract_soft_overflow_limit"]:
        issues.append({"issue_type": "Abstract Length", "evidence_quote": f"Abstract contains {abs_words} words (>200 target).", "severity": "Minor"})

    return {
        "structural_section_checks": structural_checks,
        "visual_asset_checks": visual_assets,
        "citation_and_formatting_issues": issues,
        "keywords": keywords,
        "abstract": abstract,
        "sections": sections,
    }

# ============================================================
# PARALLEL OPENALEX REVIEWER HARVESTING
# ============================================================

def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    if mailto:
        params = dict(params)
        params["mailto"] = mailto
    res = requests.get(url, params=params, timeout=10)
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

def harvest_openalex_reviewers(keywords: List[str], research_area: str, mailto: str = "") -> Dict[str, List[Dict[str, Any]]]:
    clean_terms = [normalize(k) for k in (keywords[:3] + [research_area]) if normalize(k)]
    query = " ".join(clean_terms[:4]).strip()
    out = {"India": [], "Northeast India": [], "Assam": []}
    if not query:
        return out

    try:
        data = openalex_get(OPENALEX_URL, {"search": query, "per-page": 30, "sort": "publication_year:desc"}, mailto)
    except Exception as exc:
        return {k: [{"error": f"OpenAlex query error: {exc}"}] for k in out}

    candidates = {}
    for work in data.get("results", []):
        year = work.get("publication_year") or 0
        title = work.get("display_name") or ""
        for authorship in work.get("authorships", []):
            author = authorship.get("author") or {}
            aid = author.get("id")
            aname = author.get("display_name")
            insts = authorship.get("institutions") or []
            if not aid or not aname or not insts:
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
            candidates[aid]["publications"].append({"title": title, "year": year})
            candidates[aid]["score"] += max(0, year - 2018)

    top_candidates = sorted(candidates.values(), key=lambda x: x["score"], reverse=True)[:15]

    # Parallelize metadata fetch with ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_cand = {executor.submit(get_openalex_author, c["id"], mailto): c for c in top_candidates}
        for future in as_completed(future_to_cand):
            c = future_to_cand[future]
            try:
                auth_info = future.result()
                lk = (auth_info.get("last_known_institutions") or [{}])[0]
                geo = lk.get("geo") or {}
                c["last_institution"] = lk.get("display_name") or c["institution"]
                c["last_country"] = lk.get("country_code") or c["country"]
                c["last_city"] = geo.get("city") or c["city"]
                c["last_region"] = geo.get("region") or c["region"]
            except Exception:
                c["last_institution"] = c["institution"]
                c["last_country"] = c["country"]
                c["last_city"] = c["city"]
                c["last_region"] = c["region"]

    for reg in ["Assam", "Northeast India", "India"]:
        matched = [
            c for c in top_candidates
            if candidate_matches_region({
                "institution": c.get("last_institution", ""),
                "country": c.get("last_country", ""),
                "city": c.get("last_city", ""),
                "region": c.get("last_region", ""),
            }, reg)
        ]
        out[reg] = matched[:5]

    return out

# ============================================================
# SAFE BLIND COPY GENERATOR
# ============================================================

def blind_copy_docx(original_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(original_bytes))
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""

    paras = list(doc.paragraphs)
    abstract_idx = None
    for idx, p in enumerate(paras):
        if strip_numerical_prefix(p.text) == "abstract":
            abstract_idx = idx
            break

    if abstract_idx is not None:
        nonempty = [i for i, p in enumerate(paras[:abstract_idx]) if p.text.strip()]
        if len(nonempty) > 1:
            for i in nonempty[1:]:
                p_elem = paras[i]._element
                if p_elem.getparent() is not None:
                    p_elem.getparent().remove(p_elem)

    # Remove sensitive sections including numeric prefixes
    removable = {"funding", "acknowledgments", "acknowledgements", "conflicts of interest", "competing interests"}
    deleting = False
    for p in list(doc.paragraphs):
        head = strip_numerical_prefix(p.text)
        if any(head.startswith(r) for r in removable):
            deleting = True
        elif any(head == s.lower() for s in MRJ_RULES["required_sections"]):
            deleting = False

        if deleting:
            p_elem = p._element
            if p_elem.getparent() is not None:
                p_elem.getparent().remove(p_elem)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
